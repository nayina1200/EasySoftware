from __future__ import annotations

import contextlib
import importlib.util
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
# 被测模块按文件路径加载，需要把 scripts 加入 sys.path，
# 否则被测模块内部的同级 import（如 archive_types）无法解析。
sys.path.insert(0, str(ROOT / "scripts"))
SPEC = importlib.util.spec_from_file_location("export_pdf_scope_tests", ROOT / "scripts" / "unified_workflow.py")
WORKFLOW = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(WORKFLOW)


def material_dir(parent: Path, name: str) -> Path:
    """建一个含单个 Word 文档的材料目录，返回材料目录本身。"""
    folder = parent / name
    (folder / "项目").mkdir(parents=True, exist_ok=True)
    (folder / "项目" / "代码.docx").write_bytes(b"docx")
    return folder


class ExportDirectoryResolutionTests(unittest.TestCase):
    def test_selected_folder_is_used_regardless_of_its_name(self):
        # 原拦截：目录名不以“修改版”结尾时，程序改找同级/子级的“修改版”，
        # 用户实际选择的目录里没有任何 PDF 被更新。
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            selected = material_dir(base, "解压版")
            self.assertEqual(WORKFLOW.resolve_export_directory(base / "修改版", selected), selected)

    def test_selected_batch_prefers_its_modified_subdirectory(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            batch = base / "批次"
            material_dir(batch, "解压版")
            modified = material_dir(batch, "修改版")
            self.assertEqual(WORKFLOW.resolve_export_directory(modified, batch), modified)

    def test_selected_folder_wins_over_sibling_modified(self):
        # 兄弟目录里恰好存在“修改版”时，也不能导出到用户没有选择的那个目录。
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            selected = material_dir(base, "解压版")
            modified = material_dir(base, "修改版")
            self.assertEqual(WORKFLOW.resolve_export_directory(modified, selected), selected)

    def test_selected_folder_without_docx_falls_back_to_modified(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            selected = base / "空材料"
            selected.mkdir()
            (selected / "说明.txt").write_bytes(b"txt")
            modified = material_dir(base, "修改版")
            self.assertEqual(WORKFLOW.resolve_export_directory(modified, selected), modified)

    def test_archive_input_falls_back_to_modified(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            archive = base / "材料.zip"
            archive.write_bytes(b"zip")
            modified = material_dir(base, "材料_修改版")
            self.assertEqual(WORKFLOW.resolve_export_directory(modified, archive), modified)

    def test_returns_selected_folder_when_nothing_has_docx(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            selected = base / "材料"
            selected.mkdir()
            (selected / "说明.txt").write_bytes(b"txt")
            self.assertEqual(WORKFLOW.resolve_export_directory(base / "修改版", selected), selected)

    def test_backup_and_temporary_docx_do_not_count_as_exportable(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory) / "材料"
            folder.mkdir()
            (folder / "~$代码.docx").write_bytes(b"docx")
            (folder / "backup").mkdir()
            (folder / "backup" / "旧版.docx").write_bytes(b"docx")
            self.assertEqual(WORKFLOW.exportable_docx_files(folder), [])
            (folder / "项目.docx").write_bytes(b"docx")
            self.assertEqual(len(WORKFLOW.exportable_docx_files(folder)), 1)


class ExportModifiedPdfsScopeTests(unittest.TestCase):
    def test_reports_clear_error_when_directory_has_no_docx(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory) / "材料"
            folder.mkdir()
            (folder / "说明.txt").write_bytes(b"txt")
            result = WORKFLOW.export_modified_pdfs(folder, Path(directory) / "work")
            self.assertEqual(result["status"], "FAILED")
            self.assertIn("没有 Word 文档", result["error"])

    def test_missing_directory_message_mentions_docx_fallback(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            result = WORKFLOW.export_modified_pdfs(base / "missing", base / "work")
            self.assertEqual(result["status"], "FAILED")
            self.assertIn("未找到修改版目录", result["error"])
            self.assertIn("Word 文档", result["error"])


class ExportCliScopeTests(unittest.TestCase):
    def invoke(self, arguments: list[str], captured: list[Path]) -> int:
        def record_base(base: Path, work: Path, progress=None):
            captured.append(base.resolve())
            return {"status": "EXPORT_OK", "exported": 1, "skipped": 0, "failed": 0}

        argv = ["unified_workflow.py"] + arguments
        with mock.patch.object(sys, "argv", argv), \
                mock.patch.object(WORKFLOW, "export_modified_pdfs", side_effect=record_base), \
                mock.patch.object(WORKFLOW, "find_python", return_value=Path(sys.executable)), \
                contextlib.redirect_stdout(io.StringIO()):
            return WORKFLOW.main()

    def test_export_entry_is_not_gated_by_folder_name(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            selected = material_dir(base, "解压版")
            work = base / "work"
            work.mkdir()
            captured: list[Path] = []
            code = self.invoke([str(selected), "--work-dir", str(work), "--export-modified-pdfs"], captured)
            self.assertEqual(code, 0)
            self.assertEqual(captured, [selected.resolve()])

    def test_export_entry_from_extracted_flag_uses_selected_folder(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            selected = material_dir(base, "客户材料")
            work = base / "work"
            work.mkdir()
            captured: list[Path] = []
            code = self.invoke(
                [str(selected), "--work-dir", str(work), "--from-extracted", "--export-modified-pdfs"], captured
            )
            self.assertEqual(code, 0)
            self.assertEqual(captured, [selected.resolve()])

    def test_export_entry_still_targets_modified_subdirectory(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            batch = base / "批次"
            material_dir(batch, "解压版")
            modified = material_dir(batch, "修改版")
            work = base / "work"
            work.mkdir()
            captured: list[Path] = []
            code = self.invoke([str(batch), "--work-dir", str(work), "--export-modified-pdfs"], captured)
            self.assertEqual(code, 0)
            self.assertEqual(captured, [modified.resolve()])

    def test_export_entry_selected_folder_wins_over_sibling_modified(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            selected = material_dir(base, "解压版")
            material_dir(base, "修改版")
            work = base / "work"
            work.mkdir()
            captured: list[Path] = []
            code = self.invoke([str(selected), "--work-dir", str(work), "--export-modified-pdfs"], captured)
            self.assertEqual(code, 0)
            self.assertEqual(captured, [selected.resolve()])


if __name__ == "__main__":
    unittest.main()

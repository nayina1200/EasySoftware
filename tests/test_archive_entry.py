from __future__ import annotations

import importlib.util
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import archive_types  # noqa: E402
import process_batch  # noqa: E402


def load(alias: str, filename: str):
    spec = importlib.util.spec_from_file_location(alias, ROOT / "scripts" / filename)
    module = importlib.util.module_from_spec(spec)
    sys.modules[alias] = module
    assert spec.loader
    spec.loader.exec_module(module)
    return module


CLEANER = load("archive_entry_cleaner", "code_comment_cleaner.py")
WORKFLOW = load("archive_entry_workflow", "unified_workflow.py")


class ArchiveTypeTests(unittest.TestCase):
    def test_common_archive_types_are_accepted_as_input(self):
        for suffix in (".7z", ".rar", ".rarx", ".zip", ".jar", ".iso", ".cab",
                       ".tar", ".gz", ".tgz", ".bz2", ".xz", ".lz4", ".zst",
                       ".tar.gz", ".tar.xz", ".tar.bz2", ".ace", ".msi", ".deb"):
            self.assertTrue(archive_types.is_archive_file(f"材料{suffix}"), suffix)

    def test_delivery_formats_are_only_zip_and_rar(self):
        self.assertEqual(archive_types.DELIVERY_SUFFIXES, {".zip", ".rar"})
        self.assertEqual(archive_types.delivery_format("材料.zip"), "zip")
        self.assertEqual(archive_types.delivery_format("材料.RAR"), "rar")
        self.assertEqual(archive_types.delivery_format("材料.7z"), "")
        self.assertEqual(archive_types.delivery_format("材料.tar.gz"), "")

    def test_plain_files_are_not_archives(self):
        for suffix in (".txt", ".docx", ".pdf", ".py", ".jpg", ".mp4"):
            self.assertFalse(archive_types.is_archive_file(f"文件{suffix}"), suffix)

    def test_compound_suffix_is_matched_as_a_whole(self):
        self.assertEqual(Path("材料.tar.gz").suffix, ".gz")
        self.assertEqual(archive_types.archive_suffix("材料.tar.gz"), ".tar.gz")
        self.assertEqual(archive_types.archive_suffix("材料.TAR.XZ"), ".tar.xz")
        self.assertEqual(archive_types.archive_suffix("材料.zip"), ".zip")

    def test_suffix_matching_is_case_insensitive(self):
        self.assertTrue(archive_types.is_archive_file("材料.RAR"))
        self.assertEqual(archive_types.delivery_format("材料.ZIP"), "zip")

    def test_delivery_set_is_a_subset_of_the_input_set(self):
        self.assertTrue(archive_types.DELIVERY_SUFFIXES <= archive_types.ARCHIVE_SUFFIXES)

    def test_leftover_archive_is_detected_by_its_same_named_sibling_dir(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            archive = base / "材料.rar"
            self.assertFalse(archive_types.is_leftover_archive(archive))
            (base / "材料").mkdir()
            self.assertTrue(archive_types.is_leftover_archive(archive))
            self.assertEqual(archive_types.extracted_sibling(archive), base / "材料")


class LauncherArchiveListTests(unittest.TestCase):
    """启动脚本与 Python 的后缀清单必须一致，这正是原 bug 的根因。"""

    def test_launcher_extension_list_matches_python_module(self):
        launcher = (ROOT / "启动EasySoftware.ps1").read_text(encoding="utf-8-sig")
        block = re.search(r"\$script:ArchiveExtensions\s*=\s*@\((.*?)\n\)", launcher, re.DOTALL)
        self.assertIsNotNone(block, "启动脚本缺少 $script:ArchiveExtensions 清单")
        extensions = set(re.findall(r"'([^']+)'", block.group(1)))
        self.assertEqual(extensions, archive_types.ARCHIVE_SUFFIXES)


class DiscoverArchivesTests(unittest.TestCase):
    def _code_material(self, folder: Path) -> None:
        (folder / "项目代码.docx").write_bytes(b"")
        (folder / "项目代码.pdf").write_bytes(b"")

    def test_extracted_folder_with_leftover_archive_returns_folder_only(self):
        """用户的核心诉求：文件夹入口不再处理就地解压后残留的压缩包。"""
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            self._code_material(base)
            (base / "材料.rar").write_bytes(b"archive")
            (base / "材料").mkdir()
            self.assertEqual(CLEANER.discover_archives(base), [base])

    def test_batch_of_archives_returns_them_all(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            names = ["项目甲.rar", "项目乙.zip", "项目丙.7z", "项目丁.tar.gz"]
            for name in names:
                (base / name).write_bytes(b"archive")
            self.assertEqual(
                CLEANER.discover_archives(base),
                [base / name for name in sorted(names, key=str.casefold)],
            )

    def test_extracted_folder_without_archives_returns_folder(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            self._code_material(base)
            (base / "项目说明.txt").write_text("内容", encoding="utf-8")
            self.assertEqual(CLEANER.discover_archives(base), [base])

    def test_processed_outputs_are_excluded(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "材料_已处理_20260101_000000").mkdir()
            (base / "材料_已处理_20260101_000000" / "材料.rar").write_bytes(b"archive")
            self.assertEqual(CLEANER.discover_archives(base), [])

    def test_archive_file_input_accepts_any_archive_type(self):
        with tempfile.TemporaryDirectory() as directory:
            for suffix in (".rar", ".zip", ".7z", ".tar.gz"):
                archive = Path(directory) / f"材料{suffix}"
                archive.write_bytes(b"archive")
                self.assertEqual(CLEANER.discover_archives(archive), [archive])

    def test_non_archive_file_input_raises_a_clear_error(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "材料.txt"
            source.write_text("不是压缩包", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "压缩包"):
                CLEANER.discover_archives(source)

    def test_run_raises_clear_error_when_folder_has_nothing_processable(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "图片.jpg").write_bytes(b"jpg-stub")
            with self.assertRaisesRegex(ValueError, "没有压缩包"):
                CLEANER.run(base, base / "work", ROOT, apply=False)


class DetectFormatTests(unittest.TestCase):
    def test_non_delivery_archive_does_not_drive_delivery_format(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "材料.7z").write_bytes(b"archive")
            self.assertIsNone(WORKFLOW.detect_format(base))

    def test_pure_rar_batch_returns_rar(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "材料.rar").write_bytes(b"archive")
            self.assertEqual(WORKFLOW.detect_format(base), "rar")

    def test_mixed_delivery_formats_return_none(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "项目甲.zip").write_bytes(b"zip")
            (base / "项目乙.rar").write_bytes(b"rar")
            self.assertIsNone(WORKFLOW.detect_format(base))


class ProcessBatchFolderEntryTests(unittest.TestCase):
    """文件夹入口按已解压材料处理，就地解压后残留的压缩包不再解压。"""

    def test_extracted_folder_skips_leftover_archive_and_records_it(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "材料").mkdir()
            (base / "材料" / "项目代码.docx").write_bytes(b"")
            (base / "项目代码.docx").write_bytes(b"")
            (base / "项目代码.pdf").write_bytes(b"")
            (base / "项目说明.txt").write_text("软件名称：项目", encoding="utf-8")
            (base / "材料.rar").write_bytes(b"archive")
            work = base / "work"
            work.mkdir()

            result = process_batch.ensure_extracted(
                base, base / "解压版", work, process_batch.HashCache(work / "hash.json"), 1
            )

            state = json.loads((work / "extraction.json").read_text(encoding="utf-8"))
            self.assertEqual(state["archives"], [])
            self.assertEqual(state["ignored_archives"], ["材料.rar"])
            self.assertEqual(len(result), 1)
            self.assertEqual(result[0]["mode"], "PREEXTRACTED_COPY")
            # 残留压缩包不进入只读审计副本
            self.assertFalse(list((base / "解压版").rglob("*.rar")))

    def test_folder_without_archive_or_material_raises_clear_error(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            # 既无压缩包，也没有 txt/pdf/docx 软著材料
            (base / "图片.jpg").write_bytes(b"jpg-stub")
            work = base / "work"
            work.mkdir()
            with self.assertRaisesRegex(RuntimeError, "没有压缩包"):
                process_batch.ensure_extracted(
                    base, base / "解压版", work, process_batch.HashCache(work / "hash.json"), 1
                )


if __name__ == "__main__":
    unittest.main()

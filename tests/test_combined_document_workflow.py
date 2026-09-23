from __future__ import annotations

import contextlib
import io
import importlib.util
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
# 被测模块按文件路径加载，需要把 scripts 加入 sys.path，
# 否则被测模块内部的同级 import（如 archive_types）无法解析。
sys.path.insert(0, str(ROOT / "scripts"))
SPEC = importlib.util.spec_from_file_location(
    "combined_document_workflow", ROOT / "scripts" / "combined_document_workflow.py"
)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
assert SPEC.loader
SPEC.loader.exec_module(MODULE)


def make_docx(path: Path, lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    document = etree.Element(f"{{{ns}}}document", nsmap={"w": ns})
    body = etree.SubElement(document, f"{{{ns}}}body")
    for line in lines:
        paragraph = etree.SubElement(body, f"{{{ns}}}p")
        run = etree.SubElement(paragraph, f"{{{ns}}}r")
        etree.SubElement(run, f"{{{ns}}}t").text = line
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
        archive.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8"))


class Reporter:
    def __init__(self) -> None:
        self.total = 0
        self.events: list[dict] = []

    def set_total(self, total: int) -> None:
        self.total = total

    def update(self, phase: str, current: int = 0, name: str = "", **details) -> None:
        self.events.append({"phase": phase, "current": current, "name": name, **details})


class CombinedDocumentWorkflowTests(unittest.TestCase):
    def test_report_shows_actual_before_after_year_images(self):
        from PIL import Image, ImageDraw

        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            source_root, output_root, work = (base / name for name in ("source", "output", "work"))
            for root in (source_root, output_root):
                root.mkdir()
            original = Image.new("RGB", (320, 100), "white")
            processed = original.copy()
            ImageDraw.Draw(original).text((30, 30), "2023", fill="black")
            ImageDraw.Draw(processed).text((30, 30), "2025", fill="black")
            for root, picture in ((source_root, original), (output_root, processed)):
                stream = io.BytesIO()
                picture.save(stream, format="PNG")
                with zipfile.ZipFile(root / "项目说明.docx", "w") as archive:
                    archive.writestr("word/media/image1.png", stream.getvalue())
                    other = io.BytesIO()
                    Image.new("RGB", (320, 100), "red" if root == source_root else "blue").save(other, format="PNG")
                    archive.writestr("word/media/image2.png", other.getvalue())
            report = {
                "material_root": str(source_root), "output": str(output_root), "applied": True,
                "status": "DOCUMENT_PROCESSING_OK", "failures": 0,
                "stages": {
                    "manual": {"status": "APPLIED", "failures": 0, "projects": [{
                        "relative_path": "项目说明.docx", "status": "APPLIED",
                        "samples": [{"image": "word/media/image1.png", "status": "unknown", "visual_review": {
                            "status": "irrelevant", "reason": "画面展示其他领域的设备"}}],
                        "date_rewrite": {"status": "CHANGED", "changed": True, "changed_images": 1,
                                         "changes": [{"media_path": "word/media/image1.png", "ocr_text": "2023-10-25",
                                                      "source_year": "2023", "target_year": "2025", "changed_image": True}]},
                    }]},
                    "code": {"status": "APPLIED", "failures": 0},
                    "txt": {"status": "PLACEHOLDER_NOT_IMPLEMENTED", "failures": 0},
                },
            }
            path = work / "result.json"
            MODULE._write_report(path, report)
            content = path.with_suffix(".html").read_text(encoding="utf-8")
            self.assertIn("图片变化前后对比（待核验）", content)
            self.assertIn("不能仅据此认定年份改写正确", content)
            self.assertEqual(content.count("待核验：这张图片"), 1)
            self.assertIn("原稿", content)
            self.assertIn("处理后", content)
            self.assertIn("修改前局部", content)
            self.assertIn("视觉复核仅提供人工审核证据", content)
            self.assertIn("画面展示其他领域的设备", content)
            assets = path.with_name("result_assets")
            self.assertEqual(len(list(assets.glob("*-before.png"))), 1)
            with Image.open(assets / "001-before.png") as before, Image.open(assets / "001-after.png") as after:
                self.assertNotEqual(before.tobytes(), after.tobytes())
            report["applied"] = False
            MODULE._write_report(path, report)
            self.assertIn("预览阶段尚未修改图片", path.with_suffix(".html").read_text(encoding="utf-8"))

    def make_materials(self, root: Path) -> Path:
        material = root / "材料"
        material.mkdir()
        make_docx(material / "项目说明.docx", ["截至2023年为最新版本"])
        make_docx(material / "项目代码.docx", ["// 中文注释", "run();"])
        (material / "项目代码.pdf").write_bytes(b"old-pdf")
        (material / "项目.txt").write_text("软件名称：项目", encoding="utf-8")
        return material

    @staticmethod
    def manual_options() -> dict:
        return {
            "ocr": lambda *_: {"text": ""},
            "relevance_classifier": lambda *_: {},
            "year_classifier": lambda *_: {"decision": "KEEP", "confidence": 1},
            "date_rewriter": lambda *_: {"status": "UNCHANGED", "changed": False},
        }

    def test_preview_uses_both_preflights_without_creating_processed_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            material = self.make_materials(base)
            work = base / "work"
            reporter = Reporter()

            with mock.patch.object(
                MODULE.code_comment_cleaner,
                "inspect_docx",
                wraps=MODULE.code_comment_cleaner.inspect_docx,
            ) as inspect_docx:
                result = MODULE.preview(
                    material,
                    work,
                    ROOT,
                    reporter,
                    manual_options=self.manual_options(),
                )

            self.assertEqual(result["status"], "NEEDS_DOCUMENT_PROCESSING_REVIEW")
            self.assertEqual(inspect_docx.call_count, 1)
            self.assertEqual(list(base.glob("材料_已处理_*")), [])
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertFalse(report["applied"])
            self.assertEqual(report["stage_order"], ["manual", "code", "txt"])
            self.assertEqual(report["stages"]["manual"]["manual_count"], 1)
            self.assertIn("body_relevance", report["stages"]["manual"]["projects"][0])
            self.assertEqual(report["stages"]["code"]["documents"][0]["deletions"], 1)
            self.assertEqual(report["stages"]["txt"]["status"], "PLACEHOLDER_NOT_IMPLEMENTED")
            self.assertTrue(Path(result["report"]).is_file())
            self.assertIn("标题与正文", Path(result["report"]).read_text(encoding="utf-8"))
            self.assertEqual(reporter.total, 3)
            self.assertEqual([item["current"] for item in reporter.events], sorted(item["current"] for item in reporter.events))
            self.assertEqual(reporter.events[-1]["current"], reporter.total)

    def test_execute_copies_once_then_runs_manual_code_and_txt_on_same_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            material = self.make_materials(base)
            work = base / "work"
            events: list[tuple[str, Path]] = []
            reporter = Reporter()

            original_manual = MODULE.manual_content_stage.run_manual_content_stage
            original_clean = MODULE.code_comment_cleaner.clean_docx
            original_txt = MODULE._run_txt_placeholder

            def tracked_manual(input_root, *args, **kwargs):
                events.append(("manual", Path(input_root)))
                return original_manual(input_root, *args, **kwargs)

            def tracked_clean(path):
                events.append(("code", Path(path).parent))
                return original_clean(path)

            def tracked_txt(root, **kwargs):
                events.append(("txt", Path(root)))
                return original_txt(root, **kwargs)

            def fake_cover_plan(root, path):
                plan = {"packages": [{"name": "项目", "folder": ".", "docx": "项目说明.docx", "pdf": "项目说明.pdf", "actions": ["EXPORT_PDF"]}], "audit_errors": []}
                path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
                return plan

            def tracked_word(root, *_args, **_kwargs):
                events.append(("cover", Path(root)))
                return {"packages": [{"name": "项目", "status": "OK", "failures": []}]}

            with mock.patch.object(MODULE.shutil, "copytree", wraps=shutil_copytree()) as copytree, \
                    mock.patch.object(MODULE.manual_content_stage, "run_manual_content_stage", side_effect=tracked_manual), \
                    mock.patch.object(MODULE, "_build_manual_cover_plan", side_effect=fake_cover_plan), \
                    mock.patch.object(MODULE.process_batch, "invoke_word_pipeline", side_effect=tracked_word) as invoke_word_pipeline, \
                    mock.patch.object(MODULE.process_batch, "word_automation_lock", return_value=contextlib.nullcontext()) as word_lock, \
                    mock.patch.object(MODULE.code_comment_cleaner, "clean_docx", side_effect=tracked_clean), \
                    mock.patch.object(MODULE.code_comment_cleaner, "_export_pdf", return_value=(True, "mock export")) as export_pdf, \
                    mock.patch.object(MODULE, "_run_txt_placeholder", side_effect=tracked_txt):
                result = MODULE.execute(
                    material,
                work,
                ROOT,
                progress=reporter,
                manual_options=self.manual_options(),
                )

            output = Path(result["output"])
            self.assertEqual(copytree.call_count, 1)
            self.assertEqual([name for name, _ in events], ["manual", "cover", "code", "txt"])
            self.assertTrue(all(path == output for _, path in events))
            self.assertEqual(export_pdf.call_count, 1)
            word_lock.assert_called_once()
            self.assertIs(invoke_word_pipeline.call_args.kwargs.get("progress"), reporter)
            self.assertEqual(result["outputs"], [str(output)])
            self.assertEqual(len(list(base.glob("材料_已处理_*"))), 1)
            self.assertEqual((output / "项目.txt").read_text(encoding="utf-8"), "软件名称：项目")
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "DOCUMENT_PROCESSING_OK")
            self.assertEqual(report["stages"]["code"]["documents"][0]["pdf_status"], "REPLACED")

    def test_stage_selection_skips_unselected_readers_and_counts_only_selected_items(self):
        choices = (
            ("manual", 1, ("manual",)),
            ("code", 1, ("code",)),
            ("manual,code", 2, ("manual", "code")),
            ("manual,txt_placeholder", 2, ("manual", "txt")),
            ("code,txt_placeholder", 2, ("code", "txt")),
            ("manual,code,txt_placeholder", 3, ("manual", "code", "txt")),
        )
        for selected, total, active in choices:
            with self.subTest(selected=selected), tempfile.TemporaryDirectory() as directory:
                base = Path(directory)
                material = self.make_materials(base)
                reporter = Reporter()
                with mock.patch.object(MODULE.manual_content_stage, "run_manual_content_stage", wraps=MODULE.manual_content_stage.run_manual_content_stage) as manual, \
                        mock.patch.object(MODULE, "_build_manual_cover_plan", wraps=MODULE._build_manual_cover_plan) as cover, \
                        mock.patch.object(MODULE.code_comment_cleaner, "_code_files", wraps=MODULE.code_comment_cleaner._code_files) as code_files, \
                        mock.patch.object(MODULE, "_run_txt_placeholder", wraps=MODULE._run_txt_placeholder) as txt:
                    result = MODULE.preview(material, base / "work", ROOT, reporter,
                                            manual_options=self.manual_options(), document_stages=selected)
                report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
                self.assertEqual(reporter.total, total)
                self.assertEqual(reporter.events[-1]["current"], total)
                self.assertEqual(report["selected_stages"], [part for part in MODULE.DOCUMENT_STAGES if part in selected.split(",")])
                for key, called in (("manual", manual.called), ("code", code_files.called), ("txt", txt.called)):
                    self.assertEqual(called, key in active, key)
                    self.assertEqual(report["stages"][key]["status"] == "SKIPPED", key not in active)
                self.assertEqual(cover.called, "manual" in active)

    def test_invalid_stage_selection_rejected_before_reading_materials(self):
        with tempfile.TemporaryDirectory() as directory:
            source = self.make_materials(Path(directory))
            for selected in ("", "txt_placeholder", "unknown", "manual,unknown"):
                with self.subTest(selected=selected), mock.patch.object(MODULE, "_material_root") as material_root:
                    with self.assertRaises(ValueError):
                        MODULE.preview(source, Path(directory) / "work", ROOT, document_stages=selected)
                    material_root.assert_not_called()

    def test_code_only_execute_does_not_open_manual_or_txt_stage(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            material = self.make_materials(base)
            reporter = Reporter()
            with mock.patch.object(MODULE.manual_content_stage, "run_manual_content_stage") as manual, \
                    mock.patch.object(MODULE, "_build_manual_cover_plan") as cover, \
                    mock.patch.object(MODULE, "_run_txt_placeholder") as txt, \
                    mock.patch.object(MODULE.code_comment_cleaner, "clean_docx", return_value={
                        "docx": str(material / "项目代码.docx"), "status": "CLEANED", "deletions": 0,
                    }) as clean:
                result = MODULE.execute(material, base / "work", ROOT, reporter, document_stages="code")
            manual.assert_not_called()
            cover.assert_not_called()
            txt.assert_not_called()
            clean.assert_called_once()
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(report["stages"]["manual"]["status"], "SKIPPED")
            self.assertEqual(report["stages"]["txt"]["status"], "SKIPPED")
            self.assertEqual(reporter.total, 1)
            self.assertEqual(reporter.events[-1]["current"], 1)

    def test_rar_input_is_extracted_once_and_original_is_untouched(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            material = self.make_materials(base)
            archive = base / "交付.rar"
            archive.write_bytes(b"rar-source")

            with mock.patch.object(MODULE.code_comment_cleaner, "extract_cached", return_value=material) as extract, \
                    mock.patch.object(MODULE.code_comment_cleaner, "material_root", return_value=material), \
                    mock.patch.object(MODULE.process_batch, "expand_nested_archives", return_value=[]) as expand_nested:
                result = MODULE.preview(
                    archive,
                    base / "work",
                    ROOT,
                    manual_options=self.manual_options(),
                )

            extract.assert_called_once()
            expand_nested.assert_called_once()
            self.assertEqual(Path(expand_nested.call_args.args[0]), material)
            self.assertEqual(archive.read_bytes(), b"rar-source")
            self.assertIsNone(result["output"])

    def test_rejects_non_archive_file(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            source = base / "材料.txt"
            source.write_text("不是压缩包", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "压缩包"):
                MODULE.preview(source, base / "work", ROOT)

    def test_accepts_any_archive_type_as_input(self):
        """交付格式只有 RAR/ZIP，但输入读取接受全部压缩类型。"""
        for suffix in (".rar", ".zip", ".7z", ".tar.gz"):
            with tempfile.TemporaryDirectory() as directory:
                base = Path(directory)
                material = self.make_materials(base)
                archive = base / f"材料{suffix}"
                archive.write_bytes(b"archive")
                with mock.patch.object(MODULE.code_comment_cleaner, "extract_cached", return_value=material) as extract,                         mock.patch.object(MODULE.code_comment_cleaner, "material_root", return_value=material):
                    MODULE.preview(archive, base / "work", ROOT, manual_options=self.manual_options())
                extract.assert_called_once()


def shutil_copytree():
    import shutil

    return shutil.copytree


if __name__ == "__main__":
    unittest.main()

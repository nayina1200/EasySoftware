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
    def test_code_page_count_allows_one_page_deviation_and_rejects_two(self):
        for actual_pages, expected_failures in ((61, 0), (59, 0), (62, 1), (58, 1)):
            with self.subTest(actual_pages=actual_pages), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                docx = root / "代码.docx"
                pdf = root / "代码.pdf"
                discovered = {"docx": str(docx), "pdf": str(pdf), "status": "READY"}
                translated = {
                    "docx": str(docx), "status": "CLEANED", "deletions": 0, "findings": [],
                    "target_pages": 60, "actual_pages": actual_pages, "page_match": False,
                    "removed_translated_comments": 0,
                }
                exported = expected_failures == 0
                export_detail = (
                    f"目标 60 页，实际 {actual_pages} 页（偏差 1 页）"
                    if exported else f"页数偏差超过 1 页：原 PDF 60 页，翻译后 {actual_pages} 页；已保留原 PDF"
                )
                with mock.patch.object(MODULE, "_code_inventory", return_value=[discovered]), \
                        mock.patch.object(MODULE.code_comment_cleaner, "inspect_docx", return_value={}), \
                        mock.patch.object(MODULE.code_comment_cleaner, "pdf_page_count", return_value=60), \
                        mock.patch.object(MODULE.code_comment_cleaner, "translate_and_fit_docx", return_value=translated), \
                        mock.patch.object(MODULE.code_comment_cleaner, "_export_pdf", return_value=(exported, export_detail)) as export_pdf:
                    stage, failures = MODULE._run_code_stage(
                        root, ROOT, root / "work", apply=True, progress=Reporter(), offset=0,
                        failures=0, inspect=True, cleanup=True, export_pdf=True,
                    )

                row = stage["documents"][0]
                export_pdf.assert_called_once()
                self.assertEqual(row["pdf_status"], "REPLACED" if exported else "PRESERVED")
                self.assertEqual(failures, expected_failures)
                self.assertEqual(stage["status"], "PARTIAL" if expected_failures else "APPLIED")
                self.assertEqual(row["target_pages"], 60)
                self.assertEqual(row["actual_pages"], actual_pages)
                if expected_failures:
                    self.assertEqual(row["status"], "CLEANED_PDF_FAILED")
                    self.assertFalse(row["page_tolerance_match"])
                    self.assertIn(f"目标 60 页，实际 {actual_pages} 页", row["acceptance_error"])
                    self.assertIn("已保留原 PDF", row["pdf_detail"])
                    self.assertIn("拟合过程已移除 0 条译文注释", row["acceptance_error"])
                else:
                    deviation = actual_pages - 60
                    self.assertEqual(row["status"], "CLEANED")
                    self.assertFalse(row["page_match"])  # 精确相等结果仍如实保留
                    self.assertTrue(row["page_tolerance_match"])
                    self.assertEqual(row["page_deviation_pages"], deviation)
                    self.assertIn(f"相差 {deviation:+d} 页（允许 ±1 页）", row["page_deviation_note"])

    def test_code_exact_page_match_does_not_add_deviation_note(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            docx = root / "代码.docx"
            pdf = root / "代码.pdf"
            discovered = {"docx": str(docx), "pdf": str(pdf), "status": "READY"}
            translated = {
                "docx": str(docx), "status": "CLEANED", "deletions": 0, "findings": [],
                "target_pages": 60, "actual_pages": 60, "page_match": True,
                "removed_translated_comments": 0,
            }
            with mock.patch.object(MODULE, "_code_inventory", return_value=[discovered]), \
                    mock.patch.object(MODULE.code_comment_cleaner, "inspect_docx", return_value={}), \
                    mock.patch.object(MODULE.code_comment_cleaner, "pdf_page_count", return_value=60), \
                    mock.patch.object(MODULE.code_comment_cleaner, "translate_and_fit_docx", return_value=translated):
                stage, failures = MODULE._run_code_stage(
                    root, ROOT, root / "work", apply=True, progress=Reporter(), offset=0,
                    failures=0, inspect=True, cleanup=True, export_pdf=False,
                )

        row = stage["documents"][0]
        self.assertEqual(failures, 0)
        self.assertEqual(row["status"], "CLEANED")
        self.assertTrue(row["page_match"])
        self.assertNotIn("page_deviation_pages", row)
        self.assertNotIn("page_deviation_note", row)

    def test_code_rolls_back_pair_when_word_page_read_fails_after_docx_change_and_continues(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first_docx, first_pdf = root / "一.docx", root / "一.pdf"
            next_docx, next_pdf = root / "二.docx", root / "二.pdf"
            originals = {
                first_docx: b"original docx one", first_pdf: b"original pdf one",
                next_docx: b"original docx two", next_pdf: b"original pdf two",
            }
            for path, content in originals.items():
                path.write_bytes(content)
            discovered = [
                {"docx": str(first_docx), "pdf": str(first_pdf), "status": "READY"},
                {"docx": str(next_docx), "pdf": str(next_pdf), "status": "READY"},
            ]

            def translate(path, *_args):
                path.write_bytes(b"translated " + path.name.encode())
                if path == first_docx:
                    raise RuntimeError("Word 页数读取失败")
                return {"docx": str(path), "status": "CLEANED", "deletions": 1, "findings": []}

            with mock.patch.object(MODULE, "_code_inventory", return_value=discovered), \
                    mock.patch.object(MODULE.code_comment_cleaner, "inspect_docx", return_value={}), \
                    mock.patch.object(MODULE.code_comment_cleaner, "pdf_page_count", return_value=1), \
                    mock.patch.object(MODULE.code_comment_cleaner, "translate_and_fit_docx", side_effect=translate), \
                    mock.patch.object(MODULE.code_comment_cleaner, "_export_pdf", return_value=(True, "ok")) as export_pdf:
                stage, failures = MODULE._run_code_stage(
                    root, ROOT, root / "work", apply=True, progress=Reporter(), offset=0,
                    failures=0, inspect=True, cleanup=True, export_pdf=True,
                )

            first, second = stage["documents"]
            self.assertEqual(failures, 1)
            self.assertEqual(first["status"], "FAILED")
            self.assertIn("Word 页数读取失败", first["error"])
            self.assertEqual(first["transaction_rollback"], "DOCX/PDF 已恢复至处理前字节")
            self.assertEqual(first_docx.read_bytes(), originals[first_docx])
            self.assertEqual(first_pdf.read_bytes(), originals[first_pdf])
            self.assertEqual(second["status"], "CLEANED")
            self.assertEqual(next_docx.read_bytes(), b"translated "+next_docx.name.encode())
            export_pdf.assert_called_once()

    def test_code_rolls_back_docx_and_pdf_if_export_replaces_pdf_then_fails(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first_docx, first_pdf = root / "一.docx", root / "一.pdf"
            next_docx, next_pdf = root / "二.docx", root / "二.pdf"
            originals = {
                first_docx: b"original docx one", first_pdf: b"original pdf one",
                next_docx: b"original docx two", next_pdf: b"original pdf two",
            }
            for path, content in originals.items():
                path.write_bytes(content)
            discovered = [
                {"docx": str(first_docx), "pdf": str(first_pdf), "status": "READY"},
                {"docx": str(next_docx), "pdf": str(next_pdf), "status": "READY"},
            ]

            def translate(path, *_args):
                path.write_bytes(b"translated " + path.name.encode())
                return {
                    "docx": str(path), "status": "CLEANED", "deletions": 1, "findings": [],
                    "target_pages": 1, "actual_pages": 1, "page_match": True,
                }

            export_calls = 0

            def export(path, *_args):
                nonlocal export_calls
                export_calls += 1
                pdf_path = path.with_suffix(".pdf")
                pdf_path.write_bytes(b"replacement pdf")
                return (False, "PDF 导出后验收失败") if export_calls == 1 else (True, "ok")

            with mock.patch.object(MODULE, "_code_inventory", return_value=discovered), \
                    mock.patch.object(MODULE.code_comment_cleaner, "inspect_docx", return_value={}), \
                    mock.patch.object(MODULE.code_comment_cleaner, "pdf_page_count", return_value=1), \
                    mock.patch.object(MODULE.code_comment_cleaner, "translate_and_fit_docx", side_effect=translate), \
                    mock.patch.object(MODULE.code_comment_cleaner, "_export_pdf", side_effect=export):
                stage, failures = MODULE._run_code_stage(
                    root, ROOT, root / "work", apply=True, progress=Reporter(), offset=0,
                    failures=0, inspect=True, cleanup=True, export_pdf=True,
                )

            first, second = stage["documents"]
            self.assertEqual(failures, 1)
            self.assertEqual(first["status"], "CLEANED_PDF_FAILED")
            self.assertEqual(first["transaction_rollback"], "DOCX/PDF 已恢复至处理前字节")
            self.assertEqual(first_docx.read_bytes(), originals[first_docx])
            self.assertEqual(first_pdf.read_bytes(), originals[first_pdf])
            self.assertEqual(second["status"], "CLEANED")
            self.assertEqual(next_docx.read_bytes(), b"translated "+next_docx.name.encode())
            self.assertEqual(next_pdf.read_bytes(), b"replacement pdf")

    def test_manual_cover_plan_uses_unique_txt_name_and_always_schedules_full_audit(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "文件夹中的旧名称"
            project.mkdir()
            (project / "TXT中的软件全称.txt").write_text("名称：TXT中的软件全称", encoding="utf-8")
            docx = project / "DOCX中的旧名称说明.docx"
            make_docx(docx, ["无关正文"])
            plan = MODULE._build_manual_cover_plan(root, root / "plan.json", selected_options=("manual_cover",))

        package = plan["packages"][0]
        self.assertEqual(package["name"], "TXT中的软件全称")
        self.assertEqual(package["name_source"], "unique_txt_filename")
        self.assertEqual(package["docx_title"], "DOCX中的旧名称")
        self.assertIn("ENSURE_COVER_FIRST_PAGE", package["actions"])
        self.assertIn("AUDIT_ALL_PAGES", package["actions"])
        self.assertIn("EXPORT_PDF", package["actions"])
        self.assertEqual(plan["unchanged"], [])

    def test_manual_page_numbers_plan_checks_field_count_but_preserves_existing_alignment(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "医院项目"
            project.mkdir()
            (project / "医院软件.txt").write_text("医院软件", encoding="utf-8")
            docx = project / "医院软件说明.docx"
            make_docx(docx, ["正文"])
            cases = (
                ({"header_page_field_counts": [1], "footer_page_field_counts": [0],
                  "textframe_page_paragraphs_right_aligned": False}, False),
                ({"header_page_field_counts": [1], "footer_page_field_counts": [0],
                  "textframe_page_paragraphs_right_aligned": True}, False),
                ({"header_page_field_counts": [2], "footer_page_field_counts": [0],
                  "textframe_page_paragraphs_right_aligned": True}, True),
            )
            with mock.patch.object(MODULE.process_batch, "inspect_docx", return_value={}), \
                    mock.patch.object(MODULE.process_batch, "inspect_docx_header_only") as inspect_header:
                for header_audit, should_normalize in cases:
                    with self.subTest(header_audit=header_audit):
                        inspect_header.return_value = header_audit
                        plan = MODULE._build_manual_cover_plan(
                            root, root / "plan.json", selected_options=("manual_page_numbers",)
                        )
                        scheduled = bool(plan["packages"])
                        self.assertEqual(scheduled, should_normalize)
                        if should_normalize:
                            self.assertIn("NORMALIZE_PAGE_NUMBERS", plan["packages"][0]["actions"])

    def test_manual_pdf_requests_python_pdf_audit_without_cover_word_action(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "医院项目"
            project.mkdir()
            (project / "医院软件.txt").write_text("医院软件", encoding="utf-8")
            docx = project / "医院软件说明.docx"
            make_docx(docx, ["正文"])
            with mock.patch.object(MODULE.process_batch, "inspect_docx", return_value={}), \
                    mock.patch.object(MODULE.process_batch, "inspect_docx_header_only", return_value={
                        "header_page_field_counts": [1], "footer_page_field_counts": [0],
                    }):
                plan = MODULE._build_manual_cover_plan(
                    root, root / "plan.json", selected_options=("manual_pdf",)
                )
        actions = plan["packages"][0]["actions"]
        self.assertIn("AUDIT_PAGE_LABELS", actions)
        self.assertIn("EXPORT_PDF", actions)
        self.assertNotIn("AUDIT_ALL_PAGES", actions)

    def test_manual_cover_automatically_adds_pdf_dependency(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            source = self.make_materials(base)
            with mock.patch.object(MODULE, "_build_manual_cover_plan", return_value={"packages": [], "audit_errors": []}) as cover:
                result = MODULE.preview(
                    source, base / "work", ROOT,
                    document_stages=("manual_cover",),
                )
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
        self.assertIn("manual_pdf", report["effective_stages"])
        self.assertTrue(any(row["option"] == "manual_pdf" for row in report["auto_dependencies"]))
        cover.assert_called_once()

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
                    }], "cover_result": {"packages": [{
                        "name": "项目说明", "status": "OK", "pdf_audit": {"bookmark_errors": [
                            {"page": 7, "type": "未定义书签", "snippet": "目录 3.1 系统登录 错误!未定义书签。"}
                        ]}, "broken_bookmark_fields_unlinked": [
                            {"type": "PAGEREF", "target": "_Toc32659", "cached_result": "3"}
                        ]
                    }]}},
                    "code": {"status": "PARTIAL", "failures": 1, "documents": [{
                        "docx": "代码.docx", "status": "CLEANED_PAGE_MISMATCH",
                        "deletions": 0, "pdf_status": "REPLACED",
                        "acceptance_error": "页数仍不匹配：目标 60 页，实际 61 页；拟合过程已移除 36 条译文注释",
                    }]},
                    "txt": {"status": "PLACEHOLDER_NOT_IMPLEMENTED", "failures": 0},
                },
            }
            path = work / "result.json"
            MODULE._write_report(path, report)
            content = path.with_suffix(".html").read_text(encoding="utf-8")
            saved_report = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(
                saved_report["stages"]["manual"]["cover_result"]["packages"][0]["broken_bookmark_fields_unlinked"][0]["cached_result"],
                "3",
            )
            self.assertIn("缺失书签引用已保留缓存结果：PAGEREF _Toc32659 → 3", content)
            self.assertIn("书签错误文本：第 7 页（未定义书签）：目录 3.1 系统登录 错误!未定义书签。", content)
            self.assertIn("图片变化前后对比（待核验）", content)
            self.assertIn("不能仅据此认定年份改写正确", content)
            self.assertEqual(content.count("待核验：这张图片"), 1)
            self.assertIn("原稿", content)
            self.assertIn("处理后", content)
            self.assertIn("修改前局部", content)
            self.assertIn("视觉复核仅提供人工审核证据", content)
            self.assertIn("画面展示其他领域的设备", content)
            self.assertIn("页数仍不匹配：目标 60 页，实际 61 页；拟合过程已移除 36 条译文注释", content)
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
                    document_stages=("match_manual_ocr", "manual_cover", "code_cleanup", "code_pdf", "txt_repair"),
                    skip_topic_mismatch=False,
                )

            self.assertEqual(result["status"], "NEEDS_DOCUMENT_PROCESSING_REVIEW")
            self.assertEqual(inspect_docx.call_count, 2)
            self.assertEqual(list(base.glob("材料_已处理_*")), [])
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertFalse(report["applied"])
            self.assertIn("TXT 检查与修复", report["stage_order"])
            self.assertEqual(report["stages"]["manual"]["manual_count"], 1)
            self.assertIn("body_relevance", report["stages"]["manual"]["projects"][0])
            self.assertEqual(report["stages"]["code"]["documents"][0]["deletions"], 1)
            self.assertEqual(report["stages"]["txt"]["status"], "READY_FOR_REVIEW")
            self.assertTrue(Path(result["report"]).is_file())
            self.assertIn("标题与正文", Path(result["report"]).read_text(encoding="utf-8"))
            self.assertEqual(reporter.total, 3)
            self.assertEqual([item["current"] for item in reporter.events], sorted(item["current"] for item in reporter.events))
            self.assertEqual(reporter.events[-1]["current"], reporter.total)

    def test_next_workflow_start_recovers_interrupted_processed_sibling(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            source = base / "材料"
            source.mkdir()
            output = base / "材料_已处理_20260925"
            project = output / "项目"
            backup = output / MODULE.process_batch.COVER_BACKUP_DIRECTORY / "txn" / "项目"
            project.mkdir(parents=True)
            backup.mkdir(parents=True)
            docx = project / "说明.docx"
            pdf = project / "说明.pdf"
            docx.write_bytes(b"partially written docx")
            pdf.write_bytes(b"original pdf")
            (backup / docx.name).write_bytes(b"original docx")
            (backup / pdf.name).write_bytes(b"original pdf")
            MODULE.process_batch.atomic_json(output / MODULE.process_batch.COVER_TRANSACTION_FILENAME, {
                "schema_version": 1, "transaction_id": "txn", "state": "PENDING",
                "documents": [{
                    "name": "项目", "folder": "项目", "docx": "说明.docx", "pdf": "说明.pdf",
                    "backup_dir": str(backup), "had_pdf": True,
                    "original_hashes": {
                        "docx": MODULE.process_batch._sha256_path(backup / docx.name),
                        "pdf": MODULE.process_batch._sha256_path(backup / pdf.name),
                    },
                    "state": "PENDING",
                }],
            })

            result = MODULE.preview(
                source, base / "work", ROOT, document_stages=(), skip_topic_mismatch=False,
            )

            self.assertEqual(docx.read_bytes(), b"original docx")
            self.assertEqual(pdf.read_bytes(), b"original pdf")
            self.assertFalse((output / MODULE.process_batch.COVER_TRANSACTION_FILENAME).exists())
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(report["recovered_cover_transactions"][0]["actions"][0]["action"], "恢复修复前 DOCX/PDF 快照")

    def test_execute_copies_once_then_runs_manual_code_and_txt_on_same_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            material = self.make_materials(base)
            work = base / "work"
            events: list[tuple[str, Path]] = []
            reporter = Reporter()

            original_manual = MODULE.manual_content_stage.run_manual_content_stage
            original_txt = MODULE._run_txt_stage

            def tracked_manual(input_root, *args, **kwargs):
                events.append(("manual", Path(input_root)))
                return original_manual(input_root, *args, **kwargs)

            def tracked_translate(path, _app_root, _work, target_pages):
                events.append(("code", Path(path).parent))
                return {
                    "docx": str(path), "status": "CLEANED", "deletions": 0,
                    "findings": [], "target_pages": target_pages,
                }

            def tracked_txt(root, **kwargs):
                events.append(("txt", Path(root)))
                return original_txt(root, **kwargs)

            def fake_cover_plan(root, path, *_args):
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
                    mock.patch.object(MODULE.code_comment_cleaner, "pdf_page_count", return_value=1), \
                    mock.patch.object(MODULE.code_comment_cleaner, "translate_and_fit_docx", side_effect=tracked_translate), \
                    mock.patch.object(MODULE.code_comment_cleaner, "_export_pdf", return_value=(True, "mock export")) as export_pdf, \
                    mock.patch.object(MODULE, "_run_txt_stage", side_effect=tracked_txt):
                result = MODULE.execute(
                    material,
                work,
                ROOT,
                progress=reporter,
                manual_options=self.manual_options(),
                document_stages=("match_manual_ocr", "manual_cover", "code_cleanup", "code_pdf", "txt_repair"),
                skip_topic_mismatch=False,
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
            ("code", 2, ("code",)),
            ("manual,code", 2, ("manual", "code")),
            ("manual,txt", 2, ("manual", "txt")),
            ("code,txt", 3, ("code", "txt")),
            ("manual,code,txt", 3, ("manual", "code", "txt")),
        )
        for selected, total, active in choices:
            with self.subTest(selected=selected), tempfile.TemporaryDirectory() as directory:
                base = Path(directory)
                material = self.make_materials(base)
                reporter = Reporter()
                with mock.patch.object(MODULE.manual_content_stage, "run_manual_content_stage", wraps=MODULE.manual_content_stage.run_manual_content_stage) as manual, \
                        mock.patch.object(MODULE, "_build_manual_cover_plan", wraps=MODULE._build_manual_cover_plan) as cover, \
                        mock.patch.object(MODULE.code_comment_cleaner, "_code_files", wraps=MODULE.code_comment_cleaner._code_files) as code_files, \
                        mock.patch.object(MODULE, "_run_txt_stage", wraps=MODULE._run_txt_stage) as txt:
                    result = MODULE.preview(material, base / "work", ROOT, reporter,
                                            manual_options=self.manual_options(), document_stages=selected)
                report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
                self.assertEqual(reporter.total, total)
                self.assertEqual(reporter.events[-1]["current"], total)
                self.assertEqual(report["selected_stages"], list(MODULE._selected_stages(selected)))
                for key, called in (("manual", manual.called), ("code", code_files.called), ("txt", txt.called)):
                    self.assertEqual(called, key in active, key)
                    self.assertEqual(report["stages"][key]["status"] == "SKIPPED", key not in active)
                self.assertEqual(cover.called, "manual" in active)

    def test_invalid_stage_selection_rejected_before_reading_materials(self):
        with tempfile.TemporaryDirectory() as directory:
            source = self.make_materials(Path(directory))
            for selected in ("unknown", "manual,unknown"):
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
                    mock.patch.object(MODULE, "_run_txt_stage") as txt, \
                    mock.patch.object(MODULE.code_comment_cleaner, "pdf_page_count", return_value=1), \
                    mock.patch.object(MODULE.code_comment_cleaner, "translate_and_fit_docx", return_value={
                        "docx": str(material / "项目代码.docx"), "status": "CLEANED", "deletions": 0,
                    }) as translate:
                result = MODULE.execute(material, base / "work", ROOT, reporter, document_stages="code")
            manual.assert_not_called()
            cover.assert_not_called()
            txt.assert_not_called()
            translate.assert_called_once()
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(report["stages"]["manual"]["status"], "SKIPPED")
            self.assertEqual(report["stages"]["txt"]["status"], "SKIPPED")
            self.assertEqual(reporter.total, 2)
            self.assertEqual(reporter.events[-1]["current"], 2)

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

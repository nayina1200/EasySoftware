from __future__ import annotations

import contextlib
import importlib.util
import base64
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from lxml import etree
from pypdf import PdfWriter


ROOT = Path(__file__).resolve().parents[1]
# 被测模块按文件路径加载，需要把 scripts 加入 sys.path，
# 否则被测模块内部的同级 import（如 archive_types）无法解析。
sys.path.insert(0, str(ROOT / "scripts"))
SPEC = importlib.util.spec_from_file_location("process_batch_cover_tests", ROOT / "scripts" / "process_batch.py")
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


def make_docx(path: Path, lines: list[str]) -> None:
    document = etree.Element(f"{{{MODULE.W}}}document", nsmap={"w": MODULE.W})
    body = etree.SubElement(document, f"{{{MODULE.W}}}body")
    for line in lines:
        paragraph = etree.SubElement(body, f"{{{MODULE.W}}}p")
        if line:
            run = etree.SubElement(paragraph, f"{{{MODULE.W}}}r")
            text = etree.SubElement(run, f"{{{MODULE.W}}}t")
            text.text = line
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>",
        )
        archive.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8"))


def make_textframe_page_docx(path: Path, right_aligned: bool) -> None:
    document = etree.Element(f"{{{MODULE.W}}}document", nsmap={"w": MODULE.W, "r": MODULE.R})
    body = etree.SubElement(document, f"{{{MODULE.W}}}body")
    section = etree.SubElement(body, f"{{{MODULE.W}}}sectPr")
    header_ref = etree.SubElement(section, f"{{{MODULE.W}}}headerReference")
    header_ref.set(f"{{{MODULE.W}}}type", "default")
    header_ref.set(f"{{{MODULE.R}}}id", "rId1")

    header = etree.Element(f"{{{MODULE.W}}}hdr", nsmap={"w": MODULE.W})
    textbox_content = etree.SubElement(header, f"{{{MODULE.W}}}txbxContent")
    paragraph = etree.SubElement(textbox_content, f"{{{MODULE.W}}}p")
    if right_aligned:
        properties = etree.SubElement(paragraph, f"{{{MODULE.W}}}pPr")
        alignment = etree.SubElement(properties, f"{{{MODULE.W}}}jc")
        alignment.set(f"{{{MODULE.W}}}val", "right")
    run = etree.SubElement(paragraph, f"{{{MODULE.W}}}r")
    instruction = etree.SubElement(run, f"{{{MODULE.W}}}instrText")
    instruction.text = " PAGE "

    relationships = etree.Element(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships",
        nsmap={None: "http://schemas.openxmlformats.org/package/2006/relationships"},
    )
    relationship = etree.SubElement(
        relationships, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    )
    relationship.set("Id", "rId1")
    relationship.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
    relationship.set("Target", "header1.xml")

    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8"))
        archive.writestr("word/header1.xml", etree.tostring(header, xml_declaration=True, encoding="UTF-8"))
        archive.writestr(
            "word/_rels/document.xml.rels",
            etree.tostring(relationships, xml_declaration=True, encoding="UTF-8"),
        )


class ManualCoverAuditTests(unittest.TestCase):
    def inspect(self, lines: list[str], title: str = "示例软件") -> dict:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "manual.docx"
            make_docx(path, lines)
            return MODULE.inspect_docx(path, title, check_login=False)

    def test_cover_label_may_follow_empty_paragraphs(self):
        audit = self.inspect(["示例软件", "", "", "使用说明书", "第一章 功能介绍"])

        self.assertTrue(audit["cover_found"])
        self.assertEqual(audit["cover_label"], "使用说明书")

    def test_active_textframe_page_alignment_is_not_an_acceptance_requirement(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            misaligned = root / "misaligned.docx"
            aligned = root / "aligned.docx"
            make_textframe_page_docx(misaligned, right_aligned=False)
            make_textframe_page_docx(aligned, right_aligned=True)

            misaligned_audit = MODULE.inspect_docx_header_only(misaligned, "示例软件")
            aligned_audit = MODULE.inspect_docx_header_only(aligned, "示例软件")

        self.assertEqual(misaligned_audit["header_page_field_counts"], [1])
        self.assertNotIn("textframe_page_paragraphs_right_aligned", misaligned_audit)
        self.assertNotIn("textframe_page_paragraph_count", misaligned_audit)
        self.assertNotIn("textframe_page_paragraphs_right_aligned", aligned_audit)
        self.assertNotIn("textframe_page_paragraph_count", aligned_audit)

    def test_cover_scan_stops_at_first_visible_non_label(self):
        audit = self.inspect(["示例软件", "第一章 功能介绍", "使用说明书"])

        self.assertFalse(audit["cover_found"])
        self.assertEqual(audit["cover_blocks"], [])
        self.assertEqual([row["text"] for row in audit["residual_manual_labels"]], ["使用说明书"])

    def test_incomplete_manual_label_is_not_completed_or_accepted(self):
        audit = self.inspect(["示例软件", "用户操作", "第一章 功能介绍"])

        self.assertFalse(MODULE.is_manual_label_candidate("用户操作"))
        self.assertFalse(audit["cover_found"])
        self.assertIn("用户操作", audit["body_text"])

    def test_pdf_audit_scans_every_page_and_renders_header_only_candidates(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "manual.pdf"
            writer = PdfWriter()
            writer.add_blank_page(width=612, height=792)
            writer.add_blank_page(width=612, height=792)
            with path.open("wb") as stream:
                writer.write(stream)
            with mock.patch.object(MODULE, "_render_pdf_page", return_value={
                "status": "rendered", "foreground_ratio": 0.002, "visually_blank": True,
                "image": str(root / "page.png"),
            }) as render:
                audit = MODULE.inspect_pdf(path, "示例软件", assets_dir=root / "assets", expected_label="用户手册")

        self.assertEqual(audit["pages"], 2)
        self.assertEqual([row["page"] for row in audit["all_pages"]], [1, 2])
        self.assertEqual(audit["confirmed_blank_pages"], [1, 2])
        self.assertEqual(render.call_count, 2)
        self.assertFalse(audit["first_has_title"])
        self.assertFalse(audit["first_has_manual_label"])

    def test_pdf_audit_finds_chinese_and_english_word_bookmark_errors_on_every_page(self):
        class FakePage:
            def __init__(self, text: str):
                self.text = text

            def extract_text(self):
                return self.text

            def get(self, key, default=None):
                return default

        class FakeReader:
            pages = [
                FakePage("示例软件\n使用说明书"),
                FakePage("目录 3.1 系统登录 错误 ！ 未定义 书签 。"),
                FakePage("Error ! Bookmark not   defined ."),
            ]

        with mock.patch.object(MODULE, "PdfReader", return_value=FakeReader()):
            audit = MODULE.inspect_pdf(Path("mock.pdf"), "示例软件", expected_label="使用说明书")

        self.assertEqual(
            [(row["page"], row["type"]) for row in audit["bookmark_errors"]],
            [(2, "未定义书签"), (3, "Bookmark not defined")],
        )
        self.assertIn("错误 ！ 未定义 书签", audit["bookmark_errors"][0]["snippet"])
        self.assertIn("Error ! Bookmark not defined", audit["bookmark_errors"][1]["snippet"])
        findings = MODULE._cover_pdf_findings(audit, word_pages=3, expected_label="使用说明书")
        self.assertTrue(any("第 2 页" in item and "第 3 页" in item for item in findings))


class _Callable:
    def __init__(self, result=None):
        self.result = result

    def __call__(self, *_args, **_kwargs):
        return self.result


class _TimedOutKernel32:
    def __init__(self):
        self.CreateMutexW = _Callable(1)
        self.WaitForSingleObject = _Callable(0x102)
        self.ReleaseMutex = _Callable()
        self.CloseHandle = _Callable()


class WordQueueTests(unittest.TestCase):
    def test_word_queue_times_out_instead_of_waiting_forever(self):
        reporter = mock.Mock()
        kernel32 = _TimedOutKernel32()
        fake_ctypes = mock.Mock(windll=mock.Mock(kernel32=kernel32))

        with mock.patch.object(MODULE.os, "name", "nt"), \
                mock.patch.object(MODULE.ctypes, "windll", fake_ctypes.windll, create=True), \
                mock.patch.object(MODULE.time, "monotonic", side_effect=[0.0, 1.0]):
            with self.assertRaisesRegex(TimeoutError, "等待 Word 处理队列超时"):
                with MODULE.word_automation_lock(reporter, "说明书处理", timeout_seconds=0):
                    self.fail("mutex must not be acquired")

        reporter.update.assert_called_once_with("等待其他任务完成 Word 导出", 0, "说明书处理")


class TransactionalCoverPipelineTests(unittest.TestCase):
    def test_crash_recovery_restores_both_files_from_pending_journal(self):
        with tempfile.TemporaryDirectory() as directory:
            source_root = Path(directory) / "source"
            output_root = Path(directory) / "sample_已处理_20260925"
            project = output_root / "project"
            backup_dir = output_root / MODULE.COVER_BACKUP_DIRECTORY / "txn-1" / "project"
            project.mkdir(parents=True)
            backup_dir.mkdir(parents=True)
            docx = project / "manual.docx"
            pdf = project / "manual.pdf"
            docx.write_bytes(b"new docx written before interruption")
            pdf.write_bytes(b"original pdf")
            (backup_dir / docx.name).write_bytes(b"original docx")
            (backup_dir / pdf.name).write_bytes(b"original pdf")
            journal = output_root / MODULE.COVER_TRANSACTION_FILENAME
            MODULE.atomic_json(journal, {
                "schema_version": 1,
                "transaction_id": "txn-1",
                "state": "PENDING",
                "documents": [{
                    "name": "软件", "folder": "project", "docx": "manual.docx", "pdf": "manual.pdf",
                    "backup_dir": str(backup_dir), "had_pdf": True,
                    "original_hashes": {
                        "docx": MODULE._sha256_path(backup_dir / docx.name),
                        "pdf": MODULE._sha256_path(backup_dir / pdf.name),
                    },
                    "state": "PENDING",
                }],
            })

            recovered = MODULE.recover_pending_cover_transactions(source_root, output_prefix="sample_已处理_")

            self.assertEqual(len(recovered), 1)
            self.assertEqual(docx.read_bytes(), b"original docx")
            self.assertEqual(pdf.read_bytes(), b"original pdf")
            self.assertFalse(journal.exists())
            self.assertFalse((output_root / MODULE.COVER_BACKUP_DIRECTORY).exists())
            self.assertEqual(recovered[0]["actions"][0]["action"], "恢复修复前 DOCX/PDF 快照")

    def test_successful_cover_pair_is_committed_and_transaction_files_are_removed(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            docx = project / "manual.docx"
            pdf = project / "manual.pdf"
            docx.write_bytes(b"original docx")
            pdf.write_bytes(b"original pdf")
            plan_path = root / "plan.json"
            report_path = root / "cover_result.json"
            script_path = root / "pipeline.ps1"
            script_path.write_text("", encoding="utf-8")
            package = {
                "name": "软件", "folder": "project", "docx": docx.name, "pdf": pdf.name,
                "actions": ["ENSURE_COVER_FIRST_PAGE", "AUDIT_ALL_PAGES", "EXPORT_PDF"],
            }
            plan_path.write_text(json.dumps({"base": str(root), "packages": [package]}, ensure_ascii=False), encoding="utf-8")

            class FakeProcess:
                returncode = 0
                pid = 12345

                def __init__(self, command, **_kwargs):
                    docx.write_bytes(b"committed docx")
                    pdf.write_bytes(b"committed pdf")
                    report_file = Path(command[command.index("-ReportPath") + 1])
                    MODULE.atomic_json(report_file, {"packages": [{
                        "name": "软件", "folder": "project", "status": "OK", "word_pages": 2,
                        "cover_label": "用户手册", "failures": [], "attempts": [],
                        "page_audit": [{"page": 1, "has_body_content": True}, {"page": 2, "has_body_content": True}],
                    }]})

                def communicate(self, timeout=None):
                    return b"completed", None

            clean_pdf_audit = {
                "pages": 2, "first_has_title": True, "first_has_manual_label": True,
                "first_label_matches_expected": True, "confirmed_blank_pages": [], "unresolved_blank_pages": [],
                "blank_page_candidates": [],
            }
            with mock.patch.object(MODULE.subprocess, "Popen", side_effect=FakeProcess), \
                    mock.patch.object(MODULE, "inspect_pdf", return_value=clean_pdf_audit), \
                    mock.patch.object(MODULE.page_label_audit, "audit_docx_pdf_page_labels", return_value={"status": "complete", "pages": [], "issues": []}), \
                    mock.patch.object(MODULE, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(MODULE, "list_word_pids", return_value=set()), \
                    mock.patch.object(MODULE, "close_owned_word"):
                report = MODULE.invoke_word_pipeline(root, plan_path, report_path, script_path, module="说明书处理")

            self.assertEqual(report["packages"][0]["status"], "OK")
            self.assertEqual(docx.read_bytes(), b"committed docx")
            self.assertEqual(pdf.read_bytes(), b"committed pdf")
            self.assertFalse((root / MODULE.COVER_TRANSACTION_FILENAME).exists())
            self.assertFalse((root / MODULE.COVER_BACKUP_DIRECTORY).exists())

    def test_bookmark_errors_still_fail_with_indeterminate_page_label_audit(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            docx = project / "manual.docx"
            pdf = project / "manual.pdf"
            docx.write_bytes(b"original docx")
            pdf.write_bytes(b"original pdf")
            original_docx, original_pdf = docx.read_bytes(), pdf.read_bytes()
            plan_path = root / "plan.json"
            report_path = root / "cover_result.json"
            script_path = root / "pipeline.ps1"
            script_path.write_text("", encoding="utf-8")
            plan_path.write_text(json.dumps({"base": str(root), "packages": [{
                "name": "软件", "folder": "project", "docx": docx.name, "pdf": pdf.name,
                "actions": ["ENSURE_COVER_FIRST_PAGE", "AUDIT_ALL_PAGES", "AUDIT_PAGE_LABELS", "EXPORT_PDF"],
            }]}, ensure_ascii=False), encoding="utf-8")
            launches = []

            class FakeProcess:
                returncode = 0
                pid = 12345

                def __init__(self, command, **_kwargs):
                    launches.append(command)
                    docx.write_bytes(b"modified docx")
                    pdf.write_bytes(b"modified pdf")
                    report_file = Path(command[command.index("-ReportPath") + 1])
                    MODULE.atomic_json(report_file, {"packages": [{
                        "name": "软件", "folder": "project", "status": "OK", "word_pages": 2,
                        "cover_label": "用户手册", "failures": [], "attempts": [],
                        "page_audit": [
                            {"page": 1, "has_body_content": True},
                            {"page": 2, "has_body_content": True},
                        ],
                    }]})

                def communicate(self, timeout=None):
                    return b"completed", None

            bookmark_only_audit = {
                "pages": 2, "first_has_title": True, "first_has_manual_label": True,
                "first_label_matches_expected": True, "confirmed_blank_pages": [],
                "unresolved_blank_pages": [], "blank_page_candidates": [],
                "bookmark_errors": [
                    {"page": 2, "type": "未定义书签", "snippet": "错误!未定义书签。"},
                    {"page": 2, "type": "Bookmark not defined", "snippet": "Error! Bookmark not defined."},
                ],
            }
            with mock.patch.object(MODULE.subprocess, "Popen", side_effect=FakeProcess), \
                    mock.patch.object(MODULE, "inspect_pdf", return_value=bookmark_only_audit), \
                    mock.patch.object(MODULE.page_label_audit, "audit_docx_pdf_page_labels", return_value={
                        "status": "indeterminate", "pages": [], "issues": ["缺少 PDF 坐标证据"],
                    }), \
                    mock.patch.object(MODULE, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(MODULE, "list_word_pids", return_value=set()), \
                    mock.patch.object(MODULE, "close_owned_word"):
                report = MODULE.invoke_word_pipeline(root, plan_path, report_path, script_path, module="说明书处理")
            restored_docx, restored_pdf = docx.read_bytes(), pdf.read_bytes()

        result = report["packages"][0]
        self.assertEqual(len(launches), 1)
        self.assertEqual(result["status"], "ROLLED_BACK")
        self.assertTrue(result["snapshot_restored"])
        self.assertFalse(result.get("retry_performed", False))
        self.assertEqual(result["pdf_audit"]["confirmed_blank_pages"], [])
        self.assertEqual(len(result["pdf_audit"]["bookmark_errors"]), 2)
        self.assertEqual(result["pdf_audit"]["page_label_audit"]["status"], "indeterminate")
        self.assertEqual(result["pdf_audit"]["page_label_audit"]["enforcement"], "informational_only")
        self.assertTrue(any("Word 书签引用错误文本" in item for item in result["failures"]))
        self.assertEqual(restored_docx, original_docx)
        self.assertEqual(restored_pdf, original_pdf)

    def test_failed_pdf_blank_page_retry_restores_original_pair(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            docx = project / "manual.docx"
            pdf = project / "manual.pdf"
            docx.write_bytes(b"original docx")
            pdf.write_bytes(b"original pdf")
            original_docx, original_pdf = docx.read_bytes(), pdf.read_bytes()
            plan_path = root / "plan.json"
            report_path = root / "cover_result.json"
            script_path = root / "pipeline.ps1"
            script_path.write_text("", encoding="utf-8")
            plan = {"base": str(root), "packages": [{
                "name": "软件", "name_source": "unique_txt_filename", "docx_title": "旧名称",
                "folder": "project", "docx": docx.name, "pdf": pdf.name,
                "actions": ["ENSURE_COVER_FIRST_PAGE", "AUDIT_ALL_PAGES", "EXPORT_PDF"],
            }]}
            plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
            launches = []

            class FakeProcess:
                returncode = 0
                pid = 12345

                def __init__(self, command, **_kwargs):
                    launches.append(command)
                    docx.write_bytes(b"modified docx")
                    pdf.write_bytes(b"modified pdf")
                    report_file = Path(command[command.index("-ReportPath") + 1])
                    package = json.loads(Path(command[command.index("-PlanPath") + 1]).read_text(encoding="utf-8"))["packages"][0]
                    MODULE.atomic_json(report_file, {"packages": [{
                        "name": package["name"], "folder": package["folder"], "status": "OK",
                        "word_pages": 2, "cover_label": "用户手册", "failures": [], "attempts": [],
                        "page_audit": [
                            {"page": 1, "has_body_content": True},
                            {"page": 2, "has_body_content": False},
                        ],
                    }]})

                def communicate(self, timeout=None):
                    return b"completed", None

            blank_pdf_audit = {
                "pages": 2, "first_has_title": True, "first_has_manual_label": True,
                "first_label_matches_expected": True, "confirmed_blank_pages": [2],
                "unresolved_blank_pages": [],
                "blank_page_candidates": [{"page": 2, "confirmed_blank": True, "suspected_causes": ["Word 页映射为空"]}],
                "bookmark_errors": [],
            }
            bookmark_error_pdf_audit = {
                "pages": 2, "first_has_title": True, "first_has_manual_label": True,
                "first_label_matches_expected": True, "confirmed_blank_pages": [],
                "unresolved_blank_pages": [], "blank_page_candidates": [],
                "bookmark_errors": [
                    {"page": 2, "type": "未定义书签", "snippet": "错误!未定义书签。"},
                    {"page": 2, "type": "Bookmark not defined", "snippet": "Error! Bookmark not defined."},
                ],
            }
            with mock.patch.object(MODULE.subprocess, "Popen", side_effect=FakeProcess), \
                    mock.patch.object(MODULE, "inspect_pdf", side_effect=[blank_pdf_audit, bookmark_error_pdf_audit]), \
                    mock.patch.object(MODULE.page_label_audit, "audit_docx_pdf_page_labels", return_value={"status": "complete", "pages": [], "issues": []}), \
                    mock.patch.object(MODULE, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(MODULE, "list_word_pids", return_value=set()), \
                    mock.patch.object(MODULE, "close_owned_word"):
                report = MODULE.invoke_word_pipeline(root, plan_path, report_path, script_path, module="说明书处理")
            restored_docx = docx.read_bytes()
            restored_pdf = pdf.read_bytes()

        self.assertEqual(len(launches), 2)
        result = report["packages"][0]
        self.assertEqual(result["status"], "ROLLED_BACK")
        self.assertTrue(result["snapshot_restored"])
        self.assertTrue(result["retry_performed"])
        self.assertEqual(max(item["attempt"] for item in result["attempts"]), 2)
        self.assertEqual(len(result["pdf_audit"]["bookmark_errors"]), 2)
        self.assertTrue(any("Word 书签引用错误文本" in item for item in result["failures"]))
        self.assertEqual(restored_docx, original_docx)
        self.assertEqual(restored_pdf, original_pdf)

    def test_confirmed_page_label_truncation_is_reported_without_repair_or_rollback(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            docx, pdf = project / "manual.docx", project / "manual.pdf"
            docx.write_bytes(b"original docx")
            pdf.write_bytes(b"original pdf")
            plan_path, report_path, script_path = root / "plan.json", root / "result.json", root / "pipeline.ps1"
            script_path.write_text("", encoding="utf-8")
            locator = {
                "id": "word/header1.xml#p0", "story_type": "header", "section_indices": [1],
                "paragraph_index": 0, "prefix": "第", "suffix": "页", "ambiguous": False,
            }
            package = {
                "name": "软件", "folder": "project", "docx": docx.name, "pdf": pdf.name,
                "actions": ["AUDIT_ALL_PAGES", "AUDIT_PAGE_LABELS", "EXPORT_PDF"],
            }
            plan_path.write_text(json.dumps({"base": str(root), "packages": [package]}, ensure_ascii=False), encoding="utf-8")
            launched_plans = []

            class FakeProcess:
                returncode = 0
                pid = 12345

                def __init__(self, command, **_kwargs):
                    item_plan = json.loads(Path(command[command.index("-PlanPath") + 1]).read_text(encoding="utf-8"))
                    launched_plans.append(item_plan["packages"][0])
                    docx.write_bytes(b"modified docx")
                    pdf.write_bytes(b"modified pdf")
                    result_file = Path(command[command.index("-ReportPath") + 1])
                    MODULE.atomic_json(result_file, {"packages": [{
                        "name": "软件", "folder": "project", "status": "OK", "word_pages": 2,
                        "failures": [], "attempts": [], "page_audit": [],
                    }]})

                def communicate(self, timeout=None):
                    return b"completed", None

            clean_pdf_audit = {
                "pages": 2, "first_has_title": True, "first_has_manual_label": True,
                "first_label_matches_expected": True, "confirmed_blank_pages": [], "unresolved_blank_pages": [],
                "blank_page_candidates": [],
            }
            truncated = {
                "status": "truncated", "issues": [],
                "template_audit": {"templates": [locator]},
                "pages": [{"page_number": 1, "status": "truncated", "matched_template_id": locator["id"]}],
            }
            with mock.patch.object(MODULE.subprocess, "Popen", side_effect=FakeProcess), \
                    mock.patch.object(MODULE, "inspect_pdf", return_value=clean_pdf_audit), \
                    mock.patch.object(MODULE.page_label_audit, "audit_docx_pdf_page_labels", return_value=truncated), \
                    mock.patch.object(MODULE, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(MODULE, "list_word_pids", return_value=set()), \
                    mock.patch.object(MODULE, "close_owned_word"):
                report = MODULE.invoke_word_pipeline(root, plan_path, report_path, script_path, module="说明书处理")

        self.assertEqual(len(launched_plans), 1)
        self.assertEqual(report["packages"][0]["status"], "OK")
        self.assertFalse(report["packages"][0]["snapshot_restored"])
        self.assertEqual(report["packages"][0]["pdf_audit"]["page_label_audit"]["status"], "truncated")
        self.assertEqual(report["packages"][0]["pdf_audit"]["page_label_audit"]["enforcement"], "informational_only")

    def test_indeterminate_page_label_audit_is_reported_without_rollback(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            docx, pdf = project / "manual.docx", project / "manual.pdf"
            docx.write_bytes(b"original docx")
            pdf.write_bytes(b"original pdf")
            plan_path, report_path, script_path = root / "plan.json", root / "result.json", root / "pipeline.ps1"
            script_path.write_text("", encoding="utf-8")
            plan_path.write_text(json.dumps({"base": str(root), "packages": [{
                "name": "软件", "folder": "project", "docx": docx.name, "pdf": pdf.name,
                "actions": ["AUDIT_ALL_PAGES", "AUDIT_PAGE_LABELS", "EXPORT_PDF"],
            }]}, ensure_ascii=False), encoding="utf-8")

            class FakeProcess:
                returncode = 0
                pid = 12345

                def __init__(self, command, **_kwargs):
                    docx.write_bytes(b"modified docx")
                    pdf.write_bytes(b"modified pdf")
                    report_file = Path(command[command.index("-ReportPath") + 1])
                    MODULE.atomic_json(report_file, {"packages": [{
                        "name": "软件", "folder": "project", "status": "OK", "word_pages": 1,
                        "failures": [], "attempts": [],
                    }]})

                def communicate(self, timeout=None):
                    return b"completed", None

            uncertain = {"status": "indeterminate", "pages": [], "issues": ["缺少 PDF 坐标证据"]}
            clean_pdf_audit = {
                "pages": 1, "first_has_title": True, "first_has_manual_label": True,
                "first_label_matches_expected": True, "confirmed_blank_pages": [],
                "unresolved_blank_pages": [], "blank_page_candidates": [], "bookmark_errors": [],
            }
            with mock.patch.object(MODULE.subprocess, "Popen", side_effect=FakeProcess), \
                    mock.patch.object(MODULE, "inspect_pdf", return_value=clean_pdf_audit), \
                    mock.patch.object(MODULE.page_label_audit, "audit_docx_pdf_page_labels", return_value=uncertain), \
                    mock.patch.object(MODULE, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(MODULE, "list_word_pids", return_value=set()), \
                    mock.patch.object(MODULE, "close_owned_word"):
                report = MODULE.invoke_word_pipeline(root, plan_path, report_path, script_path, module="说明书处理")
            final_docx, final_pdf = docx.read_bytes(), pdf.read_bytes()

        self.assertEqual(report["packages"][0]["status"], "OK")
        self.assertFalse(report["packages"][0]["snapshot_restored"])
        self.assertIn("缺少 PDF 坐标证据", json.dumps(report, ensure_ascii=False))
        self.assertEqual(report["packages"][0]["pdf_audit"]["page_label_audit"]["enforcement"], "informational_only")
        self.assertEqual(final_docx, b"modified docx")
        self.assertEqual(final_pdf, b"modified pdf")

    def test_page_label_audit_error_is_reported_without_rollback(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            docx, pdf = project / "manual.docx", project / "manual.pdf"
            docx.write_bytes(b"original docx")
            pdf.write_bytes(b"original pdf")
            plan_path, report_path, script_path = root / "plan.json", root / "result.json", root / "pipeline.ps1"
            script_path.write_text("", encoding="utf-8")
            plan_path.write_text(json.dumps({"base": str(root), "packages": [{
                "name": "软件", "folder": "project", "docx": docx.name, "pdf": pdf.name,
                "actions": ["AUDIT_PAGE_LABELS", "EXPORT_PDF"],
            }]}, ensure_ascii=False), encoding="utf-8")

            class FakeProcess:
                returncode = 0
                pid = 12345

                def __init__(self, command, **_kwargs):
                    docx.write_bytes(b"modified docx")
                    pdf.write_bytes(b"modified pdf")
                    report_file = Path(command[command.index("-ReportPath") + 1])
                    MODULE.atomic_json(report_file, {"packages": [{
                        "name": "软件", "folder": "project", "status": "OK", "word_pages": 1,
                        "failures": [], "attempts": [],
                    }]})

                def communicate(self, timeout=None):
                    return b"completed", None

            with mock.patch.object(MODULE.subprocess, "Popen", side_effect=FakeProcess), \
                    mock.patch.object(MODULE.page_label_audit, "audit_docx_pdf_page_labels", side_effect=RuntimeError("读取页码坐标失败")), \
                    mock.patch.object(MODULE, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(MODULE, "list_word_pids", return_value=set()), \
                    mock.patch.object(MODULE, "close_owned_word"):
                report = MODULE.invoke_word_pipeline(root, plan_path, report_path, script_path, module="说明书处理")
            final_docx, final_pdf = docx.read_bytes(), pdf.read_bytes()
        result = report["packages"][0]
        self.assertEqual(result["status"], "OK")
        self.assertFalse(result["snapshot_restored"])
        self.assertEqual(result["pdf_audit"]["page_label_audit"]["status"], "error")
        self.assertEqual(result["pdf_audit"]["page_label_audit"]["enforcement"], "informational_only")
        self.assertIn("读取页码坐标失败", json.dumps(result, ensure_ascii=False))
        self.assertEqual(final_docx, b"modified docx")
        self.assertEqual(final_pdf, b"modified pdf")


class PowerShellCoverSafetyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.script = (ROOT / "scripts" / "word_manual_pipeline_stable.ps1").read_text(encoding="utf-8-sig")

    def test_residual_label_with_content_or_objects_is_report_only(self):
        start = self.script.index("if ($null -ne $residualLabel)")
        end = self.script.index("if ($null -ne $isolated)", start)
        branch = self.script[start:end]

        self.assertIn("$otherText.Count -gt 0 -or $hasObjects", branch)
        self.assertIn("$minimumPage = [Math]::Max($minimumPage, $residualLabel.Page + 1)", branch)
        self.assertIn("continue", branch)
        self.assertLess(branch.index("continue"), branch.index(".Delete()"))

    def test_cover_boundary_includes_images_and_tables(self):
        start = self.script.index("function Ensure-CoverPageBoundary")
        end = self.script.index("function Remove-BodyLeadingBlankParagraphs", start)
        function = self.script[start:end]

        self.assertIn("$Document.InlineShapes", function)
        self.assertIn("$Document.Tables", function)
        self.assertIn("$extra.Count -eq 0 -and -not $hasObjects", function)

    def test_cover_keeps_existing_page_dimensions(self):
        start = self.script.index("function Add-Cover")
        end = self.script.index("function Remove-DuplicateCoverBlocks", start)
        function = self.script[start:end]

        self.assertNotIn("PageWidth", function)
        self.assertNotIn("PageHeight", function)
        self.assertNotIn("wdPaperA4", function)

    def test_two_line_title_is_allowed_but_short_last_line_is_flagged(self):
        start = self.script.index("function Set-CoverTypography")
        end = self.script.index("function Remove-CoverVersion", start)
        function = self.script[start:end]
        checks = r'''
$ast = (Get-Command Set-CoverTypography).ScriptBlock.Ast
$guardClause = $null
foreach ($node in $ast.FindAll({ param($candidate) $candidate -is [System.Management.Automation.Language.IfStatementAst] }, $true)) {
    foreach ($clause in $node.Clauses) {
        $variables = @($clause.Item1.FindAll({ param($part) $part -is [System.Management.Automation.Language.VariableExpressionAst] }, $true) | ForEach-Object { $_.VariablePath.UserPath.ToLowerInvariant() })
        if ($variables -contains 'lines' -and $variables -contains 'orphanedlastline' -and $variables -contains 'labellayoutfailed' -and $clause.Item2.Extent.Text -match '\bbreak\b') {
            $guardClause = $clause.Item1
            break
        }
    }
    if ($null -ne $guardClause) { break }
}
if ($null -eq $guardClause) { throw 'Could not find the typography convergence exit condition' }
$guard = [scriptblock]::Create($guardClause.Extent.Text)
$cases = @(
    [pscustomobject]@{ Lines = 2; Orphaned = $false; LabelFailed = $false; Expected = $true },
    [pscustomobject]@{ Lines = 2; Orphaned = $false; LabelFailed = $true; Expected = $false },
    [pscustomobject]@{ Lines = 3; Orphaned = $false; LabelFailed = $false; Expected = $false },
    [pscustomobject]@{ Lines = 2; Orphaned = $true; LabelFailed = $false; Expected = $false }
)
foreach ($case in $cases) {
    $lines = $case.Lines; $orphanedLastLine = $case.Orphaned; $labelLayoutFailed = $case.LabelFailed
    if ([bool](& $guard) -ne $case.Expected) { throw "Unexpected convergence decision for $($case | ConvertTo-Json -Compress)" }
}
'''
        encoded = base64.b64encode((function + "\n" + checks).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr or result.stdout)
        self.assertIn('if ($doCover -and $coverAudit.Lines -gt 2)', self.script)
        self.assertIn('if ($doCover -and $coverAudit.OrphanedLastLine)', self.script)

    def test_confirmed_title_only_cover_gets_missing_label_without_second_cover(self):
        start = self.script.index("function Ensure-FirstPageCover")
        end = self.script.index("function Get-WordPageAudit", start)
        function = self.script[start:end]
        self.assertIn('Result = "LABEL_ADDED"', function)
        self.assertIn('InsertBefore("使用说明书`r")', function)
        self.assertIn("Add-Cover $Document $Title", function)
        self.assertIn("EXISTING_COVER", function)

    def test_cover_style_can_format_existing_label_text(self):
        start = self.script.index("function Set-CoverTypography")
        end = self.script.index("function Remove-CoverVersion", start)
        function = self.script[start:end]

        # Dynamic table/TOC/body protection is covered by test_cover_performance.
        self.assertIn("$_.Page -eq 1 -and -not $_.InTable -and $_.Style -notmatch '(?i)TOC|目录'", function)
        self.assertIn("$items = @(Get-FrontParagraphs $Document 2)", function)
        self.assertIn("$block = Find-CoverBlock $items $Title 1 1 $true", function)
        self.assertIn("if ($null -ne $labelItem)", function)
        self.assertNotIn("$labelRange.Text =", function)
        self.assertIn("LabelPresent", function)
        self.assertIn("$between | Sort-Object Index -Descending", function)
        self.assertIn("$_.Text -ne \"\"", function)

    def test_duplicate_qa_requires_a_standalone_cover_page(self):
        start = self.script.index("function Test-DuplicateCoverBlock")
        end = self.script.index("function Get-CoverLastLineUnitCount", start)
        function = self.script[start:end]

        self.assertIn("$otherText.Count -eq 0", function)
        self.assertIn("Test-PageHasObjects", function)
        self.assertIn("Find-IsolatedCombinedCoverFragment", function)

    def test_redundant_post_cover_breaks_are_removed_after_first_boundary(self):
        start = self.script.index("function Remove-RedundantPostCoverFrontMatter")
        end = self.script.index("function Story-HasContent", start)
        function = self.script[start:end]

        self.assertIn("$protectedBoundaryFound", function)
        self.assertIn("$Document.Sections.Item($sectionIndex).Range.End", function)
        self.assertIn("Test-CombinedCoverFragment", function)
        self.assertIn("$realBody.Page -lt 3", function)
        self.assertIn("Sort-Object Index -Descending", function)

    def test_real_output_post_cover_fragments_match_only_front_matter(self):
        function_names = (
            "Clean-Text", "Compact-Text", "Strip-OrdinalPrefix",
            "Test-CoverTitle", "Test-ManualLabel", "Test-PostCoverFragment",
        )
        definitions = []
        for name in function_names:
            start = self.script.index(f"function {name}(")
            end = self.script.find("\nfunction ", start + 1)
            definitions.append(self.script[start:end if end != -1 else None])
        cases = [
            ("弘元绿能金刚线蓝宝石切片机控制系统软件", "弘元绿能金刚线蓝宝石切片机控制系统软件", True),
            ("用户手册", "医院人力资源综合管理系统", True),
            ("技术说明手册", "建筑项目数据综合分析平台", True),
            ("文档编号:ARCH-ENERGY-2024-UM01 系统状态:正在运行 机密程度:内部使用", "建筑项目数据综合分析平台", True),
            ("目录", "建筑项目数据综合分析平台", False),
            ("1. 系统概述", "建筑项目数据综合分析平台", False),
            ("如上图所示可以看到切割过程中，参数在发生变化。", "弘元绿能金刚线蓝宝石切片机控制系统软件", False),
            ("建筑项目数据综合分析平台是一款面向现代化大型建筑工程的数字化能效管理平台", "建筑项目数据综合分析平台", False),
        ]
        rule_path = ROOT / "scripts" / "cover_rules.ps1"
        commands = [f". '{str(rule_path)}'", *definitions]
        for text, title, expected in cases:
            commands.append(f"if ((Test-PostCoverFragment '{text}' '{title}') -ne ${str(expected).lower()}) {{ throw 'fragment rule mismatch' }}")
        encoded = base64.b64encode("\n".join(commands).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_content_page_cleanup_deletes_only_leading_fragment_paragraph(self):
        start = self.script.index("function Remove-DuplicateCoverBlocks")
        end = self.script.index("function Test-DuplicateCoverBlock", start)
        function = self.script[start:end]
        self.assertIn("Test-PostCoverFragment", function)
        self.assertIn("Get-FrontParagraphs $Document 5", function)
        self.assertIn("$leadingFragment.Page -eq 2", function)
        self.assertIn("$Document.Range($leadingFragment.Start, $leadingFragment.End).Delete()", function)
        self.assertIn("$wdWithInTable", function)
        self.assertIn("Test-RangeHasProtectedContent $Document $fragmentRange $true", function)
        self.assertIn('or (Test-DuplicateCoverBlock $doc $title)', self.script)

    def test_blank_range_protection_rules_execute(self):
        start = self.script.index("function Test-RangeHasProtectedContent(")
        end = self.script.index("\nfunction ", start + 1)
        function = self.script[start:end]
        checks = r'''
function New-Range($text = '', $pageBreak = $false) {
    return [pscustomobject]@{ Start = 10; End = 20; Text = $text; ParagraphFormat = [pscustomobject]@{ PageBreakBefore = $pageBreak } }
}
function New-Document {
    return [pscustomobject]@{ Sections = @(); InlineShapes = @(); Shapes = @(); Tables = @() }
}
$doc = New-Document
if (Test-RangeHasProtectedContent $doc (New-Range)) { throw 'plain blank was protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)))) { throw 'page break was not protected' }
if (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)) $false $true) { throw 'cover-only page break was protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range ([string][char]14)) $false $true $true)) { throw 'column break was not protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range '' $true) $true)) { throw 'PageBreakBefore was not protected' }
$doc.Sections = @([pscustomobject]@{ Range = [pscustomobject]@{ End = 19 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'section end was not protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)) $false $true)) { throw 'section end was lost when allowing page break' }
if (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)) $false $true $true) { throw 'standalone cover section break was protected' }
$doc.Sections = @()
$doc.InlineShapes = @([pscustomobject]@{ Range = [pscustomobject]@{ Start = 15; End = 16 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'inline image was not protected' }
$doc.InlineShapes = @()
$doc.Shapes = @([pscustomobject]@{ Anchor = [pscustomobject]@{ Start = 15 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'floating image was not protected' }
$doc.Shapes = @()
$doc.Tables = @([pscustomobject]@{ Range = [pscustomobject]@{ Start = 15; End = 18 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'table was not protected' }
'''
        encoded = base64.b64encode((function + "\n" + checks).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_floating_shape_counts_as_page_object(self):
        start = self.script.index("function Test-PageHasObjects(")
        end = self.script.index("\nfunction ", start + 1)
        function = self.script[start:end]
        checks = r'''
$wdActiveEndPageNumber = 3
$anchor = [pscustomobject]@{}
$anchor | Add-Member -MemberType ScriptMethod -Name Information -Value { param($kind) return 2 }
$doc = [pscustomobject]@{ InlineShapes = @(); Tables = @(); Shapes = @([pscustomobject]@{ Anchor = $anchor }) }
if (-not (Test-PageHasObjects $doc 2)) { throw 'floating shape on page 2 was missed' }
if (Test-PageHasObjects $doc 1) { throw 'floating shape was assigned to wrong page' }
'''
        encoded = base64.b64encode((function + "\n" + checks).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_cover_style_uses_com_double_values(self):
        start = self.script.index("function Set-CoverTypography")
        end = self.script.index("function Remove-CoverVersion", start)
        function = self.script[start:end]
        checks = r'''
$ast = (Get-Command Set-CoverTypography).ScriptBlock.Ast
$assignments = @($ast.FindAll({ param($candidate) $candidate -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))
$doubleSizeAssignments = @($assignments | Where-Object {
    $left = $_.Left.Extent.Text -replace '\s', ''
    $right = $_.Right.Extent.Text -replace '\s', ''
    $left -match '(?i)\.Font\.Size$' -and $right -match '^\[double\]\$size$'
})
$labelSpacingAssignments = @($assignments | Where-Object {
    $left = $_.Left.Extent.Text -replace '\s', ''
    $right = $_.Right.Extent.Text -replace '\s', ''
    $left -match '(?i)\.ParagraphFormat\.SpaceBefore$' -and $right -match '^\[double\]48$'
})
if ($doubleSizeAssignments.Count -lt 2) { throw 'Title and label font sizes must both use double values' }
if ($labelSpacingAssignments.Count -lt 1) { throw 'The label must retain 48 points of spacing using a double value' }
'''
        encoded = base64.b64encode((function + "\n" + checks).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr or result.stdout)


if __name__ == "__main__":
    unittest.main()

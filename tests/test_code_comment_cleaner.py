from __future__ import annotations

import importlib.util
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
SPEC = importlib.util.spec_from_file_location("code_comment_cleaner", ROOT / "scripts" / "code_comment_cleaner.py")
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


def make_docx(path: Path, lines: list[str]) -> None:
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    document = etree.Element(f"{{{ns}}}document", nsmap={"w": ns})
    body = etree.SubElement(document, f"{{{ns}}}body")
    for line in lines:
        paragraph = etree.SubElement(body, f"{{{ns}}}p")
        run = etree.SubElement(paragraph, f"{{{ns}}}r")
        text = etree.SubElement(run, f"{{{ns}}}t")
        text.text = line
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
        archive.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8"))


def read_lines(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    return ["".join(p.xpath(".//w:t/text()", namespaces=MODULE.NS)) for p in MODULE._paragraphs(root)]


class CommentLexerTests(unittest.TestCase):
    def cleaned(self, source: str) -> str:
        ranges, _ = MODULE.find_comment_ranges(source)
        for start, end, _ in sorted(ranges, reverse=True):
            source = source[:start] + source[end:]
        return source

    def test_supported_comments_and_english_preservation(self):
        source = "\n".join([
            "int x = 1; // 中文注释",
            "// English comment",
            "#中文注释",
            "value = 1  # Python行尾中文",
            "-- SQL中文",
            "/* CSS中文 */ body {}",
            "<!-- HTML中文 --> <p>中文页面</p>",
        ])
        cleaned = self.cleaned(source)
        self.assertIn("int x = 1; ", cleaned)
        self.assertIn("// English comment", cleaned)
        self.assertIn("body {}", cleaned)
        self.assertIn("<p>中文页面</p>", cleaned)
        self.assertNotIn("中文注释", cleaned)
        self.assertNotIn("Python中文", cleaned)
        self.assertIn("value = 1  ", cleaned)

    def test_css_hex_color_is_not_python_comment(self):
        source = "color: #fff;\nbackground: #中文;\n#中文组件 { color: red; }\ndiv#中文组件 { color: red; }\na { background: url(//中文/图.png); }"
        self.assertEqual(self.cleaned(source), source)

    def test_python_and_sql_comments_without_separator_space_are_removed(self):
        source = "if ready: # 中文\nx=1# 中文\nSELECT 1-- 中文\nFROM users-- 中文\nVALUES(1)-- 中文"
        cleaned = self.cleaned(source)
        self.assertEqual(cleaned, "if ready: \nx=1\nSELECT 1\nFROM users\nVALUES(1)")

    def test_css_custom_properties_are_not_sql_comments(self):
        source = "--中文主题: red;\ncolor: var(--中文主题);"
        self.assertEqual(self.cleaned(source), source)

    def test_python_floor_division_is_preserved_before_hash_comment(self):
        source = "\n".join([
            "pages = total // size  # 中文注释",
            "negative = total // -size #负数",
            "positive = total // +2 #正数",
            "masked = total // ~mask #按位取反",
            "total //= size  # 更新页数",
        ])
        self.assertEqual(self.cleaned(source), "\n".join([
            "pages = total // size  ",
            "negative = total // -size ",
            "positive = total // +2 ",
            "masked = total // ~mask ",
            "total //= size  ",
        ]))

    def test_decrement_operators_are_preserved_before_real_comments(self):
        source = "count--; // 中文注释\ni--; /* 中文注释 */\n--index; // 前置递减"
        self.assertEqual(self.cleaned(source), "count--; \ni--; \n--index; ")

    def test_unicode_python_operand_is_retained_for_review_when_ambiguous(self):
        source = "result = value // 中文变量"
        ranges, uncertain = MODULE.find_comment_ranges(source)
        self.assertEqual(ranges, [])
        self.assertEqual(len(uncertain), 1)
        self.assertEqual(self.cleaned(source), source)

    def test_return_with_han_after_double_slash_is_ambiguous(self):
        source = "return a // 中文注释"
        ranges, uncertain = MODULE.find_comment_ranges(source)
        self.assertEqual(ranges, [])
        self.assertEqual(len(uncertain), 1)

    def test_indented_line_only_double_slash_comment_is_deleted(self):
        source = "    // 创建场景"
        self.assertEqual(self.cleaned(source), "    ")

    def test_javascript_context_wins_over_accidentally_valid_python_ast(self):
        source = "\n".join([
            "const state = {};",
            "this.size = 200 // 视口基准尺度",
            "setState(value) // 更新状态",
        ])
        self.assertEqual(self.cleaned(source), "\n".join([
            "const state = {};",
            "this.size = 200 ",
            "setState(value) ",
        ]))

    def test_cjk_extension_b_triggers_deletion(self):
        self.assertEqual(self.cleaned("// \U00020000"), "")

    def test_literals_templates_regex_and_html_text_are_protected(self):
        source = "\n".join([
            'String url = "https://example.test/中文";',
            "const s = `// 中文不是注释`;",
            r"const re = /\/\/ 中文/;",
            "text = '''# 中文文档字符串''';",
            "<h1>中文页面文字</h1>",
        ])
        self.assertEqual(self.cleaned(source), source)

    def test_unclosed_block_comment_is_retained_for_review(self):
        source = "code(); /* 未闭合中文"
        ranges, uncertain = MODULE.find_comment_ranges(source)
        self.assertEqual(ranges, [])
        self.assertEqual(len(uncertain), 1)


class DocxTests(unittest.TestCase):
    def test_clean_docx_deletes_comment_paragraph_and_keeps_code(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "项目代码.docx"
            make_docx(path, ["// 整行中文注释", "run(); // 行尾中文", 'print("中文不是注释")', "// English"])
            result = MODULE.clean_docx(path)
            self.assertEqual(result["deletions"], 2)
            self.assertEqual(read_lines(path), ["run();", 'print("中文不是注释")', "// English"])
            with zipfile.ZipFile(path) as archive:
                self.assertIsNone(archive.testzip())

    def test_only_code_named_docx_matches(self):
        self.assertTrue(MODULE.CODE_DOCX.search("产品代码.docx"))
        self.assertFalse(MODULE.CODE_DOCX.search("产品说明.docx"))

    def test_only_blank_lines_joined_by_deleted_comment_are_collapsed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "项目代码.docx"
            make_docx(path, ["", "", "code();", "", "// 删除", "", "next();"])
            MODULE.clean_docx(path)
            self.assertEqual(read_lines(path), ["", "", "code();", "", "next();"])

    def test_section_paragraph_and_last_table_paragraph_are_retained(self):
        ns = MODULE.W_NS
        document = etree.Element(f"{{{ns}}}document", nsmap={"w": ns})
        body = etree.SubElement(document, f"{{{ns}}}body")
        section_p = etree.SubElement(body, f"{{{ns}}}p")
        ppr = etree.SubElement(section_p, f"{{{ns}}}pPr")
        etree.SubElement(ppr, f"{{{ns}}}sectPr")
        run = etree.SubElement(section_p, f"{{{ns}}}r")
        etree.SubElement(run, f"{{{ns}}}t").text = "// 中文"
        table = etree.SubElement(body, f"{{{ns}}}tbl")
        cell = etree.SubElement(etree.SubElement(table, f"{{{ns}}}tr"), f"{{{ns}}}tc")
        cell_p = etree.SubElement(cell, f"{{{ns}}}p")
        cell_run = etree.SubElement(cell_p, f"{{{ns}}}r")
        etree.SubElement(cell_run, f"{{{ns}}}t").text = "// 中文"
        root, _, changed = MODULE.analyze_document_xml(etree.tostring(document))
        self.assertEqual(changed, 2)
        self.assertEqual(len(root.xpath("//w:pPr/w:sectPr", namespaces=MODULE.NS)), 1)
        self.assertEqual(len(root.xpath("//w:tc/w:p", namespaces=MODULE.NS)), 1)


class WorkflowSafetyTests(unittest.TestCase):
    class Reporter:
        def __init__(self):
            self.total = 0
            self.events = []

        def set_total(self, total):
            self.total = total

        def update(self, phase, current=0, name="", status="RUNNING", error=None, **details):
            self.events.append({"phase": phase, "current": current, "name": name, "status": status, "error": error, **details})

    def _preview_report(self, base: Path, extracted: Path) -> tuple[dict, dict]:
        archive = base / "batch.rar"
        archive.write_bytes(b"archive")
        with mock.patch.object(MODULE, "extract_cached", return_value=extracted), mock.patch.object(MODULE, "material_root", return_value=extracted):
            result = MODULE.run(archive, base / "work", ROOT, apply=False)
        report = __import__("json").loads(Path(result["report_json"]).read_text(encoding="utf-8"))
        return result, report

    def test_archive_failure_is_isolated_and_report_is_written(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            first, second = base / "a.rar", base / "b.rar"
            first.write_bytes(b"a")
            second.write_bytes(b"b")
            extracted = base / "extracted"
            extracted.mkdir()
            with mock.patch.object(MODULE, "extract_cached", side_effect=[RuntimeError("bad archive"), extracted]):
                result = MODULE.run(base, base / "work", ROOT, apply=False)
            report = __import__("json").loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(len(report["archives"]), 2)
            self.assertEqual(report["archives"][0]["status"], "FAILED")
            self.assertEqual(report["archives"][1]["status"], "PARTIAL")
            self.assertEqual(report["archives"][1]["documents"][0]["status"], "MISSING_CODE_DOCX")

    def test_progress_is_batch_wide_monotonic_and_accumulates_failures(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            first, second = base / "a.rar", base / "b.rar"
            first.write_bytes(b"a")
            second.write_bytes(b"b")
            extracted = base / "source"
            extracted.mkdir()
            make_docx(extracted / "项目代码.docx", ["// 中文", "run();"])
            (extracted / "项目代码.pdf").write_bytes(b"pdf")
            reporter = self.Reporter()
            with mock.patch.object(MODULE, "extract_cached", side_effect=[RuntimeError("bad archive"), extracted]), mock.patch.object(MODULE, "material_root", return_value=extracted):
                result = MODULE.run(base, base / "work", ROOT, apply=False, progress=reporter)
            currents = [event["current"] for event in reporter.events]
            failure_counts = [event.get("failures", 0) for event in reporter.events]
            self.assertEqual(reporter.total, 2)
            self.assertEqual(currents, sorted(currents))
            self.assertEqual(currents[-1], reporter.total)
            self.assertEqual(failure_counts, sorted(failure_counts))
            self.assertEqual(failure_counts[-1], result["failures"])

    def test_pdf_failure_returns_partial_status_and_preserves_old_pdf(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            archive = base / "batch.rar"
            archive.write_bytes(b"archive")
            extracted = base / "source"
            extracted.mkdir()
            docx = extracted / "项目代码.docx"
            pdf = extracted / "项目代码.pdf"
            make_docx(docx, ["// 中文", "run();"])
            pdf.write_bytes(b"old-pdf")
            with mock.patch.object(MODULE, "extract_cached", return_value=extracted), mock.patch.object(MODULE, "material_root", return_value=extracted), mock.patch.object(MODULE, "_export_pdf", return_value=(False, "failed")):
                result = MODULE.run(archive, base / "work", ROOT, apply=True)
            self.assertEqual(result["status"], "CODE_COMMENT_CLEANUP_PARTIAL")
            self.assertGreater(result["failures"], 0)
            output_pdf = next(Path(result["outputs"][0]).rglob("项目代码.pdf"))
            self.assertEqual(output_pdf.read_bytes(), b"old-pdf")

    def test_unsafe_archive_member_is_rejected(self):
        completed = mock.Mock(returncode=0, stdout="Path = C:\\safe\\a.rar\n----------\nPath = ..\\escape.txt\n")
        with mock.patch.object(MODULE.subprocess, "run", return_value=completed):
            with self.assertRaises(ValueError):
                MODULE._validate_archive_members(Path("7z.exe"), Path("C:/safe/a.rar"))

    def test_docx_without_pdf_is_reported_as_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            extracted = base / "source"
            extracted.mkdir()
            make_docx(extracted / "项目代码.docx", ["run();"])
            result, report = self._preview_report(base, extracted)
            self.assertGreater(result["failures"], 0)
            self.assertEqual(report["archives"][0]["status"], "PARTIAL")
            self.assertEqual(report["archives"][0]["documents"][0]["status"], "SKIPPED_MISSING_PDF")

    def test_pdf_without_docx_is_reported_as_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            extracted = base / "source"
            extracted.mkdir()
            (extracted / "项目代码.pdf").write_bytes(b"pdf")
            result, report = self._preview_report(base, extracted)
            self.assertGreater(result["failures"], 0)
            self.assertEqual(report["archives"][0]["status"], "PARTIAL")
            self.assertEqual(report["archives"][0]["documents"][0]["status"], "SKIPPED_MISSING_DOCX")

    def test_archive_without_code_docx_or_pdf_is_reported_as_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            extracted = base / "source"
            extracted.mkdir()
            (extracted / "项目说明.txt").write_text("不处理", encoding="utf-8")
            result, report = self._preview_report(base, extracted)
            self.assertGreater(result["failures"], 0)
            self.assertEqual(report["archives"][0]["status"], "PARTIAL")
            self.assertEqual(report["archives"][0]["documents"][0]["status"], "MISSING_CODE_DOCX")

    def test_extracted_folder_is_scanned_directly_without_rar(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            make_docx(base / "项目代码.docx", ["run();"])
            (base / "项目代码.pdf").write_bytes(b"pdf")
            result = MODULE.run(base, base / "work", ROOT, apply=False)
            report = __import__("json").loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(len(report["archives"]), 1)
            self.assertEqual(report["archives"][0]["archive"], str(base.resolve()))
            self.assertEqual(report["archives"][0]["documents"][0]["status"], "READY")

    def test_extracted_folder_with_project_subfolders_is_scanned(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            for project in ("甲项目", "乙项目"):
                folder = base / project
                folder.mkdir()
                make_docx(folder / "代码.docx", ["run();"])
                (folder / "代码.pdf").write_bytes(b"pdf")
            result = MODULE.run(base, base / "work", ROOT, apply=False)
            report = __import__("json").loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(len(report["archives"]), 1)
            documents = report["archives"][0]["documents"]
            self.assertEqual(len(documents), 2)
            self.assertTrue(all(item["status"] == "READY" for item in documents))

    def test_folder_without_archive_or_code_material_raises_clear_error(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "项目说明.txt").write_text("不处理", encoding="utf-8")
            with self.assertRaises(ValueError) as raised:
                MODULE.run(base, base / "work", ROOT, apply=False)
            message = str(raised.exception)
            self.assertIn("没有压缩包", message)
            self.assertIn("代码材料", message)
            # 文案不再把用户限定到 RAR：压缩包类型是宽集合
            self.assertNotIn("未找到 RAR", message)


class LauncherTests(unittest.TestCase):
    def test_cmd_launcher_uses_windows_crlf_line_endings(self):
        content = (ROOT / "启动EasySoftware.cmd").read_bytes()
        self.assertNotIn(b"\n", content.replace(b"\r\n", b""))


if __name__ == "__main__":
    unittest.main()

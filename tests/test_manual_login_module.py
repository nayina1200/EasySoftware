from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from lxml import etree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import manual_login_module as login
import process_batch

W, NS = login.W, login.NS
LOGIN = "请在登录界面输入用户名和密码，点击登录按钮进入系统。"


def paragraph(value: str = "", *, outline: int | None = None, num: tuple[str, int] | None = None) -> ET._Element:
    p = ET.Element(W + "p")
    if outline is not None or num:
        props = ET.SubElement(p, W + "pPr")
        if num:
            n = ET.SubElement(props, W + "numPr")
            ET.SubElement(n, W + "ilvl").set(W + "val", str(num[1]))
            ET.SubElement(n, W + "numId").set(W + "val", num[0])
        if outline is not None:
            ET.SubElement(props, W + "outlineLvl").set(W + "val", str(outline))
    if value:
        ET.SubElement(ET.SubElement(p, W + "r"), W + "t").text = value
    return p


def picture() -> ET._Element:
    p = paragraph()
    drawing = ET.SubElement(ET.SubElement(p, W + "r"), W + "drawing")
    inline = ET.SubElement(drawing, "{" + NS["wp"] + "}inline")
    ET.SubElement(inline, "{" + NS["a"] + "}blip").set("{" + NS["r"] + "}embed", "rId7")
    return p


def write_docx(path: Path, nodes: list[ET._Element], *, numbering: bytes | None = None) -> None:
    doc = ET.Element(W + "document", nsmap={k: v for k, v in NS.items()})
    body = ET.SubElement(doc, W + "body")
    body.extend(nodes)
    ET.SubElement(body, W + "sectPr")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", ET.tostring(doc))
        archive.writestr("word/media/image1.png", b"unchanged original image bytes")
        archive.writestr("word/header1.xml", b"unchanged header")
        if numbering:
            archive.writestr("word/numbering.xml", numbering)


def body_nodes(path: Path) -> list[ET._Element]:
    with zipfile.ZipFile(path) as archive:
        return list(ET.fromstring(archive.read("word/document.xml")).find("w:body", NS))


def numbered_definition() -> bytes:
    numbering = ET.Element(W + "numbering", nsmap={"w": W[1:-1]})
    abstract = ET.SubElement(numbering, W + "abstractNum")
    abstract.set(W + "abstractNumId", "0")
    for i in range(3):
        level = ET.SubElement(abstract, W + "lvl")
        level.set(W + "ilvl", str(i))
        for tag, value in (("start", "1"), ("numFmt", "decimal"), ("lvlText", ".".join("%" + str(j + 1) for j in range(i + 1)))):
            ET.SubElement(level, W + tag).set(W + "val", value)
    num = ET.SubElement(numbering, W + "num")
    num.set(W + "numId", "1")
    ET.SubElement(num, W + "abstractNumId").set(W + "val", "0")
    return ET.tostring(numbering)


class LoginModuleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.docx = Path(self.temp.name) / "测试系统说明.docx"

    def plain(self, *extra: ET._Element, reverse: bool = False) -> list[ET._Element]:
        block = [picture(), paragraph(LOGIN)] if reverse else [paragraph(LOGIN), picture()]
        return [paragraph("测试系统"), paragraph("使用说明书"), paragraph("一、软件介绍"), paragraph("这是介绍正文。"), *block, paragraph("二、操作介绍"), paragraph("操作章节说明。"), paragraph("2.1数据管理"), paragraph("原数据管理正文。"), *extra, paragraph("2.2查询"), paragraph("查询功能正文。"), paragraph("三、其他说明")]

    def test_plain_text_headings_shift_descendants_without_losing_content(self) -> None:
        nodes = self.plain(paragraph("2.1.1新增数据"), paragraph("输入后点击保存。"), reverse=True)
        write_docx(self.docx, nodes)
        original = self.docx.read_bytes()
        preview = login.normalize(self.docx, "测试系统")
        self.assertEqual(preview["status"], "READY")
        self.assertEqual(self.docx.read_bytes(), original)
        result = login.normalize(self.docx, "测试系统", apply=True)
        self.assertEqual(result["status"], "APPLIED")
        after = body_nodes(self.docx)
        texts = [login.text(p) for p in after]
        position = texts.index("2.1登录页")
        self.assertEqual(texts[position + 1], LOGIN)
        self.assertTrue(after[position + 2].xpath(".//w:drawing", namespaces=NS))
        self.assertEqual(texts[position + 3], "2.2数据管理")
        self.assertIn("2.2.1新增数据", texts)
        self.assertIn("2.3查询", texts)
        self.assertIn("三、其他说明", texts)
        self.assertEqual(texts.count(LOGIN), 1)
        self.assertEqual(sum(bool(p.xpath(".//w:drawing", namespaces=NS)) for p in after), 1)
        with zipfile.ZipFile(self.docx) as archive:
            self.assertEqual(archive.read("word/media/image1.png"), b"unchanged original image bytes")
            self.assertEqual(archive.read("word/header1.xml"), b"unchanged header")
        first_output = self.docx.read_bytes()
        self.assertEqual(login.normalize(self.docx, "测试系统", apply=True)["status"], "ALREADY_COMPLIANT")
        self.assertEqual(self.docx.read_bytes(), first_output)

    def test_number_replacement_preserves_split_runs_and_heading_font(self) -> None:
        nodes = self.plain()
        heading = next(p for p in nodes if login.text(p) == "2.1数据管理")
        heading.remove(heading[0])
        for value in ("2.", "1", "数据管理"):
            run = ET.SubElement(heading, W + "r")
            props = ET.SubElement(run, W + "rPr")
            ET.SubElement(props, W + "b")
            ET.SubElement(run, W + "t").text = value
        write_docx(self.docx, nodes)
        self.assertEqual(login.normalize(self.docx, "测试系统", apply=True)["status"], "APPLIED")
        output = next(p for p in body_nodes(self.docx) if login.text(p) == "2.2数据管理")
        self.assertEqual(len(output.findall("w:r/w:rPr/w:b", NS)), 3)

    def test_automatic_numbering_keeps_list_and_descendant_definitions(self) -> None:
        nodes = [paragraph("简介", num=("1", 0)), paragraph(LOGIN), picture(), paragraph("功能介绍", num=("1", 0)), paragraph("数据管理", num=("1", 1)), paragraph("新增", num=("1", 2)), paragraph("查询", num=("1", 1))]
        definition = numbered_definition()
        write_docx(self.docx, nodes, numbering=definition)
        result = login.normalize(self.docx, "测试系统", apply=True)
        self.assertEqual(result["status"], "APPLIED")
        after = body_nodes(self.docx)
        titles = [p for p in after if login.text(p) in {"登录页", "数据管理", "查询"}]
        for p in titles:
            self.assertEqual(login._value(p.find("w:pPr", NS), "w:numPr/w:numId"), "1")
            self.assertEqual(login._value(p.find("w:pPr", NS), "w:numPr/w:ilvl"), "1")
        with zipfile.ZipFile(self.docx) as archive:
            self.assertEqual(archive.read("word/numbering.xml"), definition)

    def test_auto_heading_number_chain_cannot_include_body_step_list(self) -> None:
        nodes = [paragraph("简介", num=("1", 0)), paragraph(LOGIN), picture(), paragraph("功能介绍", num=("1", 0)), paragraph("数据管理", num=("1", 1)), paragraph("请填写所有字段，点击按钮后即可保存。", num=("1", 1)), paragraph("查询", num=("1", 1))]
        write_docx(self.docx, nodes, numbering=numbered_definition())
        before = self.docx.read_bytes()
        result = login.normalize(self.docx, "测试系统", apply=True)
        self.assertEqual(result["status"], "NEEDS_MANUAL_REVIEW")
        self.assertIn("普通列表", result["message"])
        self.assertEqual(before, self.docx.read_bytes())

    def test_sole_function_chapter_is_allowed_but_sole_overview_is_not(self) -> None:
        for title, status in (("平台功能", "READY"), ("基本介绍", "NEEDS_MANUAL_REVIEW")):
            with self.subTest(title=title):
                write_docx(self.docx, [paragraph(LOGIN), picture(), paragraph(title, outline=0), paragraph("数据管理", outline=1)])
                self.assertEqual(login.normalize(self.docx, "测试系统")["status"], status)

    def test_plain_cover_label_is_excluded_before_part_headings(self) -> None:
        write_docx(self.docx, [paragraph("测试系统", outline=0), paragraph("操作说明"), paragraph("第一部分：系统介绍", outline=0), paragraph("1.1组成"), paragraph(LOGIN), picture(), paragraph("第二部分：操作说明", outline=0), paragraph("2.1菜单", outline=1)])
        result = login.normalize(self.docx, "测试系统")
        self.assertEqual(result["status"], "READY")
        self.assertEqual(result["target_chapter"], "操作说明")
        self.assertEqual(result["template_heading"], "2.1菜单")

    def test_two_neighbor_images_are_reported_without_changes(self) -> None:
        nodes = self.plain()
        i = next(i for i, p in enumerate(nodes) if login.text(p) == LOGIN)
        nodes.insert(i, picture())
        write_docx(self.docx, nodes)
        before = self.docx.read_bytes()
        result = login.normalize(self.docx, "测试系统", apply=True)
        self.assertEqual(result["status"], "NEEDS_MANUAL_REVIEW")
        self.assertIn("2 张", result["message"])
        self.assertEqual(before, self.docx.read_bytes())

    def test_login_bookmark_moves_intact_but_cross_paragraph_bookmark_blocks(self) -> None:
        for cross_paragraph in (False, True):
            with self.subTest(cross_paragraph=cross_paragraph):
                nodes = self.plain()
                p = next(p for p in nodes if login.text(p) == LOGIN)
                start, end = ET.Element(W + "bookmarkStart"), ET.Element(W + "bookmarkEnd")
                start.set(W + "id", "91")
                start.set(W + "name", "doc_refurbish_login_page_v1")
                end.set(W + "id", "91")
                p.insert(0, start)
                (nodes[-1] if cross_paragraph else p).append(end)
                write_docx(self.docx, nodes)
                result = login.normalize(self.docx, "测试系统", apply=True)
                self.assertEqual(result["status"], "NEEDS_MANUAL_REVIEW" if cross_paragraph else "APPLIED")
                if not cross_paragraph:
                    moved = next(p for p in body_nodes(self.docx) if login.text(p) == LOGIN)
                    self.assertEqual(moved.find("w:bookmarkStart", NS).get(W + "name"), "doc_refurbish_login_page_v1")

    def test_existing_plain_login_heading_is_unchanged(self) -> None:
        write_docx(self.docx, [paragraph("一、概述"), paragraph("二、操作说明"), paragraph("2.1、登录"), picture(), paragraph(LOGIN), paragraph("2.2、查询")])
        before = self.docx.read_bytes()
        self.assertEqual(login.normalize(self.docx, "测试系统", apply=True)["status"], "ALREADY_COMPLIANT")
        self.assertEqual(self.docx.read_bytes(), before)

    def test_existing_login_heading_variants_never_add_another_login_section(self) -> None:
        for title in ("登录页面", "用户登录页面", "登录模块说明", "登录界面操作", "登录页（用户身份认证）"):
            with self.subTest(title=title):
                write_docx(self.docx, [paragraph("一、概述"), paragraph("二、操作说明"), paragraph("2.1、" + title), paragraph(LOGIN), picture(), paragraph("2.2、查询")])
                before = self.docx.read_bytes()
                result = login.normalize(self.docx, "测试系统", apply=True)
                self.assertNotEqual(result["status"], "APPLIED", result)
                self.assertEqual(self.docx.read_bytes(), before)

    def test_multiple_or_uncertain_login_headings_require_review_without_insert(self) -> None:
        for titles in (("2.1登录页", "2.2用户登录页面"), ("2.1登录与权限管理",)):
            with self.subTest(titles=titles):
                write_docx(self.docx, [paragraph("一、概述"), paragraph("二、操作说明"), *(paragraph(t) for t in titles), paragraph(LOGIN), picture(), paragraph("2.3查询")])
                before = self.docx.read_bytes()
                self.assertEqual(login.normalize(self.docx, "测试系统", apply=True)["status"], "NEEDS_MANUAL_REVIEW")
                self.assertEqual(self.docx.read_bytes(), before)

    def test_existing_login_section_with_two_pictures_is_not_duplicated(self) -> None:
        write_docx(self.docx, [paragraph("一、概述"), paragraph("二、操作说明"), paragraph("2.1登录页面"), picture(), paragraph(LOGIN), picture(), paragraph("2.2查询")])
        before = self.docx.read_bytes()
        self.assertEqual(login.normalize(self.docx, "测试系统", apply=True)["status"], "ALREADY_COMPLIANT")
        self.assertEqual(self.docx.read_bytes(), before)

    def test_existing_login_heading_with_extra_intro_never_creates_duplicate(self) -> None:
        write_docx(self.docx, [paragraph("一、概述"), paragraph("二、操作说明"), paragraph("2.1、登录"), paragraph("先打开软件。"), paragraph(LOGIN), picture(), paragraph("2.2、查询")])
        before = self.docx.read_bytes()
        result = login.normalize(self.docx, "测试系统", apply=True)
        self.assertEqual(result["status"], "ALREADY_COMPLIANT")
        self.assertEqual(self.docx.read_bytes(), before)

    def test_complex_login_and_broken_number_chain_leave_original_bytes(self) -> None:
        for case in ("table", "heading", "gap", "missing_image"):
            with self.subTest(case=case):
                nodes = self.plain()
                i = next(i for i, p in enumerate(nodes) if login.text(p) == LOGIN)
                if case == "table":
                    table = ET.Element(W + "tbl")
                    ET.SubElement(ET.SubElement(table, W + "tr"), W + "tc").append(nodes[i])
                    nodes[i] = table
                elif case == "heading":
                    nodes[i] = paragraph("一、" + LOGIN)
                elif case == "gap":
                    nodes[-3] = paragraph("2.5查询")
                else:
                    nodes.pop(i + 1)
                write_docx(self.docx, nodes)
                before = self.docx.read_bytes()
                self.assertEqual(login.normalize(self.docx, "测试系统", apply=True)["status"], "NEEDS_MANUAL_REVIEW")
                self.assertEqual(self.docx.read_bytes(), before)

    def test_preview_reads_structure_only_and_does_not_load_images(self) -> None:
        write_docx(self.docx, self.plain())
        read = zipfile.ZipFile.read
        names = []
        def tracked(archive, name, *args, **kwargs):
            names.append(name)
            return read(archive, name, *args, **kwargs)
        with mock.patch.object(zipfile.ZipFile, "read", tracked):
            self.assertEqual(login.normalize(self.docx, "测试系统")["status"], "READY")
        self.assertFalse(any("media/" in str(name) for name in names))

    def test_atomic_replace_failure_keeps_source(self) -> None:
        write_docx(self.docx, self.plain())
        before = self.docx.read_bytes()
        with mock.patch.object(login.os, "replace", side_effect=PermissionError("file busy")):
            self.assertEqual(login.normalize(self.docx, "测试系统", apply=True)["status"], "FAILED")
        self.assertEqual(self.docx.read_bytes(), before)
        self.assertEqual(list(self.docx.parent.glob("*.tmp")), [])

    def test_review_result_does_not_undo_completed_word_repairs(self) -> None:
        package = {"name": "测试系统", "folder": "", "login_module": {"status": "NEEDS_MANUAL_REVIEW", "message": "两张相邻图片"}}
        report = {"packages": [{"name": "测试系统", "folder": "", "status": "OK", "actions_applied": ["NORMALIZE_HEADER"], "failures": []}]}
        process_batch._attach_login_module_results(self.docx.parent, [package], report)
        result = report["packages"][0]
        self.assertEqual(result["status"], "NEEDS_MANUAL_REVIEW")
        self.assertEqual(result["actions_applied"], ["NORMALIZE_HEADER"])
        self.assertIn("登录模块", result["failures"][0])

    def test_snapshot_restore_never_reports_login_change_as_applied(self) -> None:
        package = {"name": "测试系统", "folder": "", "login_module": {"status": "APPLIED", "changed": True}}
        report = {"packages": [{"name": "测试系统", "folder": "", "status": "ROLLED_BACK", "snapshot_restored": True, "failures": []}]}
        process_batch._attach_login_module_results(self.docx.parent, [package], report)
        result = report["packages"][0]["login_module"]
        self.assertEqual(result["status"], "NEEDS_MANUAL_REVIEW")
        self.assertFalse(result["changed"])

    def test_saved_login_requires_readable_pdf_with_matching_word_page_count(self) -> None:
        from pypdf import PdfWriter

        write_docx(self.docx, self.plain())
        applied = login.normalize(self.docx, "测试系统", apply=True)
        self.assertEqual(applied["status"], "APPLIED")
        applied.pop("typography", None)  # This test isolates PDF acceptance.
        package = {"name": "测试系统", "folder": "", "docx": self.docx.name, "login_module": applied}
        for pdf_pages, expected in ((1, "APPLIED"), (2, "FAILED"), (0, "FAILED")):
            with self.subTest(pdf_pages=pdf_pages):
                if pdf_pages:
                    writer = PdfWriter()
                    for _ in range(pdf_pages):
                        writer.add_blank_page(width=100, height=100)
                    with self.docx.with_suffix(".pdf").open("wb") as stream:
                        writer.write(stream)
                else:
                    self.docx.with_suffix(".pdf").write_bytes(b"broken pdf")
                report = {"packages": [{"name": "测试系统", "folder": "", "status": "OK", "word_pages": 1, "failures": []}]}
                process_batch._attach_login_module_results(self.docx.parent, [package], report)
                self.assertEqual(report["packages"][0]["login_module"]["status"], expected)


if __name__ == "__main__":
    unittest.main()

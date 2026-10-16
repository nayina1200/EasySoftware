from __future__ import annotations

import copy
import unittest
from test_manual_login_module import paragraph, picture, write_docx, body_nodes, LOGIN
import manual_login_module as login
import manual_login_typography as typography
from lxml import etree as ET

W, NS = login.W, login.NS


def format_run(p, font="宋体", size="24", bold="0"):
    for run in p.findall(".//w:r", NS):
        props = run.find("w:rPr", NS)
        if props is None:
            props = ET.SubElement(run, W + "rPr")
        for tag in ("rFonts", "sz", "b"):
            for old in props.findall("w:" + tag, NS):
                props.remove(old)
        fonts = ET.SubElement(props, W + "rFonts")
        for slot in typography.SLOTS:
            fonts.set(W + slot, font)
        ET.SubElement(props, W + "sz").set(W + "val", size)
        ET.SubElement(props, W + "b").set(W + "val", bold)
    return p


class LoginTypographyTests(unittest.TestCase):
    def make_manual(self, *, existing=True):
        instruction = format_run(paragraph(LOGIN), "黑体", "48", "1")
        nodes = [paragraph("一、系统介绍"), format_run(paragraph("这是第一段用于确定正文标准格式的普通说明文字。")), paragraph("二、功能介绍"), paragraph("2.1登录页" if existing else "2.1数据管理"), instruction, picture(), paragraph("2.2查询"), format_run(paragraph("这是第二段用于确定正文标准格式的独立说明文字。")), format_run(paragraph("这是第三段用于确定正文标准格式的独立说明文字。"))]
        return nodes, instruction

    def test_existing_section_changes_only_instruction_runs_and_is_idempotent(self):
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "测试.docx"
            nodes, instruction = self.make_manual()
            write_docx(path, nodes)
            original = [ET.tostring(n) for n in body_nodes(path)[:-1]]
            preview = login.normalize(path, "测试系统")
            self.assertEqual(preview["status"], "READY")
            self.assertFalse(preview["structure_changed"])
            self.assertEqual(preview["typography"]["reference_font"], "宋体")
            result = login.normalize(path, "测试系统", apply=True)
            self.assertEqual(result["status"], "APPLIED")
            output = body_nodes(path)[:-1]
            for i, n in enumerate(output):
                if i != 4:
                    self.assertEqual(ET.tostring(n), original[i])
            self.assertEqual(login.text(output[4]), LOGIN)
            self.assertEqual(typography.Typography({}).signatures(output[4])[0][4:], ("24", "0", "0"))
            saved = path.read_bytes()
            self.assertEqual(login.normalize(path, "测试系统", apply=True)["status"], "ALREADY_COMPLIANT")
            self.assertEqual(path.read_bytes(), saved)

    def test_mixed_runs_are_aligned_but_color_and_bookmarks_survive(self):
        nodes, instruction = self.make_manual()
        run = copy.deepcopy(instruction.find("w:r", NS));run.find("w:t", NS).text = "补充说明"
        instruction.append(run)
        ET.SubElement(run.find("w:rPr", NS), W + "color").set(W + "val", "FF0000")
        root = ET.Element(W + "document");ET.SubElement(root, W + "body").extend(nodes)
        contents = {"word/document.xml": ET.tostring(root)}
        manual = login.Manual(contents, "测试")
        result = typography.plan(manual, contents, instruction)
        typography.apply(instruction, result)
        self.assertEqual(len(set(typography.Typography({}).signatures(instruction))), 1)
        self.assertEqual(run.find("w:rPr/w:color", NS).get(W + "val"), "FF0000")

    def test_tied_body_styles_require_manual_review(self):
        nodes, instruction = self.make_manual()
        nodes.pop()  # One reference of each style.
        format_run(nodes[1], "黑体")
        root = ET.Element(W + "document");ET.SubElement(root, W + "body").extend(nodes)
        contents = {"word/document.xml": ET.tostring(root)}
        result = typography.plan(login.Manual(contents, "测试"), contents, instruction)
        self.assertEqual(result["status"], "NEEDS_MANUAL_REVIEW")

    def test_inherited_default_theme_and_direct_override(self):
        styles = f'<w:styles xmlns:w="{W[1:-1]}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:lang w:eastAsia="zh-CN" w:bidi="ar-SA"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:b/></w:rPr></w:style><w:style w:styleId="Child"><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style></w:styles>'
        theme = f'<a:theme xmlns:a="{NS["a"]}"><a:themeElements><a:fontScheme><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hans" typeface="宋体"/><a:font script="Arab" typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'
        resolver = typography.Typography({"word/styles.xml": styles.encode(), "word/theme/theme1.xml": theme.encode()})
        p = paragraph(LOGIN);pp = ET.SubElement(p, W + "pPr");ET.SubElement(pp, W + "pStyle").set(W + "val", "Child")
        self.assertEqual(resolver.signatures(p), [("Calibri", "Calibri", "宋体", "Arial", "24", "0", "0")])
        format_run(p, "微软雅黑", "28")
        self.assertEqual(resolver.signatures(p), [("微软雅黑",)*4 + ("28", "0", "0")])

    def test_revision_is_reported_without_mutation(self):
        nodes, instruction = self.make_manual()
        ET.SubElement(instruction, W + "ins")
        before = ET.tostring(instruction)
        root = ET.Element(W + "document");ET.SubElement(root, W + "body").extend(nodes)
        contents = {"word/document.xml": ET.tostring(root)}
        self.assertEqual(typography.plan(login.Manual(contents, "测试"), contents, instruction)["status"], "NEEDS_MANUAL_REVIEW")
        self.assertEqual(ET.tostring(instruction), before)

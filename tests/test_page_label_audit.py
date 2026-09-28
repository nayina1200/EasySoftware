from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import page_label_audit as audit  # noqa: E402


W = audit.W
R = audit.R
MC = audit.MC


def _el(tag: str, **attrs: str) -> ET.Element:
    return ET.Element(f"{{{W}}}{tag}", {f"{{{W}}}{key}": value for key, value in attrs.items()})


def make_docx(path: Path, *, ambiguous: bool = False) -> None:
    document = ET.Element(f"{{{W}}}document")
    body = ET.SubElement(document, f"{{{W}}}body")
    section = ET.SubElement(body, f"{{{W}}}sectPr")
    ref = ET.SubElement(section, f"{{{W}}}headerReference")
    ref.set(f"{{{W}}}type", "default")
    ref.set(f"{{{R}}}id", "rId1")
    header = ET.Element(f"{{{W}}}hdr")
    textbox = ET.SubElement(header, f"{{{W}}}txbxContent")
    paragraph = ET.SubElement(textbox, f"{{{W}}}p")
    count = 2 if ambiguous else 1
    for index in range(count):
        run = ET.SubElement(paragraph, f"{{{W}}}r")
        text = ET.SubElement(run, f"{{{W}}}t")
        text.text = "第 " if index == 0 else " / "
        begin_run = ET.SubElement(paragraph, f"{{{W}}}r")
        begin = ET.SubElement(begin_run, f"{{{W}}}fldChar")
        begin.set(f"{{{W}}}fldCharType", "begin")
        code_run = ET.SubElement(paragraph, f"{{{W}}}r")
        instr = ET.SubElement(code_run, f"{{{W}}}instrText")
        instr.text = " PAGE "
        separate_run = ET.SubElement(paragraph, f"{{{W}}}r")
        separate = ET.SubElement(separate_run, f"{{{W}}}fldChar")
        separate.set(f"{{{W}}}fldCharType", "separate")
        cached_run = ET.SubElement(paragraph, f"{{{W}}}r")
        cached = ET.SubElement(cached_run, f"{{{W}}}t")
        cached.text = "1"
        end_run = ET.SubElement(paragraph, f"{{{W}}}r")
        end = ET.SubElement(end_run, f"{{{W}}}fldChar")
        end.set(f"{{{W}}}fldCharType", "end")
        suffix_run = ET.SubElement(paragraph, f"{{{W}}}r")
        suffix = ET.SubElement(suffix_run, f"{{{W}}}t")
        suffix.text = " 页" if index == 0 else " 页"
    rels = ET.Element(f"{{{audit.PKG_REL}}}Relationships")
    rel = ET.SubElement(rels, f"{{{audit.PKG_REL}}}Relationship")
    rel.set("Id", "rId1")
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
    rel.set("Target", "header1.xml")
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", ET.tostring(document))
        archive.writestr("word/header1.xml", ET.tostring(header))
        archive.writestr("word/_rels/document.xml.rels", ET.tostring(rels))


def make_pdf(path: Path, labels: list[str], *, x: float = 500, crop: tuple[float, float, float, float] | None = None) -> None:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    c = canvas.Canvas(str(path), pagesize=(612, 792))
    for label in labels:
        c.setFont("STSong-Light", 12)
        c.drawString(x, 760, label)
        c.drawString(72, 400, "正文文本")
        c.showPage()
    c.save()
    if crop:
        reader = PdfReader(str(path))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
            writer.pages[-1].cropbox.lower_left = (crop[0], crop[1])
            writer.pages[-1].cropbox.upper_right = (crop[2], crop[3])
        with path.open("wb") as output:
            writer.write(output)


class PageLabelAuditTests(unittest.TestCase):
    def test_extracts_active_textbox_page_field_template_and_preserves_literals(self):
        with tempfile.TemporaryDirectory() as directory:
            docx = Path(directory) / "sample.docx"
            make_docx(docx)
            result = audit.extract_page_label_templates(docx)
        self.assertEqual(result["status"], "ok")
        self.assertEqual(len(result["templates"]), 1)
        template = result["templates"][0]
        self.assertEqual((template["prefix"], template["suffix"]), ("第", "页"))
        self.assertTrue(template["in_textbox"])
        self.assertEqual(template["field_count"], 1)
        self.assertIn("header1.xml", template["part"])
        self.assertEqual(template["section_indices"], [1])

    def test_nested_textbox_paragraph_and_mc_fallback_are_not_double_counted(self):
        with tempfile.TemporaryDirectory() as directory:
            docx = Path(directory) / "nested.docx"
            make_docx(docx)
            with zipfile.ZipFile(docx, "r") as source:
                payloads = {name: source.read(name) for name in source.namelist()}
            header = ET.Element(f"{{{W}}}hdr")
            outer = ET.SubElement(header, f"{{{W}}}p")

            def append_page_field(paragraph: ET.Element, prefix: str, suffix: str) -> None:
                run = ET.SubElement(paragraph, f"{{{W}}}r")
                ET.SubElement(run, f"{{{W}}}t").text = prefix
                for kind, tag, value in (("begin", "fldChar", None), (None, "instrText", " PAGE "),
                                         ("separate", "fldChar", None), (None, "t", "1"),
                                         ("end", "fldChar", None)):
                    run = ET.SubElement(paragraph, f"{{{W}}}r")
                    node = ET.SubElement(run, f"{{{W}}}{tag}")
                    if kind:
                        node.set(f"{{{W}}}fldCharType", kind)
                    if value:
                        node.text = value
                ET.SubElement(ET.SubElement(paragraph, f"{{{W}}}r"), f"{{{W}}}t").text = suffix

            append_page_field(outer, "Outer ", " label")
            alt = ET.SubElement(ET.SubElement(outer, f"{{{W}}}r"), f"{{{MC}}}AlternateContent")
            choice = ET.SubElement(alt, f"{{{MC}}}Choice")
            textbox = ET.SubElement(choice, f"{{{W}}}txbxContent")
            append_page_field(ET.SubElement(textbox, f"{{{W}}}p"), "第", "页")
            fallback = ET.SubElement(alt, f"{{{MC}}}Fallback")
            fallback_textbox = ET.SubElement(fallback, f"{{{W}}}txbxContent")
            append_page_field(ET.SubElement(fallback_textbox, f"{{{W}}}p"), "第", "页")
            payloads["word/header1.xml"] = ET.tostring(header)
            with zipfile.ZipFile(docx, "w") as output:
                for name, payload in payloads.items():
                    output.writestr(name, payload)

            result = audit.extract_page_label_templates(docx)
        self.assertEqual(result["status"], "ok")
        self.assertEqual(len(result["templates"]), 2)
        self.assertEqual([row["field_count"] for row in result["templates"]], [1, 1])
        self.assertEqual([row["section_indices"] for row in result["templates"]], [[1], [1]])
        self.assertEqual(result["templates"][0]["prefix"], "Outer")
        self.assertEqual(result["templates"][1]["prefix"], "第")

    def test_normal_pdf_page_labels_are_complete(self):
        with tempfile.TemporaryDirectory() as directory:
            docx, pdf = Path(directory) / "sample.docx", Path(directory) / "sample.pdf"
            make_docx(docx)
            make_pdf(pdf, ["第 一 页", "第 二 页", "第 三 页"])
            templates = audit.extract_page_label_templates(docx)
            result = audit.audit_pdf_page_labels(pdf, templates)
        self.assertEqual(result["status"], "complete")
        self.assertEqual([row["status"] for row in result["pages"]], ["complete"] * 3)
        self.assertEqual([row["field_page_number"] for row in result["pages"]], [1, 2, 3])
        self.assertTrue(all(row["bbox"] and row["crop_box"] and row["media_box"] for row in result["pages"]))

    def test_prefix_only_label_touching_crop_edge_is_confirmed_truncated(self):
        with tempfile.TemporaryDirectory() as directory:
            docx, pdf = Path(directory) / "sample.docx", Path(directory) / "bad.pdf"
            make_docx(docx)
            make_pdf(pdf, ["第"], x=609)
            result = audit.audit_docx_pdf_page_labels(docx, pdf)
        self.assertEqual(result["status"], "truncated")
        self.assertEqual(result["pages"][0]["status"], "truncated")
        self.assertIn("CropBox 右边界", " ".join(result["pages"][0]["evidence"]))

    def test_complete_page_label_wrapped_to_neighbor_line_is_truncated(self):
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.pdfgen import canvas

        with tempfile.TemporaryDirectory() as directory:
            docx, pdf = Path(directory) / "sample.docx", Path(directory) / "wrapped.pdf"
            make_docx(docx)
            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
            writer = canvas.Canvas(str(pdf), pagesize=(612, 792))
            writer.setFont("STSong-Light", 12)
            writer.drawString(469, 760, "第")
            writer.drawString(469, 748, "一 页")
            writer.showPage()
            writer.save()
            result = audit.audit_docx_pdf_page_labels(docx, pdf)
        self.assertEqual(result["status"], "truncated")
        self.assertEqual(result["pages"][0]["status"], "truncated")
        self.assertIn("拆分到相邻", " ".join(result["pages"][0]["evidence"]))

    def test_incomplete_label_without_crop_evidence_fails_closed(self):
        with tempfile.TemporaryDirectory() as directory:
            docx, pdf = Path(directory) / "sample.docx", Path(directory) / "partial.pdf"
            make_docx(docx)
            make_pdf(pdf, ["第"], x=500)
            result = audit.audit_docx_pdf_page_labels(docx, pdf)
        self.assertEqual(result["status"], "indeterminate")
        self.assertIn("CropBox 未证明裁切", " ".join(result["pages"][0]["issues"]))

    def test_crop_box_clipping_is_reported(self):
        with tempfile.TemporaryDirectory() as directory:
            docx, pdf = Path(directory) / "sample.docx", Path(directory) / "cropped.pdf"
            make_docx(docx)
            make_pdf(pdf, ["第 一 页"], x=580, crop=(0, 0, 582, 792))
            result = audit.audit_pdf_page_labels(pdf, audit.extract_page_label_templates(docx))
        self.assertIn(result["pages"][0]["status"], {"truncated", "indeterminate"})
        self.assertNotEqual(result["pages"][0]["crop_box"], result["pages"][0]["media_box"])

    def test_ambiguous_docx_template_is_indeterminate(self):
        with tempfile.TemporaryDirectory() as directory:
            docx, pdf = Path(directory) / "ambiguous.docx", Path(directory) / "sample.pdf"
            make_docx(docx, ambiguous=True)
            make_pdf(pdf, ["第 一 页"])
            templates = audit.extract_page_label_templates(docx)
            result = audit.audit_pdf_page_labels(pdf, templates)
        self.assertEqual(templates["status"], "indeterminate")
        self.assertEqual(result["status"], "indeterminate")

    def test_missing_pdfminer_reports_each_page_indeterminate(self):
        with tempfile.TemporaryDirectory() as directory:
            docx, pdf = Path(directory) / "sample.docx", Path(directory) / "sample.pdf"
            make_docx(docx)
            make_pdf(pdf, ["第 一 页", "第 二 页"])
            templates = audit.extract_page_label_templates(docx)
            with mock.patch.dict(sys.modules, {"pdfminer": None, "pdfminer.high_level": None}):
                result = audit.audit_pdf_page_labels(pdf, templates)
        self.assertEqual(result["status"], "indeterminate")
        self.assertEqual(len(result["pages"]), 2)
        self.assertEqual([page["status"] for page in result["pages"]], ["indeterminate"] * 2)


if __name__ == "__main__":
    unittest.main()

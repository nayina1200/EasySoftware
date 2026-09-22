from __future__ import annotations

import argparse
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
LABEL_RE = re.compile(r"^[【\[]?(使用说明书|使用说明|说明书|使用手册|用户手册|用户使用手册|用户操作手册|用户操作说明|操作说明书|操作说明|操作手册)[】\]]?$")


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def text_of(p) -> str:
    return re.sub(r"\s+", " ", "".join(p.xpath(".//w:t/text()", namespaces=NS))).strip()


def ensure_child(parent, local: str):
    child = parent.find(qn(local))
    if child is None:
        child = etree.SubElement(parent, qn(local))
    return child


def style_paragraph(p, half_points: str):
    ppr = p.find(qn("pPr"))
    if ppr is None:
        ppr = etree.Element(qn("pPr"))
        p.insert(0, ppr)
    numpr = ppr.find(qn("numPr"))
    if numpr is not None:
        ppr.remove(numpr)
    jc = ensure_child(ppr, "jc")
    jc.set(qn("val"), "center")
    for r in p.xpath(".//w:r[w:t]", namespaces=NS):
        rpr = r.find(qn("rPr"))
        if rpr is None:
            rpr = etree.Element(qn("rPr"))
            r.insert(0, rpr)
        fonts = ensure_child(rpr, "rFonts")
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(attr), "宋体")
        ensure_child(rpr, "b").set(qn("val"), "1")
        ensure_child(rpr, "bCs").set(qn("val"), "1")
        ensure_child(rpr, "sz").set(qn("val"), half_points)
        ensure_child(rpr, "szCs").set(qn("val"), half_points)


def patch_docx(path: Path, title: str) -> tuple[bool, str]:
    with zipfile.ZipFile(path, "r") as src:
        content = {info.filename: src.read(info.filename) for info in src.infolist()}
        infos = src.infolist()
    root = etree.fromstring(content["word/document.xml"])
    paras = root.xpath("/w:document/w:body/w:p", namespaces=NS)[:20]
    title_p = next((p for p in paras if text_of(p) == title), None)
    label_p = next((p for p in paras if LABEL_RE.fullmatch(text_of(p))), None)
    if title_p is None or label_p is None:
        return False, "cover block not found"
    sizes = title_p.xpath(".//w:rPr/w:sz/@w:val", namespaces=NS)
    half_points = sizes[0] if sizes else "52"
    style_paragraph(title_p, half_points)
    style_paragraph(label_p, half_points)
    content["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    fd, temp_name = tempfile.mkstemp(suffix=".docx", dir=path.parent)
    os.close(fd)
    Path(temp_name).unlink(missing_ok=True)
    try:
        with zipfile.ZipFile(temp_name, "w") as dst:
            for info in infos:
                dst.writestr(info, content[info.filename])
        Path(temp_name).replace(path)
    finally:
        try:
            Path(temp_name).unlink()
        except FileNotFoundError:
            pass
    return True, half_points


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("base", type=Path)
    args = parser.parse_args()
    failures = 0
    for folder in sorted(p for p in args.base.iterdir() if p.is_dir()):
        docxs = list(folder.glob("*说明.docx"))
        if not docxs:
            continue
        txts = list(folder.glob("*.txt"))
        title = txts[0].stem if txts else folder.name
        ok, detail = patch_docx(docxs[0], title)
        print(f"{title}\t{ok}\t{detail}")
        failures += int(not ok)
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()

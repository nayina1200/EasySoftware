from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image
from pypdf import PdfReader

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
LABEL_RE = re.compile(r"使用说明书|使用说明|说明书|使用手册|用户手册|用户使用手册|用户操作手册|用户操作说明|操作说明书|操作说明|操作手册")


def decode_txt(path: Path):
    raw = path.read_bytes()
    candidates = []
    if raw.startswith(b"\xef\xbb\xbf"):
        candidates.append(("utf-8-sig", "utf-8-sig"))
    elif raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        candidates.append(("utf-16", "utf-16"))
    candidates += [("utf-8", "utf-8"), ("gb18030", "gb18030")]
    for label, codec in candidates:
        try:
            return label, raw.decode(codec)
        except UnicodeDecodeError:
            pass
    return "unknown", raw.decode("utf-8", errors="replace")


def dhash(data: bytes) -> int | None:
    try:
        im = Image.open(io.BytesIO(data)).convert("L").resize((9, 8))
    except Exception:
        return None
    value = 0
    for y in range(8):
        for x in range(8):
            value = (value << 1) | int(im.getpixel((x, y)) > im.getpixel((x + 1, y)))
    return value


def inspect_docx(path: Path, title: str):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        root = etree.fromstring(z.read("word/document.xml"))
        body_text = "\n".join(root.xpath(".//w:t/text()", namespaces=NS))
        paras = [re.sub(r"\s+", " ", "".join(p.xpath(".//w:t/text()", namespaces=NS))).strip()
                 for p in root.xpath("/w:document/w:body/w:p", namespaces=NS)[:60]]
        headers, page_field = [], False
        for name in names:
            if re.fullmatch(r"word/header\d+\.xml", name):
                h = etree.fromstring(z.read(name))
                headers.append("".join(h.xpath(".//w:t/text()", namespaces=NS)))
                page_field |= any("PAGE" in x for x in h.xpath(".//w:instrText/text()", namespaces=NS))
        media = []
        for name in sorted(n for n in names if n.startswith("word/media/")):
            data = z.read(name)
            media.append({"name": name, "sha256": hashlib.sha256(data).hexdigest(), "dhash": dhash(data), "bytes": len(data)})
    substantial = [m for m in media if m["bytes"] >= 10_000 and m["dhash"] is not None]
    similar = []
    if substantial:
        first = substantial[0]
        for m in substantial[1:]:
            distance = (first["dhash"] ^ m["dhash"]).bit_count()
            if distance <= 6:
                similar.append({"first": first["name"], "candidate": m["name"], "distance": distance})
    return {
        "cover_title": title in paras,
        "cover_label": any(LABEL_RE.search(p) for p in paras),
        "header_title": any(title in h for h in headers),
        "header_version": any("V1.0" in h for h in headers),
        "header_page_field": page_field,
        "mentions_login": "登录" in body_text,
        "image_count": len(media),
        "first_image_similarity_candidates": similar,
        "body_text": body_text,
    }


def suspicious_txt(title: str, text: str):
    issues = []
    if "软件的技术特点选项：教育软件" in text and not any(
        x in title for x in ("教育", "英语", "学习", "教学", "培训", "数学", "代数", "助教", "辅导", "训练", "测评", "评估", "思维")
    ):
        issues.append("技术特点填为教育软件，与名称用途不符")
    if "软件的技术特点选项：大数据软件" in text and not any(x in title for x in ("数据", "分析", "大数据")):
        issues.append("技术特点填为大数据软件，与名称用途不符")
    return issues


def inspect_folder(folder: Path):
    txts, code_pdfs, manuals, manual_pdfs = (list(folder.glob("*.txt")), list(folder.glob("*代码.pdf")),
                                             list(folder.glob("*说明.docx")), list(folder.glob("*说明.pdf")))
    title = txts[0].stem if txts else folder.name
    result = {"name": title, "folder": str(folder), "clear_errors": [], "candidates": []}
    if not txts or not code_pdfs or not manuals or not manual_pdfs:
        result["clear_errors"].append("缺少TXT、代码PDF、说明DOCX或说明PDF中的至少一种")
        return result
    encoding, txt = decode_txt(txts[0])
    result["txt"] = {"encoding": encoding, "name_present": title in txt, "issues": suspicious_txt(title, txt)}
    result["clear_errors"] += [f"TXT：{x}" for x in result["txt"]["issues"]]
    reader = PdfReader(str(code_pdfs[0]))
    first, last = (reader.pages[0].extract_text() or ""), (reader.pages[-1].extract_text() or "")
    result["code_pdf"] = {
        "pages": len(reader.pages), "first_has_name": title in first, "last_has_name": title in last,
        "first_has_version": "V1.0" in first, "last_has_version": "V1.0" in last,
    }
    if not 70 <= len(reader.pages) <= 130:
        result["candidates"].append(f"代码PDF页数为{len(reader.pages)}，偏离约100页")
    doc = inspect_docx(manuals[0], title)
    result["manual"] = {k: v for k, v in doc.items() if k != "body_text"}
    for key, label in (("cover_title", "封面缺少软件名称"), ("cover_label", "封面缺少说明/手册标题"),
                       ("header_title", "页眉缺少软件名称"), ("header_page_field", "页眉缺少动态页码")):
        if not doc[key]:
            result["clear_errors"].append(f"说明DOCX：{label}")
    if not doc["mentions_login"]:
        result["candidates"].append("说明正文未检出“登录”文字，需确认是否缺少登录界面")
    if doc["first_image_similarity_candidates"]:
        result["candidates"].append("存在与首张大图高度相似的图片，仅需对这些候选做视觉复核")
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("base", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    rows = [inspect_folder(p) for p in sorted(args.base.iterdir()) if p.is_dir()]
    args.out.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"packages={len(rows)} clear_errors={sum(len(x['clear_errors']) for x in rows)} candidates={sum(len(x['candidates']) for x in rows)}")


if __name__ == "__main__":
    main()

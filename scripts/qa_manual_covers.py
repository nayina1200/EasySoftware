from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
_bundled_pdftoppm = Path(sys.executable).resolve().parents[1] / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe"
PDFTOPPM = _bundled_pdftoppm if _bundled_pdftoppm.exists() else Path(shutil.which("pdftoppm") or "pdftoppm")
LABEL_RE = re.compile(r"^[【\[]?(使用说明书|使用说明|说明书|使用手册|用户手册|用户使用手册|用户操作手册|用户操作说明|操作说明书|操作说明|操作手册)[】\]]?$")


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\r", "").replace("\n", "")).strip()


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text).casefold()


def paragraph_text(p) -> str:
    return clean("".join(p.xpath(".//w:t/text()", namespaces=NS)))


def run_props(p):
    runs = p.xpath(".//w:r[w:t]", namespaces=NS)
    fonts, sizes, bold = set(), set(), []
    for r in runs:
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            value = r.xpath(f"string(./w:rPr/w:rFonts/@w:{attr})", namespaces=NS)
            if value:
                fonts.add(value)
        size = r.xpath("string(./w:rPr/w:sz/@w:val)", namespaces=NS)
        if size:
            sizes.add(size)
        b = r.xpath("./w:rPr/w:b", namespaces=NS)
        if b:
            bold.append(b[0].get(f"{{{W}}}val", "1") not in {"0", "false", "off"})
    bold_value = False if any(v is False for v in bold) else (True if bold else None)
    return sorted(fonts), sorted(sizes), bold_value


def inspect_docx(path: Path, title: str):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    paras = root.xpath("/w:document/w:body/w:p", namespaces=NS)
    title_p = next((p for p in paras[:20] if paragraph_text(p) == title), None)
    label_p = next((p for p in paras[:20] if LABEL_RE.fullmatch(paragraph_text(p))), None)
    if title_p is None or label_p is None:
        return {"cover_found": False}
    tf, ts, tb = run_props(title_p)
    lf, ls, lb = run_props(label_p)
    return {
        "cover_found": True,
        "title_fonts": tf,
        "label_fonts": lf,
        "title_sizes": ts,
        "label_sizes": ls,
        "title_bold": tb,
        "label_bold": lb,
        "same_size": ts == ls and len(ts) == 1,
        "simsun": (not tf or any(f in {"宋体", "SimSun"} for f in tf)) and (not lf or any(f in {"宋体", "SimSun"} for f in lf)),
    }


def render_page(pdf: Path, page: int, out: Path) -> Path:
    out.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [str(PDFTOPPM), "-f", str(page), "-l", str(page), "-r", "90", "-singlefile", "-png", str(pdf), str(out.with_suffix(""))],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    return out


def make_montage(cards, path: Path, cols=2):
    page_w, page_h = 744, 1053
    label_h = 44
    rows = (len(cards) + cols - 1) // cols
    canvas = Image.new("RGB", (cols * page_w, rows * (page_h + label_h)), "#d9dde2")
    font_path = Path(r"C:\Windows\Fonts\msyh.ttc")
    font = ImageFont.truetype(str(font_path), 22) if font_path.exists() else ImageFont.load_default()
    draw = ImageDraw.Draw(canvas)
    for i, (title, image_path) in enumerate(cards):
        x = (i % cols) * page_w
        y = (i // cols) * (page_h + label_h)
        draw.text((x + 8, y + 8), title, fill="black", font=font)
        im = Image.open(image_path).convert("RGB")
        im.thumbnail((page_w - 10, page_h - 10))
        canvas.paste(im, (x + (page_w - im.width) // 2, y + label_h))
    canvas.save(path)


def main(base: Path, report: Path, montage: Path, duplicate_titles: set[str]):
    results, cards = [], []
    temp = montage.parent / "qa_cover_pages"
    for folder in sorted(p for p in base.iterdir() if p.is_dir()):
        docxs = list(folder.glob("*说明.docx"))
        pdfs = list(folder.glob("*说明.pdf"))
        txts = list(folder.glob("*.txt"))
        if not docxs or not pdfs:
            continue
        title = txts[0].stem if txts else folder.name
        dc = inspect_docx(docxs[0], title)
        reader = PdfReader(str(pdfs[0]))
        first_text = reader.pages[0].extract_text() or ""
        later = [(reader.pages[i].extract_text() or "") for i in range(1, min(4, len(reader.pages)))]
        manual_words = ("使用说明书", "使用说明", "说明书", "使用手册", "用户手册", "用户使用手册", "用户操作手册", "用户操作说明", "操作说明书", "操作说明", "操作手册")
        duplicate_later = any(compact(title) in compact(t) and any(word in t for word in manual_words) for t in later)
        dc.update({
            "name": title,
            "pages": len(reader.pages),
            "pdf_first_has_title": compact(title) in compact(first_text),
            "duplicate_later": duplicate_later if title in duplicate_titles else False,
        })
        dc["all_ok"] = all([
            dc.get("cover_found"), dc.get("same_size"), dc.get("simsun"),
            dc.get("title_bold") is not False, dc.get("label_bold") is not False, dc.get("pdf_first_has_title"),
            not dc.get("duplicate_later"),
        ])
        results.append(dc)
        image_path = temp / f"{len(cards):02d}.png"
        render_page(pdfs[0], 1, image_path)
        cards.append((title, image_path))
    report.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    make_montage(cards, montage)
    for row in results:
        print(row["name"], row["all_ok"], row["pages"], row.get("title_sizes"), row.get("title_fonts"))


if __name__ == "__main__":
    main(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]), set(sys.argv[4:]))

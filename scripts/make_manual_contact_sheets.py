from __future__ import annotations

import math
import re
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader


_bundled_pdftoppm = Path(sys.executable).resolve().parents[1] / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe"
PDFTOPPM = _bundled_pdftoppm if _bundled_pdftoppm.exists() else Path(shutil.which("pdftoppm") or "pdftoppm")


def font(size):
    for path in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf"):
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def safe(name):
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in name)


def page_num(path: Path):
    m = re.search(r"-(\d+)\.png$", path.name)
    return int(m.group(1)) if m else 0


def main(base: Path, outdir: Path):
    outdir.mkdir(parents=True, exist_ok=True)
    title_font = font(25)
    page_font = font(17)
    for folder in sorted(p for p in base.iterdir() if p.is_dir()):
        pdf = next(folder.glob("*说明.pdf"))
        count = len(PdfReader(str(pdf)).pages)
        temp = outdir / ("tmp_" + safe(folder.name))
        temp.mkdir(exist_ok=True)
        prefix = temp / "page"
        subprocess.run(
            [str(PDFTOPPM), "-r", "64", "-png", str(pdf), str(prefix)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        files = sorted(temp.glob("page-*.png"), key=page_num)
        thumbs = []
        for f in files:
            img = Image.open(f).convert("RGB")
            scale = min(1.0, 330 / img.width)
            img = img.resize((round(img.width * scale), round(img.height * scale)), Image.Resampling.LANCZOS)
            thumbs.append((page_num(f), img))
        cols = 4
        cell_w = 350
        cell_h = max(i.height for _, i in thumbs) + 34
        rows = math.ceil(len(thumbs) / cols)
        canvas = Image.new("RGB", (cols * cell_w, 54 + rows * cell_h), "#cfd4da")
        d = ImageDraw.Draw(canvas)
        d.text((12, 10), f"{folder.name} | 说明书全 {count} 页", fill="black", font=title_font)
        for k, (pno, img) in enumerate(thumbs):
            x = (k % cols) * cell_w + 10
            y = 54 + (k // cols) * cell_h + 28
            d.text((x, y - 25), f"第 {pno} 页", fill="black", font=page_font)
            canvas.paste(img, (x, y))
        out = outdir / ("manual_" + safe(folder.name) + ".jpg")
        canvas.save(out, quality=90, optimize=True)
        print(out)


if __name__ == "__main__":
    main(Path(sys.argv[1]), Path(sys.argv[2]))

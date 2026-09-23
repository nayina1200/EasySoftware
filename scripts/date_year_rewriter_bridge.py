from __future__ import annotations

import importlib.util
import io
import json
import os
import re
import sys
import zipfile
from calendar import monthrange
from datetime import date, datetime as BaseDateTime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")


YEAR_RE = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
FULL_DATE_RE = re.compile(
    r"(?<!\d)(?P<year>(?:19|20|21)\d{2})[-/.年]"
    r"(?P<month>1[0-2]|0?[1-9])[-/.月]"
    r"(?P<day>3[01]|[12]\d|0?[1-9])日?(?!\d)"
)
YEAR_MONTH_RE = re.compile(
    r"(?<!\d)(?P<year>(?:19|20|21)\d{2})[-/.年]"
    r"(?P<month>1[0-2]|0?[1-9])月?(?![-/.月\d])"
)
MONTH_YEAR_RE = re.compile(
    r"(?<!\d)(?P<month>十[一二]?|[一二三四五六七八九]|1[0-2]|0?[1-9])月\s*"
    r"(?P<year>(?:19|20|21)\d{2})(?!\d)"
)
CHINESE_MONTHS = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
                  "十": 10, "十一": 11, "十二": 12}


def date_year_spans(text: str, target_year: int) -> list[tuple[int, int]]:
    """Only years anchored to a valid calendar date or an explicit year-month field.

    A bare four-digit number can be an employee ID or chart tick. In particular,
    `04-19 16:06` must never be joined into a fictitious year `1916`.
    """
    spans: set[tuple[int, int]] = set()
    for match in FULL_DATE_RE.finditer(text):
        try:
            date(int(match.group("year")), int(match.group("month")), int(match.group("day")))
            date(target_year, int(match.group("month")), int(match.group("day")))
        except ValueError:
            continue
        if int(match.group("year")) < target_year:
            spans.add(match.span("year"))
    for pattern in (YEAR_MONTH_RE, MONTH_YEAR_RE):
        for match in pattern.finditer(text):
            month = match.group("month")
            month_number = CHINESE_MONTHS.get(month, int(month) if month.isdigit() else 0)
            if 1 <= month_number <= 12 and int(match.group("year")) < target_year:
                spans.add(match.span("year"))
    return sorted(spans)


def ocr_date_fields(result, target_year: int) -> list[tuple[str, tuple[int, int, int, int], tuple[int, int]]]:
    fields = []
    boxes = getattr(result, "boxes", None)
    texts = getattr(result, "txts", None)
    for polygon, value in zip(boxes if boxes is not None else [], texts if texts is not None else []):
        text_value = str(value)
        xs = [int(point[0]) for point in polygon]
        ys = [int(point[1]) for point in polygon]
        rect = min(xs), min(ys), max(xs) + 1, max(ys) + 1
        for span in date_year_spans(text_value, target_year):
            fields.append((text_value, rect, span))
    return fields


def original_glyph_color(module, source_bgr, rect: tuple[int, int, int, int], fallback):
    """Choose the original high-contrast glyph color on light or tinted UI fields."""
    x1, y1, x2, y2 = rect
    roi = source_bgr[y1:y2, x1:x2]
    if roi.size == 0:
        return fallback
    gray = module.cv2.cvtColor(roi, module.cv2.COLOR_BGR2GRAY)
    background = module.np.median(roi.reshape(-1, 3), axis=0)
    candidates = []
    for mask in (gray >= module.np.percentile(gray, 95), gray <= module.np.percentile(gray, 5)):
        pixels = roi[mask]
        if len(pixels) >= 3:
            color = module.np.median(pixels, axis=0)
            contrast = float(module.np.linalg.norm(color - background))
            candidates.append((contrast, color))
    if not candidates:
        return fallback
    color = max(candidates, key=lambda item: item[0])[1].astype(module.np.uint8)
    return int(color[2]), int(color[1]), int(color[0])


def restore_year_background(module, output, source_bgr, rect: tuple[int, int, int, int]):
    """Restore the UI field row by row; rectangular inpainting smears blue bars."""
    x1, y1, x2, y2 = rect
    image_width = source_bgr.shape[1]
    for y in range(y1, y2):
        roi = source_bgr[y, x1:x2]
        background = module.np.median(roi, axis=0)
        sides = [source_bgr[y, max(0, x1 - 12):max(0, x1 - 3)],
                 source_bgr[y, min(image_width, x2 + 3):min(image_width, x2 + 12)]]
        candidates = [module.np.median(side, axis=0) for side in sides if side.size]
        if candidates:
            color = min(candidates, key=lambda value: float(module.np.linalg.norm(value - background)))
            output[y, x1:x2] = color.astype(module.np.uint8)
    return output


def replace_year_fields(module, engine, source_bgr, fields, target_year: int):
    """Replace only the OCR year glyphs; retain all month/day/time pixels."""
    output = source_bgr.copy()
    notes = []
    applied_fields = []
    for value, (row_x1, y1, row_x2, y2), (start, end) in fields:
        source_year = value[start:end]
        if source_year == str(target_year):
            continue
        measure = module.Image.new("RGB", (2, 2), (255, 255, 255))
        measure_draw = module.ImageDraw.Draw(measure)
        font = engine._fit_font(measure_draw, value, row_x2 - row_x1, y2 - y1)
        total = max(1.0, float(measure_draw.textlength(value, font=font)))
        left = float(measure_draw.textlength(value[:start], font=font))
        right = float(measure_draw.textlength(value[:end], font=font))
        x1 = row_x1 + round((row_x2 - row_x1) * left / total)
        x2 = row_x1 + round((row_x2 - row_x1) * right / total)
        x1 = max(0, min(output.shape[1] - 1, x1))
        x2 = max(x1 + 1, min(output.shape[1], x2))
        before_field = output[y1:y2, x1:x2].copy()
        _, fallback_color, _ = engine._background_and_text_colors(source_bgr, (x1, y1, x2, y2))
        text_color = original_glyph_color(module, source_bgr, (x1, y1, x2, y2), fallback_color)
        right_margin = 2 if value[end:end + 1] == "年" else 0
        output = restore_year_background(module, output, source_bgr,
                                         (max(0, x1 - 1), y1, min(output.shape[1], x2 + right_margin), y2))
        canvas = module.Image.fromarray(module.cv2.cvtColor(output, module.cv2.COLOR_BGR2RGB))
        draw = module.ImageDraw.Draw(canvas)
        replacement = str(target_year)
        replacement_font = engine._fit_font(draw, replacement, x2 - x1, y2 - y1)
        text_box = draw.textbbox((0, 0), replacement, font=replacement_font)
        draw_y = y1 + max(0, (y2 - y1 - (text_box[3] - text_box[1])) // 2) - text_box[1]
        draw.text((x1, draw_y), replacement, font=replacement_font, fill=text_color)
        output = module.cv2.cvtColor(module.np.asarray(canvas), module.cv2.COLOR_RGB2BGR)
        if not module.np.array_equal(before_field, output[y1:y2, x1:x2]):
            notes.append(f"{source_year} → {target_year}（原图日期字段年份）")
            applied_fields.append((value, (start, end)))
    return output, notes, applied_fields


def load_engine(tool_root: Path, target_year: int):
    internal = tool_root / "_internal"
    if hasattr(os, "add_dll_directory"):
        for directory, _children, files in os.walk(internal):
            if any(name.lower().endswith(".dll") for name in files):
                try:
                    os.add_dll_directory(directory)
                except OSError:
                    pass
    sys.path.insert(0, str(internal))
    sys._MEIPASS = str(internal)  # type: ignore[attr-defined]
    spec = importlib.util.spec_from_file_location("easysoftware_date_tool", tool_root / "app.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("无法加载图片日期工具")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)

    class FixedDateTime(BaseDateTime):
        @classmethod
        def now(cls, tz=None):
            current = BaseDateTime.now(tz)
            return cls(target_year, current.month, min(current.day, monthrange(target_year, current.month)[1]), tzinfo=current.tzinfo)

    module.datetime = FixedDateTime

    def same_month_day(_cls, _now, source_month, source_day):
        month = max(1, min(12, int(source_month)))
        day = max(1, min(int(source_day), monthrange(target_year, month)[1]))
        return module.date(target_year, month, day)

    module.DateStampProcessor._target_date_in_replacement_window = classmethod(same_month_day)
    return module, module.DateStampProcessor()


def encode_like_source(module, rgb_array, suffix: str) -> bytes:
    output = io.BytesIO()
    image = module.Image.fromarray(rgb_array)
    suffix = suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        image.convert("RGB").save(output, format="JPEG", quality=95, subsampling=0)
    elif suffix in {".tif", ".tiff"}:
        image.save(output, format="TIFF")
    elif suffix == ".bmp":
        image.save(output, format="BMP")
    elif suffix == ".webp":
        image.save(output, format="WEBP", quality=95)
    else:
        image.save(output, format="PNG")
    return output.getvalue()


def rewrite_docx(tool_root: Path, docx: Path, target_year: int) -> dict:
    module, engine = load_engine(tool_root, target_year)
    temporary = docx.with_suffix(docx.suffix + ".date-year.tmp")
    temporary.unlink(missing_ok=True)
    changed_images = 0
    detections_found = 0
    notes: list[str] = []
    changes: list[dict] = []
    with zipfile.ZipFile(docx, "r") as source, zipfile.ZipFile(temporary, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            suffix = Path(item.filename).suffix.lower()
            if item.filename.lower().startswith("word/media/") and suffix in IMAGE_SUFFIXES:
                opened = module.Image.open(io.BytesIO(data))
                rgb = module.np.asarray(module.ImageOps.exif_transpose(opened).convert("RGB"), dtype=module.np.uint8)[:, :, :3]
                bgr = module.cv2.cvtColor(rgb, module.cv2.COLOR_RGB2BGR)
                # Use the original full-image OCR boxes as the sole evidence.
                # Enhanced crop OCR can join adjacent date and time glyphs into
                # a false standalone year (for example 04-19 16:06 -> 1916).
                with engine.ocr_lock:
                    quick_result = engine.ocr(bgr)
                fields = ocr_date_fields(quick_result, target_year)
                if not fields:
                    target.writestr(item, data)
                    continue
                detections_found += len(fields)
                output_bgr, image_notes, applied_fields = replace_year_fields(module, engine, bgr, fields, target_year)
                output_rgb = module.cv2.cvtColor(output_bgr, module.cv2.COLOR_BGR2RGB)
                if not module.np.array_equal(rgb, output_rgb):
                    data = encode_like_source(module, output_rgb, suffix)
                    changed_images += 1
                    notes.extend(map(str, image_notes))
                    changes.extend({
                        "media_path": item.filename,
                        "ocr_text": value,
                        "source_year": value[span[0]:span[1]],
                        "target_year": str(target_year),
                        "changed_image": True,
                    } for value, span in applied_fields)
            target.writestr(item, data)
    os.replace(temporary, docx)
    return {
        "status": "CHANGED" if changed_images else "UNCHANGED",
        "changed": bool(changed_images),
        "changed_images": changed_images,
        "detections": detections_found,
        "target_year": target_year,
        "notes": notes,
        "changes": changes,
    }


def main() -> int:
    try:
        tool_root = Path(sys.argv[1]).resolve()
        docx = Path(sys.argv[2]).resolve()
        target_year = int(sys.argv[3])
        result = rewrite_docx(tool_root, docx, target_year)
        print(json.dumps(result, ensure_ascii=False))
        return 0
    except Exception as exc:
        if len(sys.argv) > 2:
            candidate = Path(sys.argv[2]).resolve()
            candidate.with_suffix(candidate.suffix + ".date-year.tmp").unlink(missing_ok=True)
        print(json.dumps({"status": "FAILED", "changed": False, "error": str(exc)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import importlib.util
import io
import json
import os
import re
import sys
import time
import zipfile
from calendar import monthrange
from datetime import date, datetime as BaseDateTime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")


YEAR_RE = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
FULL_DATE_RE = re.compile(
    r"(?<![A-Za-z0-9])(?P<year>(?:19|20|21)\d{2})"
    r"(?:\s*[-/.年]\s*|\s*['’‘]\s*|\s+)"
    r"(?P<month>1[0-2]|0?[1-9])\s*[-/.月]\s*"
    r"(?P<day>3[01]|[12]\d|0?[1-9])日?(?!\d)"
)
YEAR_MONTH_RE = re.compile(
    r"(?<![A-Za-z0-9])(?P<year>(?:19|20|21)\d{2})"
    r"(?:\s*[-/.年]\s*|\s*['’‘]\s*|\s+)"
    r"(?P<month>1[0-2]|0?[1-9])月?(?![-/.月A-Za-z0-9])"
)
MONTH_YEAR_RE = re.compile(
    r"(?<![A-Za-z0-9])(?P<month>十[一二]?|[一二三四五六七八九]|1[0-2]|0?[1-9])月\s*"
    r"(?P<year>(?:19|20|21)\d{2})(?![A-Za-z0-9])"
)
CHINESE_MONTHS = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
                  "十": 10, "十一": 11, "十二": 12}
SMALL_TEXT_HEIGHT = 24
OCR_DIGIT_CONFUSIONS = str.maketrans({"O": "0", "o": "0", "Q": "0", "D": "0", "I": "1", "l": "1", "|": "1", "S": "5", "B": "8", "Z": "2"})
COMPANY_TRAILING_PUNCTUATION = " \t,，.。;；:：!！?？、)]}）】》」』”’\"'"


def date_year_spans(
    text: str, target_year: int, *, include_target_year: bool = False
) -> list[tuple[int, int]]:
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
        year = int(match.group("year"))
        if year < target_year or (include_target_year and year == target_year):
            spans.add(match.span("year"))
    for pattern in (YEAR_MONTH_RE, MONTH_YEAR_RE):
        for match in pattern.finditer(text):
            month = match.group("month")
            month_number = CHINESE_MONTHS.get(month, int(month) if month.isdigit() else 0)
            year = int(match.group("year"))
            if 1 <= month_number <= 12 and (year < target_year or (include_target_year and year == target_year)):
                spans.add(match.span("year"))
    return sorted(spans)


def ocr_date_fields(
    result, target_year: int, *, include_target_year: bool = False
) -> list[tuple[str, tuple[int, int, int, int], tuple[int, int]]]:
    fields = []
    boxes = getattr(result, "boxes", None)
    texts = getattr(result, "txts", None)
    for polygon, value in zip(boxes if boxes is not None else [], texts if texts is not None else []):
        text_value = str(value)
        xs = [int(point[0]) for point in polygon]
        ys = [int(point[1]) for point in polygon]
        rect = min(xs), min(ys), max(xs) + 1, max(ys) + 1
        for span in date_year_spans(text_value, target_year, include_target_year=include_target_year):
            fields.append((text_value, rect, span))
    return fields


def _date_component_keys(
    text: str, target_year: int, *, include_target_year: bool = False
) -> set[tuple[int, int, int | None]]:
    """Return only calendar-valid date/year-month values, never bare year-like numbers."""
    keys: set[tuple[int, int, int | None]] = set()
    # Keep spacing: OCR commonly separates year and month as "2026 12/30" or
    # emits a quote before the month. The date regex accepts those separators.
    normalized = text.translate(OCR_DIGIT_CONFUSIONS)
    for match in FULL_DATE_RE.finditer(normalized):
        year, month, day = int(match.group("year")), int(match.group("month")), int(match.group("day"))
        try:
            date(year, month, day)
            date(target_year, month, day)
        except ValueError:
            continue
        if year < target_year or (include_target_year and year == target_year):
            keys.add((year, month, day))
    for pattern in (YEAR_MONTH_RE, MONTH_YEAR_RE):
        for match in pattern.finditer(normalized):
            month_text = match.group("month")
            month = CHINESE_MONTHS.get(month_text, int(month_text) if month_text.isdigit() else 0)
            year = int(match.group("year"))
            if 1 <= month <= 12 and (year < target_year or (include_target_year and year == target_year)):
                keys.add((year, month, None))
    return keys


def _ocr_rows(result) -> list[tuple[str, tuple[int, int, int, int], float]]:
    rows = []
    boxes = getattr(result, "boxes", None)
    texts = getattr(result, "txts", None)
    scores = getattr(result, "scores", None)
    boxes = boxes if boxes is not None else []
    texts = texts if texts is not None else []
    scores = scores if scores is not None else []
    for index, (polygon, value) in enumerate(zip(boxes, texts)):
        try:
            xs = [int(point[0]) for point in polygon]
            ys = [int(point[1]) for point in polygon]
            rect = min(xs), min(ys), max(xs) + 1, max(ys) + 1
            score = float(scores[index]) if index < len(scores) else 0.0
        except (TypeError, ValueError, IndexError):
            continue
        rows.append((str(value), rect, score))
    return rows


def company_text_fields(result) -> list[tuple[str, tuple[int, int, int, int], float]]:
    """Return OCR boxes whose complete text line ends with the company suffix."""
    return [
        (text, rect, score)
        for text, rect, score in _ocr_rows(result)
        if text.strip().rstrip(COMPANY_TRAILING_PUNCTUATION).endswith("公司")
    ]


def remove_company_text_fields(module, source_bgr, fields):
    """Inpaint only OCR boxes for company-name lines, preserving the source image."""
    image_height, image_width = source_bgr.shape[:2]
    mask = module.np.zeros((image_height, image_width), dtype=module.np.uint8)
    removed = []
    for text, rect, score in fields:
        x1, y1, x2, y2 = map(int, rect)
        line_height = max(1, y2 - y1)
        padding = max(1, min(4, round(line_height * 0.08)))
        x1, y1 = max(0, x1 - padding), max(0, y1 - padding)
        x2, y2 = min(image_width, x2 + padding), min(image_height, y2 + padding)
        if x2 <= x1 or y2 <= y1:
            continue
        module.cv2.rectangle(mask, (x1, y1), (x2 - 1, y2 - 1), 255, thickness=-1)
        removed.append({
            "ocr_text": text,
            "confidence": round(float(score), 4),
            "bounding_box": [int(rect[0]), int(rect[1]), int(rect[2]), int(rect[3])],
        })
    if not removed:
        return source_bgr.copy(), []
    radius = max(3, min(12, round(float(module.np.median([
        max(1, int(rect[3]) - int(rect[1])) for _text, rect, _score in fields
    ])) * 0.45)))
    output = module.cv2.inpaint(source_bgr, mask, radius, module.cv2.INPAINT_TELEA)
    if module.np.array_equal(source_bgr, output):
        return source_bgr.copy(), []
    return output, removed


def _box_overlap_ratio(left: tuple[int, int, int, int], right: tuple[int, int, int, int]) -> float:
    x1, y1 = max(left[0], right[0]), max(left[1], right[1])
    x2, y2 = min(left[2], right[2]), min(left[3], right[3])
    intersection = max(0, x2 - x1) * max(0, y2 - y1)
    left_area = max(1, (left[2] - left[0]) * (left[3] - left[1]))
    right_area = max(1, (right[2] - right[0]) * (right[3] - right[1]))
    return intersection / min(left_area, right_area)


def corroborate_year_month_field(module, engine, source_bgr, text: str, rect, target_year: int, *, include_target_year: bool = False) -> bool:
    """Confirm an ambiguous small year-month OCR field on its local image crop."""
    normalized = re.sub(r"\s+", "", text.translate(OCR_DIGIT_CONFUSIONS))
    if FULL_DATE_RE.search(normalized):
        return True
    expected = {key for key in _date_component_keys(text, target_year, include_target_year=include_target_year) if key[2] is None}
    if not expected:
        return True

    x1, y1, x2, y2 = map(int, rect)
    image_height = source_bgr.shape[0] if hasattr(source_bgr, "shape") else len(source_bgr)
    image_width = source_bgr.shape[1] if hasattr(source_bgr, "shape") else len(source_bgr[0])
    height, width = y2 - y1, x2 - x1
    if height <= 0 or width <= 0:
        return False
    pad_x = max(2, width // 10)
    pad_y = max(2, height // 4)
    bx1, by1 = max(0, x1 - pad_x), max(0, y1 - pad_y)
    bx2, by2 = min(image_width, x2 + pad_x), min(image_height, y2 + pad_y)
    if hasattr(source_bgr, "shape"):
        crop = source_bgr[by1:by2, bx1:bx2]
    else:
        crop = [row[bx1:bx2] for row in source_bgr[by1:by2]]
    enlarged = module.cv2.resize(
        crop, None, fx=2, fy=2, interpolation=module.cv2.INTER_CUBIC
    )
    try:
        with engine.ocr_lock:
            result = engine.ocr(enlarged)
    except Exception:
        return False
    confirmed = set()
    for value in getattr(result, "txts", None) or []:
        confirmed.update(
            key for key in _date_component_keys(str(value), target_year, include_target_year=include_target_year) if key[2] is None
        )
    return bool(expected.intersection(confirmed))


def adaptive_ocr_date_fields(module, engine, source_bgr, primary_result, target_year: int, *, include_target_year: bool = False):
    """Use the date engine's multi-scale OCR only around small date-like text.

    The fallback is deliberately anchored to an original OCR box and requires the
    source text itself to support the old year. This prevents crop OCR from
    inventing a year by joining a nearby date and time (for example 04-19 16:06).
    """
    primary_fields = ocr_date_fields(primary_result, target_year, include_target_year=include_target_year)
    verification_cache = {}
    rejected_year_month_candidates = 0

    def has_local_evidence(field) -> bool:
        nonlocal rejected_year_month_candidates
        value, rect, span = field
        key = value, rect, span
        if key not in verification_cache:
            verification_cache[key] = corroborate_year_month_field(
                module, engine, source_bgr, value, rect, target_year,
                include_target_year=include_target_year,
            )
            if not verification_cache[key]:
                rejected_year_month_candidates += 1
        return verification_cache[key]

    primary_fields = [field for field in primary_fields if has_local_evidence(field)]
    raw_rows = _ocr_rows(primary_result)
    small_rows = []
    for text, rect, _score in raw_rows:
        compact = re.sub(r"\s+", "", text)
        height = rect[3] - rect[1]
        date_hint = bool(re.search(
            r"(?:(?:19|20|21)[0-9OQDI|SBZ]{2}[-/.年]|"
            r"\d{1,2}[-/.]\d{1,2}(?:\D{0,5}\d{1,2}[:.]\d{2})?|"
            r"\d{1,2}月\d{1,2}日)",
            compact,
        ))
        if height <= SMALL_TEXT_HEIGHT and date_hint and _date_component_keys(text, target_year, include_target_year=include_target_year):
            small_rows.append((text, rect))
    if not small_rows:
        return primary_fields, {
            "used": bool(rejected_year_month_candidates),
            "reason": (
                "局部 OCR 未确认小字号年月候选，已保留原图"
                if rejected_year_month_candidates else "未发现需要增强的小字号日期候选"
            ),
            "fallback_candidates": 0,
            "rejected_year_month_candidates": rejected_year_month_candidates,
            "needs_skip": False,
        }

    try:
        candidates = engine.detect_ui_dates(source_bgr, confidence=0.72, document_fast_mode=True)
    except Exception as exc:
        return primary_fields, {"used": True, "reason": f"增强识别失败，保留首轮结果：{exc}", "fallback_candidates": 0}

    accepted = []
    for candidate in candidates or []:
        text = str(getattr(candidate, "text", ""))
        rect = tuple(map(int, getattr(candidate, "rect", (0, 0, 0, 0))))
        confidence = float(getattr(candidate, "confidence", 0.0) or 0.0)
        if len(rect) != 4 or rect[2] <= rect[0] or rect[3] <= rect[1]:
            continue
        spans = date_year_spans(text, target_year, include_target_year=include_target_year)
        if not spans or confidence < 0.35:
            continue
        candidate_keys = _date_component_keys(text, target_year, include_target_year=include_target_year)
        anchors = [
            raw_text for raw_text, anchor_rect in small_rows
            if _box_overlap_ratio(rect, anchor_rect) >= 0.15
        ]
        if not anchors:
            continue
        supported = False
        for span in spans:
            year = text[span[0]:span[1]]
            # OCR substitutions (O/0, I/1, etc.) are normalized for evidence only.
            for anchor in anchors:
                normalized_anchor = re.sub(r"\s+", "", anchor.translate(OCR_DIGIT_CONFUSIONS))
                if (re.search(rf"(?<!\d){re.escape(year)}(?!\d)", normalized_anchor)
                        and candidate_keys.intersection(_date_component_keys(anchor, target_year, include_target_year=include_target_year))):
                    supported = True
                    break
            if supported:
                if has_local_evidence((text, rect, span)):
                    accepted.append((text, rect, span))
                break
    if not accepted:
        return primary_fields, {
            "used": True, "reason": "小字号候选已自动重试，但未取得原图文字支持的有效日期；流程继续且不改该候选",
            "fallback_candidates": len(candidates or []),
            "rejected_year_month_candidates": rejected_year_month_candidates,
            "needs_skip": rejected_year_month_candidates == 0,
        }

    # Keep OCR text and geometry from the same recognition result. The enhanced
    # detector can return a wider box or include adjacent weekday text; replacing
    # the primary field with that box shifts the year when proportional placement
    # is calculated in replace_year_fields. Use enhanced detections only when no
    # primary field already covers the same date.
    merged = list(primary_fields)
    for enhanced in accepted:
        enhanced_keys = _date_component_keys(enhanced[0], target_year, include_target_year=include_target_year)
        has_primary_match = any(
            enhanced_keys.intersection(_date_component_keys(field[0], target_year, include_target_year=include_target_year))
            and _box_overlap_ratio(field[1], enhanced[1]) >= 0.15
            for field in primary_fields
        )
        if not has_primary_match:
            merged.append(enhanced)
    return merged, {
        "used": True, "reason": "已用小字号自适应多尺度 OCR 复核候选位置",
        "fallback_candidates": len(candidates or []), "accepted_candidates": len(accepted),
        "rejected_year_month_candidates": rejected_year_month_candidates,
    }


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


def _median_color(pixels) -> tuple[float, ...]:
    channels = list(zip(*(tuple(float(channel) for channel in pixel) for pixel in pixels)))
    return tuple(sorted(channel)[len(channel) // 2] if len(channel) % 2 else
                 (sorted(channel)[len(channel) // 2 - 1] + sorted(channel)[len(channel) // 2]) / 2
                 for channel in channels)


def _color_distance(left, right) -> float:
    return sum((float(a) - float(b)) ** 2 for a, b in zip(left, right)) ** 0.5


def restore_year_background(module, output, source_bgr, rect: tuple[int, int, int, int],
                            background_color, text_color):
    """Restore the UI field row by row; rectangular inpainting smears blue bars."""
    x1, y1, x2, y2 = rect
    image_width = source_bgr.shape[1] if hasattr(source_bgr, "shape") else len(source_bgr[0])
    background_reference = tuple(float(channel) for channel in background_color)
    foreground_reference = tuple(float(channel) for channel in reversed(text_color))
    for y in range(y1, y2):
        ranges = [
            (max(0, x1 - 12), max(0, x1 - 3)),
            (min(image_width, x2 + 3), min(image_width, x2 + 12)),
        ]
        candidates = []
        row = source_bgr[y]
        for start, end in ranges:
            if end <= start:
                continue
            pixels = list(row[start:end])
            if not pixels:
                continue
            color = _median_color(pixels)
            spread = _median_color([
                (_color_distance(pixel, color),) for pixel in pixels
            ])[0]
            background_distance = _color_distance(color, background_reference)
            foreground_distance = _color_distance(color, foreground_reference)
            # A flat stroke from a neighboring glyph can be more uniform than
            # the real background. Reject it when it is much closer to the ink.
            if background_distance > 40 and foreground_distance + 24 < background_distance:
                continue
            candidates.append((background_distance, spread, color))
        if not candidates:
            continue
        color = tuple(max(0, min(255, round(channel))) for channel in
                      min(candidates, key=lambda candidate: (candidate[0], candidate[1]))[2])
        try:
            output[y, x1:x2] = color
        except TypeError:
            output[y][x1:x2] = [list(color)] * (x2 - x1)
    return output


def text_ink_bounds(source_bgr, rect, background, text_color) -> tuple[int, int]:
    """Trim non-text margins (for example, a clock icon) from an OCR line box."""
    x1, y1, x2, y2 = rect
    roi = [row[x1:x2] for row in source_bgr[y1:y2]]
    height = len(roi)
    if not height:
        return x1, x2
    width = len(roi[0])
    if not width:
        return x1, x2

    background_bgr = tuple(float(channel) for channel in background)
    text_bgr = tuple(float(channel) for channel in reversed(text_color))
    ink_per_column = [0] * width
    for row in roi:
        for column, pixel in enumerate(row):
            pixel_bgr = tuple(float(channel) for channel in pixel)
            background_distance = sum(abs(value - background_bgr[index]) for index, value in enumerate(pixel_bgr))
            text_distance = sum(abs(value - text_bgr[index]) for index, value in enumerate(pixel_bgr))
            if background_distance > 60 and text_distance < 150:
                ink_per_column[column] += 1

    minimum_ink = max(2, round(height * 0.12))
    active_columns = [index for index, count in enumerate(ink_per_column) if count >= minimum_ink]
    if not active_columns:
        return x1, x2
    return x1 + active_columns[0], x1 + active_columns[-1] + 1


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
        background, fallback_color, _ = engine._background_and_text_colors(
            source_bgr, (row_x1, y1, row_x2, y2)
        )
        text_color = original_glyph_color(module, source_bgr, (row_x1, y1, row_x2, y2), fallback_color)
        text_x1, text_x2 = text_ink_bounds(
            source_bgr, (row_x1, y1, row_x2, y2), background, text_color
        )
        text_width = max(1, text_x2 - text_x1)
        x1 = text_x1 + round(text_width * left / total)
        x2 = text_x1 + round(text_width * right / total)
        x1 = max(0, min(output.shape[1] - 1, x1))
        x2 = max(x1 + 1, min(output.shape[1], x2))
        before_field = output[y1:y2, x1:x2].copy()
        # OCR boxes can trim anti-aliased edge pixels. Clear a small halo around
        # the estimated year glyphs, while keeping the following Chinese unit
        # inside its own untouched boundary.
        cleanup_padding = max(4, min(6, round((y2 - y1) * 0.06)))
        next_character = value[end:end + 1]
        right_padding = 0 if next_character in {"年", "月", "日"} else cleanup_padding
        output = restore_year_background(module, output, source_bgr,
                                         (max(0, x1 - cleanup_padding), y1,
                                          min(output.shape[1], x2 + right_padding), y2),
                                         background, text_color)
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


def _valid_2026_date_anchors(text: str) -> list[tuple[str, date]]:
    normalized = text.translate(OCR_DIGIT_CONFUSIONS)
    found: set[tuple[str, date]] = set()
    for match in FULL_DATE_RE.finditer(normalized):
        if match.group("year") != "2026":
            continue
        try:
            value = date(2026, int(match.group("month")), int(match.group("day")))
        except ValueError:
            continue
        found.add((value.isoformat(), value))
    for pattern in (YEAR_MONTH_RE, MONTH_YEAR_RE):
        for match in pattern.finditer(normalized):
            if match.group("year") != "2026":
                continue
            month_text = match.group("month")
            month = CHINESE_MONTHS.get(month_text, int(month_text) if month_text.isdigit() else 0)
            if not 1 <= month <= 12:
                continue
            # Month-only values use the first day of the month as their date anchor.
            anchor = date(2026, month, 1)
            found.add((anchor.isoformat(), anchor))
    return sorted(found, key=lambda item: item[1])


def _date_anchors_for_year_field(
    text: str, year_span: tuple[int, int], target_year: int
) -> list[tuple[str, date]]:
    """Project one OCR-confirmed date field to a candidate year."""
    normalized = text.translate(OCR_DIGIT_CONFUSIONS)
    found: set[tuple[str, date]] = set()
    for match in FULL_DATE_RE.finditer(normalized):
        if match.span("year") != year_span:
            continue
        month, day = int(match.group("month")), int(match.group("day"))
        try:
            date(int(match.group("year")), month, day)
            projected = date(target_year, month, day)
        except ValueError:
            continue
        found.add((projected.isoformat(), projected))

    for pattern in (YEAR_MONTH_RE, MONTH_YEAR_RE):
        for match in pattern.finditer(normalized):
            if match.span("year") != year_span:
                continue
            month_text = match.group("month")
            month = CHINESE_MONTHS.get(month_text, int(month_text) if month_text.isdigit() else 0)
            if not 1 <= month <= 12:
                continue
            try:
                date(int(match.group("year")), month, 1)
                projected = date(target_year, month, 1)
            except ValueError:
                continue
            found.add((projected.isoformat(), projected))
    return sorted(found, key=lambda item: item[1])


def _select_effective_target_year(
    target_year: int,
    *,
    observed_2026_dates: list[tuple[str, date]],
    projected_2026_dates: list[tuple[str, date]],
    today: date,
) -> tuple[int, str | None, str | None]:
    """Choose one document-wide target without creating future screenshot dates."""
    future_dates = sorted({
        value for value, parsed in (*observed_2026_dates, *projected_2026_dates)
        if parsed > today
    })
    eligible_dates = sorted({
        value for value, parsed in observed_2026_dates if parsed <= today
    })
    if future_dates:
        trigger_date = future_dates[0]
        rollback_reason = (
            f"将截图日期统一到 2026 会产生晚于本地日期 {today.isoformat()} 的日期：{trigger_date}；"
            "整份文档回退到 2025"
        )
        return 2025, trigger_date, rollback_reason
    if eligible_dates and target_year == 2025:
        return 2026, eligible_dates[0], None
    return target_year, None, None


def rewrite_docx(
    tool_root: Path,
    docx: Path,
    target_year: int,
    runtime=None,
    *,
    rewrite_dates: bool = True,
    remove_company_text: bool = False,
) -> dict:
    # Scan once against 2026 so both old dates and 2026 candidates are retained.
    # The final target is chosen only after all embedded images have been scanned.
    perf_trace = os.getenv("EASYSOFTWARE_PERF_TRACE") == "1"
    total_started = time.perf_counter() if perf_trace else None
    engine_load_started = time.perf_counter() if perf_trace and runtime is None else None
    module, engine = runtime if runtime is not None else load_engine(tool_root, 2026)
    performance = None
    if perf_trace:
        performance = {
            "engine_load_seconds": round(time.perf_counter() - engine_load_started, 4) if engine_load_started is not None else 0.0,
            "ocr_calls": 0,
            "primary_image_ocr_calls": 0,
            "ocr_seconds": 0.0,
            "scanned_image_count": 0,
            "zip_member_count": 0,
        }

        class TimedOcrEngine:
            def __init__(self, wrapped, metrics):
                self._wrapped = wrapped
                self._metrics = metrics

            def __getattr__(self, name):
                return getattr(self._wrapped, name)

            def ocr(self, image):
                started = time.perf_counter()
                try:
                    return self._wrapped.ocr(image)
                finally:
                    self._metrics["ocr_calls"] += 1
                    self._metrics["ocr_seconds"] += time.perf_counter() - started

        engine = TimedOcrEngine(engine, performance)
    temporary = docx.with_suffix(docx.suffix + ".date-year.tmp")
    temporary.unlink(missing_ok=True)
    changed_images = 0
    detections_found = 0
    company_changed_images = 0
    company_candidate_count = 0
    adaptive_images = 0
    notes: list[str] = []
    changes: list[dict] = []
    company_changes: list[dict] = []
    auto_skipped_images: list[dict[str, str]] = []
    scanned_images: list[dict] = []
    observed_2026_dates: list[tuple[str, date]] = []
    projected_2026_dates: list[tuple[str, date]] = []
    first_pass_started = time.perf_counter() if perf_trace else None
    with zipfile.ZipFile(docx, "r") as source, zipfile.ZipFile(temporary, "w") as target:
        if performance is not None:
            performance["zip_member_count"] = len(source.infolist())
        for item in source.infolist():
            data = source.read(item.filename)
            suffix = Path(item.filename).suffix.lower()
            if item.filename.lower().startswith("word/media/") and suffix in IMAGE_SUFFIXES:
                opened = module.Image.open(io.BytesIO(data))
                rgb = module.np.asarray(module.ImageOps.exif_transpose(opened).convert("RGB"), dtype=module.np.uint8)[:, :, :3]
                bgr = module.cv2.cvtColor(rgb, module.cv2.COLOR_RGB2BGR)
                # The original full-image OCR remains the source-evidence anchor.
                # Small date-like boxes may use the tool's enhanced detector to
                # refine recognition/geometry, but only when the original box
                # supports the same valid date. This rejects joins such as
                # 04-19 16:06 -> 1916.
                with engine.ocr_lock:
                    if performance is not None:
                        performance["primary_image_ocr_calls"] += 1
                    quick_result = engine.ocr(bgr)
                if rewrite_dates:
                    fields, adaptive = adaptive_ocr_date_fields(
                        module, engine, bgr, quick_result, 2026, include_target_year=True
                    )
                else:
                    fields, adaptive = [], {"used": False, "needs_skip": False}
                company_fields = company_text_fields(quick_result) if remove_company_text else []
                company_candidate_count += len(company_fields)
                for field in fields:
                    projected_dates = _date_anchors_for_year_field(field[0], field[2], 2026)
                    projected_2026_dates.extend(projected_dates)
                    source_year = field[0][field[2][0]:field[2][1]]
                    if source_year == "2026":
                        observed_2026_dates.extend(projected_dates)
                if adaptive.get("used"):
                    adaptive_images += 1
                    notes.append(f"{item.filename}：{adaptive.get('reason', '已运行自适应识别')}")
                if not fields:
                    if adaptive.get("used") and adaptive.get("needs_skip", True):
                        auto_skipped_images.append({
                            "media_path": item.filename,
                            "status": "AUTO_UNMODIFIED_LOW_CONFIDENCE",
                            "reason": str(adaptive.get("reason", "未识别到可信日期，保留原图")),
                        })
                scanned_images.append({
                    "item": item, "suffix": suffix, "fields": fields,
                    "company_fields": company_fields, "adaptive": adaptive,
                })
            else:
                target.writestr(item, data)

        if performance is not None:
            performance["scanned_image_count"] = len(scanned_images)
            performance["scan_and_first_pass_seconds"] = round(time.perf_counter() - first_pass_started, 4)
        today = BaseDateTime.now().astimezone().date()
        effective_target_year, trigger_date, rollback_reason = _select_effective_target_year(
            target_year,
            observed_2026_dates=observed_2026_dates,
            projected_2026_dates=projected_2026_dates,
            today=today,
        )

        rewrite_started = time.perf_counter() if perf_trace else None
        for scanned in scanned_images:
            item = scanned["item"]
            # Read one compressed media part at a time after target selection;
            # retain only OCR fields and metadata during the project-wide scan.
            data = source.read(item.filename)
            fields = scanned["fields"]
            company_fields = scanned["company_fields"]
            if fields or company_fields:
                detections_found += len(fields)
                with module.Image.open(io.BytesIO(data)) as opened:
                    rgb = module.np.asarray(
                        module.ImageOps.exif_transpose(opened).convert("RGB"), dtype=module.np.uint8
                    )[:, :, :3]
                bgr = module.cv2.cvtColor(rgb, module.cv2.COLOR_RGB2BGR)
                output_bgr = bgr
                image_notes, applied_fields = [], []
                if fields:
                    output_bgr, image_notes, applied_fields = replace_year_fields(
                        module, engine, output_bgr, fields, effective_target_year
                    )
                removed_fields = []
                if company_fields:
                    output_bgr, removed_fields = remove_company_text_fields(module, output_bgr, company_fields)
                output_rgb = module.cv2.cvtColor(output_bgr, module.cv2.COLOR_BGR2RGB)
                if not module.np.array_equal(rgb, output_rgb):
                    data = encode_like_source(module, output_rgb, scanned["suffix"])
                    if applied_fields:
                        changed_images += 1
                        notes.extend(map(str, image_notes))
                    changes.extend({
                        "media_path": item.filename,
                        "ocr_text": value,
                        "source_year": value[span[0]:span[1]],
                        "target_year": str(effective_target_year),
                        "changed_image": True,
                        "recognition_mode": "adaptive_small_text" if scanned["adaptive"].get("used") else "original_ocr",
                    } for value, span in applied_fields)
                    if removed_fields:
                        company_changed_images += 1
                        company_changes.extend({
                            "media_path": item.filename,
                            **field,
                            "changed_image": True,
                        } for field in removed_fields)
            target.writestr(item, data)
    os.replace(temporary, docx)
    if performance is not None:
        performance["selection_and_rewrite_seconds"] = round(time.perf_counter() - rewrite_started, 4)
        performance["total_seconds"] = round(time.perf_counter() - total_started, 4)
        performance["ocr_seconds"] = round(performance["ocr_seconds"], 4)
        performance["adaptive_or_local_ocr_calls"] = max(
            0, performance["ocr_calls"] - performance["primary_image_ocr_calls"]
        )
    if auto_skipped_images:
        result_status = "CHANGED_WITH_AUTO_SKIPS" if changed_images else "AUTO_SKIPPED_UNCERTAIN_IMAGES"
    else:
        result_status = "CHANGED" if changed_images else "UNCHANGED"
    result = {
        "status": result_status,
        "changed": bool(changed_images),
        "changed_images": changed_images,
        "detections": detections_found,
        "adaptive_small_text_images": adaptive_images,
        "auto_unmodified_images": auto_skipped_images,
        "target_year": effective_target_year,
        "effective_target_year": effective_target_year,
        "trigger_date": trigger_date,
        "rollback_reason": rollback_reason,
        "notes": notes,
        "changes": changes,
        "company_text_rewrite": {
            "status": (
                "NOT_SELECTED" if not remove_company_text
                else "CHANGED" if company_changes else "UNCHANGED"
            ),
            "selected": remove_company_text,
            "scanned_image_count": len(scanned_images) if remove_company_text else 0,
            "candidate_count": company_candidate_count,
            "changed_image_count": company_changed_images,
            "removed_text_count": len(company_changes),
            "changes": company_changes,
        },
    }
    if performance is not None:
        result["performance"] = performance
    return result


def rewrite_docx_stream(tool_root: Path, target_year: int) -> int:
    """Serve one JSONL DOCX request at a time while retaining one OCR engine."""
    started = time.perf_counter()
    runtime = load_engine(tool_root, 2026)
    perf_trace = os.getenv("EASYSOFTWARE_PERF_TRACE") == "1"
    print(json.dumps({
        "ready": True,
        "performance": {"engine_load_seconds": round(time.perf_counter() - started, 4)} if perf_trace else {},
    }), flush=True)
    for line in sys.stdin:
        try:
            request = json.loads(line)
            if request == {"op": "close"}:
                break
            if not isinstance(request, dict) or request.get("op") != "rewrite" or not isinstance(request.get("path"), str):
                raise ValueError("无效的日期 OCR 批次请求")
            docx = Path(request["path"]).resolve(strict=True)
            if not docx.is_file() or docx.suffix.lower() != ".docx":
                raise ValueError("日期 OCR 批次仅接受已有 DOCX 路径")
            try:
                result = rewrite_docx(
                    tool_root, docx, target_year, runtime=runtime,
                    rewrite_dates=bool(request.get("rewrite_dates", True)),
                    remove_company_text=bool(request.get("remove_company_text", False)),
                )
            except Exception as exc:
                docx.with_suffix(docx.suffix + ".date-year.tmp").unlink(missing_ok=True)
                result = {"status": "FAILED", "changed": False, "error": str(exc)}
        except Exception as exc:
            result = {"status": "FAILED_INVALID_REQUEST", "changed": False, "error": str(exc)}
        print(json.dumps(result, ensure_ascii=False), flush=True)
    return 0


def main() -> int:
    try:
        if len(sys.argv) >= 2 and sys.argv[1] == "--stream":
            tool_root = Path(sys.argv[2]).resolve()
            target_year = int(sys.argv[3])
            return rewrite_docx_stream(tool_root, target_year)
        tool_root = Path(sys.argv[1]).resolve()
        docx = Path(sys.argv[2]).resolve()
        target_year = int(sys.argv[3])
        flags = set(sys.argv[4:])
        result = rewrite_docx(
            tool_root, docx, target_year,
            rewrite_dates="--no-date-rewrite" not in flags,
            remove_company_text="--remove-company-text" in flags,
        )
        print(json.dumps(result, ensure_ascii=False))
        return 0
    except Exception as exc:
        if len(sys.argv) > 1 and sys.argv[1] == "--stream":
            candidates = []
        else:
            candidates = sys.argv[2:3]
        for value in candidates:
            candidate = Path(value).resolve()
            candidate.with_suffix(candidate.suffix + ".date-year.tmp").unlink(missing_ok=True)
        print(json.dumps({"status": "FAILED", "changed": False, "error": str(exc)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

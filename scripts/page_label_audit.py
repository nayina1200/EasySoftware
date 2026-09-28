"""Fail-closed PDF page-label validation against active DOCX PAGE fields.

This module reads DOCX package XML and PDF text/coordinates only.  It never
renders pages or invokes OCR, Word, or LibreOffice.
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from pypdf import PdfReader


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W, "r": R, "rel": PKG_REL, "mc": MC}
_PAGE_INSTRUCTION = re.compile(r"(?:^|\s)PAGE(?:\s|$)", re.IGNORECASE)
_PAGE_TOKEN = "\uFFF0PAGE\uFFF1"


def _norm(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _relationship_targets(archive: zipfile.ZipFile) -> dict[str, str]:
    root = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
    result: dict[str, str] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if not rel_id or not target or rel.get("TargetMode") == "External":
            continue
        result[rel_id] = posixpath.normpath(posixpath.join("word", target))
    return result


def _paragraph_tokens(paragraph: ET.Element) -> list[str]:
    """Return visible paragraph text, replacing PAGE field results by a marker."""
    tokens: list[str] = []
    fields: list[dict[str, Any]] = []

    def walk(node: ET.Element, *, root: bool = False) -> None:
        if node.tag == f"{{{MC}}}Fallback":
            return
        if node.tag == f"{{{W}}}fldSimple":
            instruction = node.get(f"{{{W}}}instr", "")
            if _PAGE_INSTRUCTION.search(instruction):
                tokens.append(_PAGE_TOKEN)
                for child in node:
                    if child.tag == f"{{{W}}}r":
                        continue
                    walk(child)
            else:
                for child in node:
                    walk(child)
            return
        if node.tag == f"{{{W}}}fldChar":
            kind = node.get(f"{{{W}}}fldCharType", "")
            if kind == "begin":
                fields.append({"instruction": "", "is_page": False, "result": False, "emitted": False})
            elif kind == "separate" and fields:
                current = fields[-1]
                current["is_page"] = bool(_PAGE_INSTRUCTION.search(current["instruction"]))
                current["result"] = True
                if current["is_page"]:
                    tokens.append(_PAGE_TOKEN)
                    current["emitted"] = True
            elif kind == "end" and fields:
                current = fields.pop()
                if current["is_page"] and not current["emitted"]:
                    tokens.append(_PAGE_TOKEN)
            return
        if node.tag == f"{{{W}}}instrText":
            if fields:
                fields[-1]["instruction"] += node.text or ""
            return
        if node.tag == f"{{{W}}}t":
            if not any(field["is_page"] and field["result"] for field in fields):
                tokens.append(node.text or "")
            return
        if node.tag == f"{{{W}}}tab":
            tokens.append(" ")
            return
        for child in node:
            # Text-box paragraphs are independent stories.  They are
            # inspected separately and must not be counted as part of the
            # containing header/footer paragraph.
            if not root and child.tag == f"{{{W}}}p":
                continue
            if child.tag == f"{{{MC}}}Fallback":
                continue
            walk(child)

    walk(paragraph, root=True)
    return tokens


def _active_paragraphs(story: ET.Element) -> list[ET.Element]:
    """List paragraphs in active markup while omitting compatibility fallback."""
    result: list[ET.Element] = []

    def visit(node: ET.Element) -> None:
        for child in node:
            if child.tag == f"{{{MC}}}Fallback":
                continue
            if child.tag == f"{{{W}}}p":
                result.append(child)
            visit(child)

    visit(story)
    return result


def _active_story_parts(archive: zipfile.ZipFile) -> tuple[dict[str, dict[str, Any]], list[str]]:
    issues: list[str] = []
    names = set(archive.namelist())
    try:
        document = ET.fromstring(archive.read("word/document.xml"))
        relationships = _relationship_targets(archive)
    except (KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        return {}, [f"DOCX OOXML 无法读取：{type(exc).__name__}: {exc}"]
    settings = None
    if "word/settings.xml" in names:
        try:
            settings = ET.fromstring(archive.read("word/settings.xml"))
        except ET.ParseError:
            issues.append("settings.xml 结构无法解析")
    even_odd = bool(settings is not None and settings.find(".//w:evenAndOddHeaders", NS) is not None)
    sections = document.findall(".//w:sectPr", NS)
    if not sections:
        issues.append("DOCX 未找到节属性，无法确定活动页眉/页脚")
        return {}, issues
    active: dict[str, dict[str, Any]] = {}
    inherited: dict[str, dict[str, str]] = {"header": {}, "footer": {}}
    for section_index, section in enumerate(sections):
        title_page = section.find("w:titlePg", NS) is not None
        for story_type in ("header", "footer"):
            refs = section.findall(f"w:{story_type}Reference", NS)
            for ref in refs:
                variant = ref.get(f"{{{W}}}type", "default")
                rel_id = ref.get(f"{{{R}}}id")
                part = relationships.get(rel_id or "")
                if part and part in names:
                    inherited[story_type][variant] = part
            variants = ["default"]
            if title_page:
                variants.append("first")
            if even_odd:
                variants.append("even")
            for variant in variants:
                part = inherited[story_type].get(variant) or inherited[story_type].get("default")
                if part:
                    item = active.setdefault(part, {"story_type": story_type, "sections": set(), "variants": set()})
                    item["sections"].add(section_index)
                    item["variants"].add(variant)
    return active, issues


def extract_page_label_templates(docx_path: str | Path) -> dict[str, Any]:
    """Extract PAGE field paragraph templates from active DOCX headers/footers.

    Each template identifies the OOXML part, section(s), paragraph index and
    XPath, textbox membership, and literal text before/after the PAGE field.
    An ambiguous paragraph remains in the result with ``ambiguous=True``.
    """
    templates: list[dict[str, Any]] = []
    issues: list[str] = []
    try:
        with zipfile.ZipFile(docx_path) as archive:
            active, active_issues = _active_story_parts(archive)
            issues.extend(active_issues)
            for part, details in sorted(active.items()):
                try:
                    story = ET.fromstring(archive.read(part))
                except (KeyError, ET.ParseError) as exc:
                    issues.append(f"活动故事部件无法解析：{part}: {type(exc).__name__}")
                    continue
                textbox_paragraphs = {
                    id(paragraph)
                    for textbox in story.findall(".//w:txbxContent", NS)
                    for paragraph in textbox.findall(".//w:p", NS)
                }
                for paragraph_index, paragraph in enumerate(_active_paragraphs(story)):
                    tokens = _paragraph_tokens(paragraph)
                    count = tokens.count(_PAGE_TOKEN)
                    if not count:
                        continue
                    ambiguous = count != 1
                    if ambiguous:
                        prefix = suffix = ""
                        issues.append(f"段落含 {count} 个 PAGE 字段，模板有歧义：{part}#{paragraph_index}")
                    else:
                        marker_index = tokens.index(_PAGE_TOKEN)
                        prefix = _norm("".join(tokens[:marker_index]))
                        suffix = _norm("".join(tokens[marker_index + 1 :]))
                    xpath = f".//w:p[{paragraph_index + 1}]"
                    templates.append({
                        "id": f"{part}#p{paragraph_index}",
                        "story_type": details["story_type"],
                        "part": part,
                        "section_indices": [section + 1 for section in sorted(details["sections"])],
                        "variants": sorted(details["variants"]),
                        "paragraph_index": paragraph_index,
                        "paragraph_xpath": xpath,
                        "in_textbox": id(paragraph) in textbox_paragraphs,
                        "prefix": prefix,
                        "suffix": suffix,
                        "field_count": count,
                        "ambiguous": ambiguous or not (prefix or suffix),
                    })
            if not templates:
                issues.append("活动页眉/页脚中未找到 PAGE 字段段落")
    except (OSError, zipfile.BadZipFile) as exc:
        issues.append(f"DOCX 文件无法读取：{type(exc).__name__}: {exc}")
    ambiguous = any(template["ambiguous"] for template in templates)
    return {"status": "ok" if templates and not issues and not ambiguous else "indeterminate",
            "templates": templates, "issues": issues}


def _chinese_number(value: str) -> int | None:
    digits = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
              "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    units = {"十": 10, "百": 100, "千": 1000}
    if not value or any(char not in digits and char not in units for char in value):
        return None
    total = section = number = 0
    for char in value:
        if char in digits:
            number = digits[char]
        else:
            unit = units[char]
            if unit == 10_000:
                section += number
                total = (total + section) * unit
                section = number = 0
            else:
                section += (number or (1 if unit == 10 else 0)) * unit
                number = 0
    return total + section + number


def _page_value(value: str) -> int | None:
    value = _norm(value)
    if value.isdigit():
        return int(value)
    return _chinese_number(value)


def _pdfminer_page_lines(layout: Any) -> list[dict[str, Any]]:
    """Read PDF text lines with pdfminer coordinates in page user space."""
    from pdfminer.layout import LTContainer, LTTextLine

    lines: list[dict[str, Any]] = []

    def visit(element: Any) -> None:
        if isinstance(element, LTTextLine):
            text = _norm(element.get_text())
            if text:
                lines.append({"text": text,
                              "bbox": [float(element.x0), float(element.y0),
                                       float(element.x1), float(element.y1)]})
            return
        if isinstance(element, LTContainer):
            for child in element:
                visit(child)

    visit(layout)
    return lines


def _box_values(box: Any) -> list[float] | None:
    try:
        return [float(box.left), float(box.bottom), float(box.right), float(box.top)]
    except Exception:
        return None


def _in_header_footer_band(line: dict[str, Any], crop: list[float]) -> bool:
    _, bottom, _, top = crop
    y = (line["bbox"][1] + line["bbox"][3]) / 2
    height = max(top - bottom, 1)
    return y >= top - height * 0.18 or y <= bottom + height * 0.18


def _near_crop_edge(line: dict[str, Any], crop: list[float], media: list[float]) -> tuple[bool, list[str]]:
    left, bottom, right, top = crop
    x0, y0, x1, y1 = line["bbox"]
    evidence = []
    if x0 <= left + 2:
        evidence.append("文本行起点接触 CropBox 左边界")
    if x1 >= right - 2:
        evidence.append("文本行终点接触 CropBox 右边界")
    if y0 <= bottom + 2:
        evidence.append("文本行接触 CropBox 下边界")
    if y1 >= top - 2:
        evidence.append("文本行接触 CropBox 上边界")
    if crop != media and (left > media[0] or bottom > media[1] or right < media[2] or top < media[3]):
        if x0 < left or x1 > right or y0 < bottom or y1 > top:
            evidence.append("文本行超出 CropBox 并被裁切")
    return bool(evidence), evidence


def _line_matches_template(line: dict[str, Any], template: dict[str, Any], total: int) -> tuple[str, int] | None:
    text = line["text"]
    prefix, suffix = template.get("prefix", ""), template.get("suffix", "")
    begin = text.find(prefix) if prefix else 0
    if begin < 0:
        return None
    after_prefix = begin + len(prefix)
    end = text.find(suffix, after_prefix) if suffix else len(text)
    if end < 0:
        return None
    value = _page_value(text[after_prefix:end])
    if value is None or not 1 <= value <= total:
        return None
    return text[begin:end + len(suffix)], value


def _split_template_match(lines: list[dict[str, Any]], template: dict[str, Any],
                          crop: list[float], total: int) -> tuple[dict[str, Any], str] | None:
    """Find a PAGE label whose known template pieces occupy nearby lines."""
    prefix, suffix = template.get("prefix", ""), template.get("suffix", "")
    if not prefix or not suffix:
        return None
    candidates = [line for line in lines if _in_header_footer_band(line, crop)]
    candidates.sort(key=lambda line: (-line["bbox"][1], line["bbox"][0]))
    for start in candidates:
        start_at = start["text"].find(prefix)
        if start_at < 0 or suffix in start["text"][start_at + len(prefix):]:
            continue
        nearby = [line for line in candidates if abs(line["bbox"][0] - start["bbox"][0]) <= 24]
        nearby.sort(key=lambda line: (-line["bbox"][1], line["bbox"][0]))
        try:
            start_index = nearby.index(start)
        except ValueError:
            continue
        for end_index in range(start_index + 1, min(start_index + 4, len(nearby))):
            end = nearby[end_index]
            if end["bbox"][1] < start["bbox"][1] - 30:
                break
            end_at = end["text"].find(suffix)
            if end_at < 0:
                continue
            pieces = [start["text"][start_at + len(prefix):]]
            pieces.extend(line["text"] for line in nearby[start_index + 1:end_index])
            pieces.append(end["text"][:end_at])
            page_value = _page_value(" ".join(pieces))
            if page_value is None or not 1 <= page_value <= total:
                continue
            bbox = [min(start["bbox"][0], end["bbox"][0]),
                    min(start["bbox"][1], end["bbox"][1]),
                    max(start["bbox"][2], end["bbox"][2]),
                    max(start["bbox"][3], end["bbox"][3])]
            merged = {"text": start["text"] + " / " + end["text"], "bbox": bbox}
            label = prefix + _norm(" ".join(pieces)) + suffix
            return merged, label
    return None


def _unverifiable_pages(reader: PdfReader, reason: str) -> list[dict[str, Any]]:
    return [{"page_number": index, "status": "indeterminate", "matched_template_id": None,
             "label_text": None, "line_text": None, "bbox": None,
             "crop_box": _box_values(page.cropbox), "media_box": _box_values(page.mediabox),
             "evidence": [], "issues": [reason]}
            for index, page in enumerate(reader.pages, start=1)]


def audit_pdf_page_labels(pdf_path: str | Path, templates: dict[str, Any] | list[dict[str, Any]]) -> dict[str, Any]:
    """Validate every page against extracted templates using PDF text lines and geometry.

    Per-page ``status`` is ``complete``, ``truncated``, or ``indeterminate``.
    Text and coordinates come from pdfminer LTTextLine objects; pypdf's text
    visitor matrices are deliberately not used because Word can emit identity
    text matrices for positioned header fragments. Missing pdfminer is a
    fail-closed indeterminate result.
    """
    if isinstance(templates, dict):
        template_rows = templates.get("templates", [])
        template_issues = list(templates.get("issues", []))
        template_ready = templates.get("status") == "ok"
    else:
        template_rows = templates
        template_issues = []
        template_ready = bool(template_rows)
    if not template_ready or not template_rows or any(row.get("ambiguous") for row in template_rows):
        return {"status": "indeterminate", "page_count": 0, "pages": [],
                "issues": template_issues + ["页码模板缺失或存在歧义"]}
    try:
        reader = PdfReader(str(pdf_path))
    except Exception as exc:
        return {"status": "indeterminate", "page_count": 0, "pages": [],
                "issues": [f"PDF 无法读取：{type(exc).__name__}: {exc}"]}
    try:
        from pdfminer.high_level import extract_pages
        layout_pages = iter(extract_pages(str(pdf_path)))
        first_layout = next(layout_pages, None)
    except ImportError:
        reason = "缺少 pdfminer.six，无法可靠提取 PDF 文本行坐标"
        return {"status": "indeterminate", "page_count": len(reader.pages),
                "pages": _unverifiable_pages(reader, reason), "issues": [reason]}
    except Exception as exc:
        reason = f"PDF 文本行/坐标解析失败：{type(exc).__name__}: {exc}"
        return {"status": "indeterminate", "page_count": len(reader.pages),
                "pages": _unverifiable_pages(reader, reason), "issues": [reason]}
    total = len(reader.pages)
    results: list[dict[str, Any]] = []
    global_issues: list[str] = []
    for index, page in enumerate(reader.pages, start=1):
        media = _box_values(page.mediabox)
        crop = _box_values(page.cropbox)
        try:
            layout = first_layout if index == 1 else next(layout_pages, None)
            if layout is None:
                raise ValueError("pdfminer 返回的页面数量不足")
            lines = _pdfminer_page_lines(layout)
        except Exception as exc:
            lines = []
            issues = [f"第 {index} 页文本行/坐标解析失败：{type(exc).__name__}: {exc}"]
        else:
            issues = []
        if not media or not crop or not lines:
            issues.append("页面缺少有效 MediaBox/CropBox 或可定位 PDF 文本行")
            results.append({"page_number": index, "status": "indeterminate", "matched_template_id": None,
                            "label_text": None, "line_text": None, "bbox": None,
                            "crop_box": crop, "media_box": media, "evidence": [], "issues": issues})
            continue
        complete = []
        partial = []
        for template in template_rows:
            if template.get("ambiguous"):
                continue
            prefix, suffix = template.get("prefix", ""), template.get("suffix", "")
            if not prefix and not suffix:
                continue
            for line in lines:
                if not _in_header_footer_band(line, crop):
                    continue
                text = line["text"]
                match = _line_matches_template(line, template, total)
                if match:
                    label, field_page = match
                    complete.append((template, line, label, field_page))
                    continue
                begin = text.find(prefix) if prefix else 0
                if begin < 0:
                    if suffix and suffix in text:
                        partial.append((template, line, "仅找到页码后缀，缺少前缀"))
                    continue
                after_prefix = begin + len(prefix)
                end = text.find(suffix, after_prefix) if suffix else len(text)
                reason = ("前后缀之间的页码为空、不可解析或超出 PDF 页数范围"
                          if end >= 0 else "找到页码前缀但缺少完整页码/后缀")
                partial.append((template, line, reason))
            split_match = _split_template_match(lines, template, crop, total)
            if split_match:
                split_line, split_label = split_match
                partial = [(template, split_line, "页码模板的前缀、实际页码和后缀被拆分到相邻 PDF 文本行",
                            split_label)]
        if len(complete) == 1:
            template, line, label, field_page = complete[0]
            x0, y0, x1, y1 = line["bbox"]
            crop_clips_label = x0 < crop[0] or y0 < crop[1] or x1 > crop[2] or y1 > crop[3]
            label_status = "truncated" if crop_clips_label else "complete"
            label_evidence = ["同一页眉/页脚文本行包含完整模板前缀、页码及后缀",
                              f"页码值在 1..{total} 范围内"]
            if crop_clips_label:
                label_evidence.append("完整文本行的坐标超出 CropBox，PDF 页面裁切了页码行")
            results.append({"page_number": index, "status": label_status, "matched_template_id": template["id"],
                            "label_text": label, "field_page_number": field_page, "line_text": line["text"],
                            "bbox": line["bbox"], "crop_box": crop, "media_box": media,
                            "evidence": label_evidence, "issues": issues})
        elif len(complete) > 1:
            issues.append("同一页匹配到多个页码模板，无法唯一定位")
            results.append({"page_number": index, "status": "indeterminate", "matched_template_id": None,
                            "label_text": None, "line_text": None, "bbox": None, "crop_box": crop,
                            "media_box": media, "evidence": [], "issues": issues})
        elif len(partial) == 1:
            template, line, reason, *label_data = partial[0]
            if label_data and "拆分到相邻" in reason:
                results.append({"page_number": index, "status": "truncated",
                                "matched_template_id": template["id"], "label_text": label_data[0],
                                "line_text": line["text"], "bbox": line["bbox"], "crop_box": crop,
                                "media_box": media,
                                "evidence": [reason, "相邻文本行坐标接近且位于页眉/页脚区域"], "issues": issues})
                continue
            clipped, edge_evidence = _near_crop_edge(line, crop, media)
            status = "truncated" if clipped else "indeterminate"
            if not clipped:
                issues.append("发现不完整页码模板，但 CropBox 未证明裁切")
            results.append({"page_number": index, "status": status,
                            "matched_template_id": template["id"], "label_text": None,
                            "line_text": line["text"], "bbox": line["bbox"], "crop_box": crop,
                            "media_box": media,
                            "evidence": [reason] + edge_evidence if clipped else [reason], "issues": issues})
        else:
            issues.append("未能在页眉/页脚区域唯一找到完整页码行")
            results.append({"page_number": index, "status": "indeterminate", "matched_template_id": None,
                            "label_text": None, "line_text": None, "bbox": None, "crop_box": crop,
                            "media_box": media, "evidence": [], "issues": issues})
    states = {item["status"] for item in results}
    status = "truncated" if "truncated" in states else "indeterminate" if "indeterminate" in states else "complete"
    return {"status": status, "page_count": total, "pages": results,
            "issues": global_issues, "method": "pdfminer_text_lines_and_pypdf_page_boxes_no_ocr"}


def audit_docx_pdf_page_labels(docx_path: str | Path, pdf_path: str | Path) -> dict[str, Any]:
    """Extract DOCX templates and audit the associated PDF in one call."""
    templates = extract_page_label_templates(docx_path)
    result = audit_pdf_page_labels(pdf_path, templates)
    result["template_audit"] = templates
    return result

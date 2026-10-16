"""Conservative, lossless relocation of a manual's login paragraph and image.

Only document.xml changes. Word subsequently resolves its original numbering
and exports the PDF through the existing, globally locked manual pipeline.
"""

from __future__ import annotations

import copy
import hashlib
import os
import re
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree as ET

from cover_rules import classify_cover_text


RULE_VERSION = "login-module-v3"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = {
    "w": W[1:-1],
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
CHINESE = "零〇一二三四五六七八九十百两"
LOGIN_TITLE = re.compile(r"^(?:(?:系统|用户|账号|账户|软件)的?)?登录(?:页面?|界面|窗口|模块|功能|操作|入口)?(?:介绍|说明|操作(?:说明)?|使用说明)?$")
FUNCTION_CHAPTER = re.compile(r"(?:功能|操作|使用|快速入门|控制面板)")
KNOWN_CHAPTER = re.compile(r"^(?:(?:软件|系统|平台|产品|基本|功能|操作|使用)(?:概述|介绍|简介|说明|功能|界面|指南|设计)?|关于系统)$")
SENTENCE = re.compile(r"[。！？；;!?，,]|(?:输入|填写|点击|单击|录入|校验|验证|即可|能够|可以)")
PROTECTED = ".//w:sectPr|.//w:fldChar|.//w:instrText|.//w:commentRangeStart|.//w:commentRangeEnd|.//w:ins|.//w:del"


class NeedsReview(ValueError):
    pass


def text(element: ET._Element) -> str:
    return "".join(element.xpath(".//w:t/text()", namespaces=NS))


def _value(element: ET._Element | None, location: str, default: str = "") -> str:
    node = element.find(location, NS) if element is not None else None
    return node.get(W + "val", default) if node is not None else default


def _chinese_number(value: str) -> int:
    if value.isdecimal():
        return int(value)
    total, digit = 0, 0
    for char in value:
        if char in "十百":
            total += (digit or 1) * (10 if char == "十" else 100)
            digit = 0
        else:
            digit = {"〇": 0, "零": 0, "两": 2}.get(char, "一二三四五六七八九".find(char) + 1)
    return total + digit


def _chinese_text(value: int) -> str:
    if not 1 <= value < 100:
        raise NeedsReview("中文标题编号超出安全顺延范围，请人工处理")
    digits = "零一二三四五六七八九"
    if value < 10:
        return digits[value]
    tens, units = divmod(value, 10)
    return (digits[tens] if tens > 1 else "") + "十" + (digits[units] if units else "")


@dataclass
class Prefix:
    family: str
    numbers: tuple[int, ...]
    spans: tuple[tuple[int, int], ...]
    end: int
    level: int


def _prefix(value: str, chinese_roots: bool = False) -> Prefix | None:
    match = re.match(r"^\s*(\d+(?:[.．]\d+)+)[.．、)）]?\s*", value)
    if match:
        parts = tuple(re.finditer(r"\d+", match.group(1)))
        start = match.start(1)
        return Prefix("decimal", tuple(int(p.group()) for p in parts), tuple((start + p.start(), start + p.end()) for p in parts), match.end(), len(parts))
    for family, pattern, level in (
        ("chapter", rf"^\s*第([{CHINESE}\d]+)(?:章|节|部分|篇)\s*[:：、]?\s*", 1),
        ("page", r"^\s*第(\d+)页\s*[:：]?\s*", 1),
        ("chinese", rf"^\s*([{CHINESE}]+)[、.．]\s*", 1),
        ("parenthesized", rf"^\s*[（(]([{CHINESE}\d]+)[）)]\s*", 2),
        ("decimal_single", r"^\s*(\d+)[、.．]\s*", 2 if chinese_roots else 1),
    ):
        match = re.match(pattern, value)
        if match:
            return Prefix(family, (_chinese_number(match.group(1)),), (match.span(1),), match.end(), level)
    return None


def _short_heading(value: str) -> bool:
    return bool(value.strip()) and len(value.strip()) <= 50 and not SENTENCE.search(value)


def _login_title(value: str) -> tuple[str, bool]:
    """Recognize an existing heading before considering any insertion."""
    prefix = _prefix(value)
    label = value[prefix.end:] if prefix else value
    label = re.sub(r"^\s*\d+[)）:：-]\s*", "", label)
    label = re.sub(r"\s+", "", label).strip("【】[]:：")
    core = re.sub(r"[（(][^()（）]{0,30}[)）]$", "", label)
    clear = bool(LOGIN_TITLE.fullmatch(core))
    plausible = clear or ("登录" in label and _short_heading(label) and not re.search(r"日志|记录|历史|统计", label))
    return label if plausible else "", clear


def _login_instruction(value: str) -> bool:
    if re.search(r"登录(?:日志|记录|历史)", value) and not re.search(r"登录(?:界面|页面|窗口|按钮)", value):
        return False
    return bool("登录" in value and re.search(r"账号|帐号|账户|用户名|登录名|密码|口令", value) and re.search(r"输入|填写|填入|录入|点击|单击", value))


@dataclass
class Heading:
    index: int
    level: int
    label: str
    source: str
    prefix: Prefix | None = None
    num_id: str = ""
    ilvl: int = 0


class Manual:
    def __init__(self, contents: dict[str, bytes], title: str):
        parser = ET.XMLParser(resolve_entities=False, no_network=True)
        self.document = ET.fromstring(contents["word/document.xml"], parser)
        self.body = self.document.find("w:body", NS)
        if self.body is None:
            raise NeedsReview("说明书缺少可识别的正文结构")
        self.nodes = list(self.body)
        self.title = re.sub(r"\s+", "", title)
        self.styles = {}
        if "word/styles.xml" in contents:
            self.styles = {s.get(W + "styleId"): s for s in ET.fromstring(contents["word/styles.xml"], parser).findall("w:style", NS)}
        self.numbering = ET.fromstring(contents["word/numbering.xml"], parser) if "word/numbering.xml" in contents else ET.Element(W + "numbering")
        self.nums = {n.get(W + "numId"): n for n in self.numbering.findall("w:num", NS)}
        self.abstracts = {n.get(W + "abstractNumId"): n for n in self.numbering.findall("w:abstractNum", NS)}

    def properties(self, p: ET._Element) -> list[ET._Element]:
        result = []
        direct = p.find("w:pPr", NS)
        if direct is not None:
            result.append(direct)
        style_id = _value(direct, "w:pStyle")
        seen = set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            style = self.styles.get(style_id)
            if style is None:
                break
            props = style.find("w:pPr", NS)
            if props is not None:
                result.append(props)
            style_id = _value(style, "w:basedOn")
        return result

    def property(self, p: ET._Element, location: str, default: str = "") -> str:
        for props in self.properties(p):
            value = _value(props, location)
            if value:
                return value
        return default

    def number(self, p: ET._Element) -> tuple[str, int, ET._Element | None]:
        num_id = self.property(p, "w:numPr/w:numId")
        ilvl = int(self.property(p, "w:numPr/w:ilvl", "0"))
        num = self.nums.get(num_id)
        abstract = self.abstracts.get(_value(num, "w:abstractNumId"))
        override = num.find(f'w:lvlOverride[@w:ilvl="{ilvl}"]/w:lvl', NS) if num is not None else None
        level = override if override is not None else (abstract.find(f'w:lvl[@w:ilvl="{ilvl}"]', NS) if abstract is not None else None)
        return num_id, ilvl, level

    def is_toc(self, p: ET._Element) -> bool:
        style = self.property(p, "w:pStyle")
        definition = self.styles.get(style)
        name = _value(definition, "w:name")
        return bool(re.search(r"(?i)^(?:TOC\s*\d*|目录\s*\d*)$", name or style) or re.search(r"[.．…]{3,}\s*\d+\s*$", text(p)))

    def headings(self) -> list[Heading]:
        if not hasattr(self, "_heading_cache"):
            self._heading_cache = self._detect_headings()
        return self._heading_cache

    def _detect_headings(self) -> list[Heading]:
        chinese_roots = any(
            (prefix is not None and prefix.family in {"chinese", "chapter"})
            or _value(self.number(p)[2], "w:numFmt") in {"chineseCounting", "chineseCountingThousand", "chineseLegalSimplified"}
            for p in self.nodes if p.tag == W + "p"
            for prefix in [_prefix(text(p))]
        )
        result = []
        body_started = False
        single_level_parent = None
        for index, p in enumerate(self.nodes):
            if p.tag != W + "p" or self.is_toc(p) or p.xpath(".//w:drawing|.//w:pict", namespaces=NS):
                continue
            value = text(p).strip()
            if not value or (not body_started and (
                classify_cover_text(value, self.title)["kind"] in {"cover_title", "manual_label"}
                or re.sub(r"\s+", "", value) in {"操作说明", "使用说明", "设计说明"}
            )):
                continue
            prefix = _prefix(text(p), chinese_roots)
            label = text(p)[prefix.end:].strip() if prefix else value
            if not _short_heading(label):
                continue
            num_id, ilvl, num_level = self.number(p)
            if prefix and num_id and num_id != "0":
                raise NeedsReview(f"第 {index + 1} 段标题同时含手写编号和自动编号，请人工确认")
            if prefix:
                heading = Heading(index, prefix.level, label, "普通文本编号", prefix)
                if prefix.family == "decimal" and single_level_parent is not None:
                    heading.level = single_level_parent.level + len(prefix.numbers) - 1
            elif num_id and num_id != "0":
                if num_level is None:
                    raise NeedsReview(f"第 {index + 1} 段的自动编号定义缺失")
                fmt, template = _value(num_level, "w:numFmt"), _value(num_level, "w:lvlText")
                if fmt in {"bullet", "none"}:
                    continue
                if fmt not in {"decimal", "chineseCounting", "chineseCountingThousand", "chineseLegalSimplified"}:
                    continue
                placeholders = re.findall(r"%[1-9]", template)
                level = len(placeholders) if len(placeholders) > 1 else ilvl + 1
                if ilvl == 0 and fmt == "decimal" and chinese_roots:
                    level = 2 if not re.search(r"[)）]", template) else 3
                heading = Heading(index, level, label, "Word 自动编号", num_id=num_id, ilvl=ilvl)
            else:
                outline = self.property(p, "w:outlineLvl")
                if outline and int(outline) < 9:
                    heading = Heading(index, int(outline) + 1, label, "Word 标题")
                elif KNOWN_CHAPTER.fullmatch(label):
                    heading = Heading(index, 1, label, "普通文本章节名")
                elif LOGIN_TITLE.fullmatch(label):
                    heading = Heading(index, 2, label, "普通文本登录标题")
                elif re.fullmatch(r"【[^【】]{1,24}】", label) and result:
                    heading = Heading(index, min(h.level for h in result) + 1, label, "普通文本括号标题")
                else:
                    continue
            result.append(heading)
            if heading.level == 1:
                single_level_parent = None
            if heading.num_id and heading.level == 2 and ilvl == 0 and len(re.findall(r"%[1-9]", _value(num_level, "w:lvlText"))) == 1:
                single_level_parent = heading
            body_started = True
        return result

    def blank(self, p: ET._Element) -> bool:
        return p.tag == W + "p" and not text(p).strip() and not p.xpath(".//w:drawing|.//w:pict|.//w:br|.//w:sectPr|.//w:fldChar|.//w:bookmarkStart|.//w:bookmarkEnd", namespaces=NS)

    def image(self, p: ET._Element) -> bool:
        return p.tag == W + "p" and not text(p).strip() and len(p.xpath(".//wp:inline", namespaces=NS)) == 1 and len(p.xpath(".//a:blip[@r:embed]", namespaces=NS)) == 1 and not p.xpath(".//wp:anchor|.//w:pict|" + PROTECTED, namespaces=NS)

    def bookmarks_safe(self, p: ET._Element) -> bool:
        starts = Counter(p.xpath(".//w:bookmarkStart/@w:id", namespaces=NS))
        ends = Counter(p.xpath(".//w:bookmarkEnd/@w:id", namespaces=NS))
        if starts != ends:
            return False
        active = set()
        for node in self.nodes:
            if node is p:
                return not active
            for marker in node.xpath(".//w:bookmarkStart|.//w:bookmarkEnd", namespaces=NS):
                if marker.tag == W + "bookmarkStart":
                    active.add(marker.get(W + "id"))
                else:
                    active.discard(marker.get(W + "id"))
        return False

    def neighbor(self, index: int, direction: int) -> int | None:
        for distance in range(1, 5):
            candidate = index + direction * distance
            if not 0 <= candidate < len(self.nodes):
                return None
            if not self.blank(self.nodes[candidate]):
                return candidate
        return None


def _result(status: str, message: str, **details: Any) -> dict[str, Any]:
    return {"rule_version": RULE_VERSION, "status": status, "message": message, "changed": False, **details}


def _existing_section(manual: Manual) -> dict[str, Any] | None:
    headers = []
    for index, p in enumerate(manual.nodes):
        if p.tag != W + "p" or manual.is_toc(p):
            continue
        label, clear = _login_title(text(p))
        if label:
            headers.append((index, label, clear))
    if not headers:
        return None
    if len(headers) > 1:
        raise NeedsReview("已存在多个登录相关标题，不再新增小节，请人工确认是否已有重复")
    index, label, clear = headers[0]
    if not clear:
        raise NeedsReview(f"已存在疑似登录标题“{label}”，不再新增，请人工确认图文归属")
    login_paragraph = image_paragraph = None
    for following in range(index + 1, min(len(manual.nodes), index + 25)):
        p = manual.nodes[following]
        value = text(p)
        if p.tag != W + "p":
            break
        prefix = _prefix(value)
        short_label = value[prefix.end:].strip() if prefix else value.strip()
        is_heading = _short_heading(short_label) and (
            prefix is not None or manual.property(p, "w:outlineLvl", "9") != "9"
            or manual.property(p, "w:numPr/w:numId") not in {"", "0"}
            or KNOWN_CHAPTER.fullmatch(short_label)
        )
        if is_heading:
            break
        if _login_instruction(value):
            login_paragraph = following + 1
        if manual.image(p):
            image_paragraph = following + 1
        if login_paragraph and image_paragraph:
            return _result("ALREADY_COMPLIANT", "已有独立登录小节，保持原样", source_paragraph=login_paragraph, image_paragraph=image_paragraph, heading=label)
    raise NeedsReview(f"已存在登录标题“{label}”，但图文归属不明确，不再新增，请人工确认")


def _plan(manual: Manual) -> dict[str, Any]:
    existing = _existing_section(manual)
    if existing is not None:
        return existing
    nodes = manual.nodes
    candidates = []
    for i, p in enumerate(nodes):
        value = text(p)
        if _login_instruction(value):
            if p.tag != W + "p" or len(value) > 400 or p.xpath(".//w:drawing|.//w:pict|" + PROTECTED, namespaces=NS) or _prefix(value) or manual.property(p, "w:numPr/w:numId") not in {"", "0"}:
                raise NeedsReview(f"第 {i + 1} 段登录内容混入标题、表格或其他复杂结构，请人工整理")
            candidates.append(i)
    if len(candidates) != 1:
        raise NeedsReview(f"找到 {len(candidates)} 段明确登录说明，无法唯一定位登录模块")
    login = candidates[0]
    neighbors = [i for d in (-1, 1) if (i := manual.neighbor(login, d)) is not None]
    images = [i for i in neighbors if manual.image(nodes[i])]
    if len(images) != 1:
        raise NeedsReview(f"第 {login + 1} 段登录说明相邻有 {len(images)} 张可安全配对图片，需要人工选择")
    image = images[0]
    if not manual.bookmarks_safe(nodes[login]) or not manual.bookmarks_safe(nodes[image]):
        raise NeedsReview("登录图文涉及跨段书签，无法安全移动，请人工整理")
    headings = manual.headings()
    if manual.document.xpath(".//w:instrText[contains(., 'TOC ')]|.//w:fldSimple[contains(@w:instr, 'TOC ')]", namespaces=NS):
        raise NeedsReview("说明书包含自动目录，新增登录小节及目录更新需人工处理")
    if not headings:
        raise NeedsReview("无法从 Word 标题及普通文本标题中识别正文层级")
    root_level = min(h.level for h in headings)
    roots = [h for h in headings if h.level == root_level]
    if len(roots) >= 2:
        target = roots[1]
        target_reason = "第二个正文一级标题（含普通文本标题）"
    elif FUNCTION_CHAPTER.search(roots[0].label):
        target = roots[0]
        target_reason = "唯一且明确的功能或操作章节"
    else:
        raise NeedsReview("只有一个正文一级标题且不是明确的功能或操作章节")
    end = next((h.index for h in roots if h.index > target.index), len(nodes))
    children = [h for h in headings if target.index < h.index < end and h.level > root_level]
    if not children:
        raise NeedsReview(f"“{target.label}”下没有可可靠识别的首个小标题")
    child_level = min(h.level for h in children)
    first = children[0]
    if first.level != child_level:
        raise NeedsReview("目标章节的小标题层级不连续，请人工处理")
    if first.index in {login, image} or nodes[first.index].xpath(PROTECTED, namespaces=NS):
        raise NeedsReview("首个小标题含分节、书签或复杂字段，无法安全借用格式")
    siblings = [h for h in children if h.level == child_level]
    edits = []
    if first.prefix:
        component = len(first.prefix.numbers) - 1
        parent = first.prefix.numbers[:component]
        expected = 1
        for heading in children:
            prefix = heading.prefix
            if heading.level > child_level and prefix is None and not heading.num_id:
                continue
            if prefix is None or len(prefix.numbers) <= component or prefix.numbers[:component] != parent:
                raise NeedsReview(f"第 {heading.index + 1} 段标题编号混用或无法联动顺延")
            if heading.level == child_level:
                if prefix.family != first.prefix.family or prefix.numbers[component] != expected:
                    raise NeedsReview("目标章节的小标题编号不连续或不是从 1 开始")
                expected += 1
            elif prefix.numbers[component] != expected - 1:
                raise NeedsReview("目标章节的下级标题与上级编号不一致")
            edits.append(_number_edit(nodes[heading.index], heading.index, prefix, component))
    elif first.num_id:
        num = manual.nums[first.num_id]
        _, _, definition = manual.number(nodes[first.index])
        override = num.find(f'w:lvlOverride[@w:ilvl="{first.ilvl}"]', NS)
        start = _value(override, "w:startOverride", _value(definition, "w:start", "1"))
        if start != "1" or definition.find("w:lvlRestart", NS) is not None:
            raise NeedsReview("小标题自动编号含特殊起始值或重启设置，需人工处理")
        previous_item = False
        for p in nodes[:first.index]:
            num_id, ilvl, _ = manual.number(p)
            if num_id == first.num_id:
                if ilvl < first.ilvl:
                    previous_item = False
                elif ilvl == first.ilvl:
                    previous_item = True
        if previous_item:
            raise NeedsReview("目标首个小标题沿用了之前的编号，无法确认从 1 开始")
        heading_indices = {h.index for h in children}
        for index in range(first.index, len(nodes)):
            num_id, ilvl, _ = manual.number(nodes[index])
            if num_id != first.num_id or ilvl != first.ilvl:
                continue
            if index < end and index not in heading_indices:
                raise NeedsReview("小标题与普通列表共用编号链，无法只顺延标题")
            if index >= end and first.ilvl == 0:
                raise NeedsReview("小标题编号链跨越后续章节，需人工确认顺延范围")
        for sibling in siblings:
            if sibling.prefix or (sibling.num_id, sibling.ilvl) != (first.num_id, first.ilvl):
                raise NeedsReview("目标章节的同级标题使用不同编号链，无法可靠顺延")
        # A Word list may have hand-typed descendants (e.g. 8.1 under auto 8).
        ordinal = 0
        for heading in children:
            if heading.level == child_level:
                ordinal += 1
            elif heading.prefix:
                prefix = heading.prefix
                component = len(re.findall(r"%[1-9]", _value(definition, "w:lvlText"))) - 1
                if component < 0 or len(prefix.numbers) <= component or prefix.numbers[component] != ordinal:
                    raise NeedsReview("自动编号下的手写子标题无法对应上级序号")
                edits.append(_number_edit(nodes[heading.index], heading.index, prefix, component))
            elif not heading.num_id:
                continue
            elif heading.num_id != first.num_id or heading.ilvl <= first.ilvl:
                raise NeedsReview("自动编号下出现无法确认的子标题编号链")
    elif any(h.prefix or h.num_id for h in children):
        raise NeedsReview("目标章节的标题编号混用，需人工处理")
    return _result(
        "READY", "可新增登录页小节，原有小节及其内容整体保留",
        source_paragraph=login + 1, image_paragraph=image + 1,
        target_chapter=target.label, target_reason=target_reason,
        target_paragraph=first.index + 1, template_heading=text(nodes[first.index]),
        heading_source=first.source, numbering="自动编号延续原列表" if first.num_id else ("手写编号及下级编号顺延" if first.prefix else "无编号标题"),
        heading_changes=edits, shifted_sibling_count=len(siblings),
        original_order="文→图" if login < image else "图→文", final_order="文→图",
    )


def _number_edit(p: ET._Element, index: int, prefix: Prefix, component: int) -> dict[str, Any]:
    if p.xpath(PROTECTED, namespaces=NS):
        raise NeedsReview(f"第 {index + 1} 段标题含字段、修订或分节，需人工调整编号")
    before = text(p)
    start, end = prefix.spans[component]
    old = before[start:end]
    value = prefix.numbers[component] + 1
    replacement = str(value).zfill(len(old)) if old.isdecimal() else _chinese_text(value)
    return {"paragraph": index + 1, "before": before, "after": before[:start] + replacement + before[end:], "start": start, "end": end, "replacement": replacement}


def _replace_span(p: ET._Element, start: int, end: int, replacement: str) -> None:
    position = 0
    inserted = False
    for node in p.iter(W + "t"):
        value = node.text or ""
        following = position + len(value)
        if position < end and following > start:
            left, right = max(0, start - position), min(len(value), end - position)
            node.text = value[:left] + (replacement if not inserted else "") + value[right:]
            inserted = True
        position = following
    if not inserted:
        raise ValueError("无法安全定位手写标题编号字符")


def _canonical(node: ET._Element) -> bytes:
    return ET.tostring(node, method="c14n")


def _apply_structure(manual: Manual, plan: dict[str, Any]) -> bytes:
    nodes = manual.nodes
    expected = Counter(_canonical(node) for node in nodes)
    first = nodes[plan["target_paragraph"] - 1]
    prefix = _prefix(text(first))
    new_heading = ET.Element(W + "p")
    props = first.find("w:pPr", NS)
    if props is not None:
        new_heading.append(copy.deepcopy(props))
    run = ET.SubElement(new_heading, W + "r")
    run_props = first.find("w:r/w:rPr", NS)
    if run_props is not None:
        run.append(copy.deepcopy(run_props))
    ET.SubElement(run, W + "t").text = (text(first)[:prefix.end] if prefix else "") + "登录页"
    for edit in plan["heading_changes"]:
        p = nodes[edit["paragraph"] - 1]
        expected[_canonical(p)] -= 1
        _replace_span(p, edit["start"], edit["end"], edit["replacement"])
        if text(p) != edit["after"]:
            raise ValueError("标题编号顺延校验失败")
        expected[_canonical(p)] += 1
    login, image = nodes[plan["source_paragraph"] - 1], nodes[plan["image_paragraph"] - 1]
    manual.body.remove(login)
    manual.body.remove(image)
    position = manual.body.index(first)
    for offset, node in enumerate((new_heading, login, image)):
        manual.body.insert(position + offset, node)
    expected[_canonical(new_heading)] += 1
    if +expected != Counter(_canonical(node) for node in manual.body):
        raise ValueError("登录整理内容保全校验失败，未写入文件")
    return ET.tostring(manual.document, xml_declaration=True, encoding="UTF-8", standalone=True)


def normalize(docx: Path, title: str, *, apply: bool = False) -> dict[str, Any]:
    """Analyze, then atomically apply a complete plan; ambiguous inputs stay intact."""
    docx = Path(docx)
    try:
        with zipfile.ZipFile(docx) as archive:
            infos = archive.infolist()
            contents = {name: archive.read(name) for name in ("word/document.xml", "word/styles.xml", "word/numbering.xml", "word/theme/theme1.xml") if name in archive.namelist()}
        if len({info.filename for info in infos}) != len(infos):
            raise NeedsReview("DOCX 包中含重复条目，请人工检查")
        manual = Manual(contents, title)
        plan = _plan(manual)
        if plan["status"] in {"READY", "ALREADY_COMPLIANT"}:
            import manual_login_typography as typography

            plan["structure_changed"] = plan["status"] == "READY"
            instruction = manual.nodes[plan["source_paragraph"] - 1]
            plan["typography"] = typography.plan(manual, contents, instruction)
            if plan["typography"]["status"] == "READY" and not plan["structure_changed"]:
                plan.update(status="READY", message="已有独立登录小节，仅对齐说明文字字体格式")
        plan["document"] = docx.name
        if not apply or plan["status"] != "READY":
            return plan
        if plan["structure_changed"]:
            _apply_structure(manual, plan)
        if plan["typography"]["status"] == "READY":
            before_text = text(instruction)
            typography.apply(instruction, plan["typography"])
            if text(instruction) != before_text:
                raise ValueError("登录说明格式对齐后文字保全校验失败")
        modified = ET.tostring(manual.document, xml_declaration=True, encoding="UTF-8", standalone=True)
        # Re-identification proves the inserted heading owns precisely the moved
        # paragraph/image, and makes repeated runs a no-op.
        final_contents = {**contents, "word/document.xml": modified}
        final_manual = Manual(final_contents, title)
        final = _plan(final_manual)
        if plan["typography"]["status"] == "READY":
            checked_font = typography.plan(final_manual, final_contents, final_manual.nodes[final["source_paragraph"] - 1])
            if checked_font["status"] != "ALREADY_COMPLIANT":
                raise ValueError("登录说明字体格式对齐复检失败，未写入文件")
        if final["status"] != "ALREADY_COMPLIANT":
            raise ValueError("登录小节整理后复检未通过，未写入文件")
        fd, temporary = tempfile.mkstemp(prefix=docx.stem + "_login_", suffix=".tmp", dir=docx.parent)
        os.close(fd)
        try:
            hashes = {}
            with zipfile.ZipFile(docx) as source, zipfile.ZipFile(temporary, "w") as output:
                for info in infos:
                    payload = modified if info.filename == "word/document.xml" else source.read(info)
                    output.writestr(info, payload)
                    hashes[info.filename] = hashlib.sha256(payload).digest()
            with zipfile.ZipFile(temporary) as checked:
                if any(hashlib.sha256(checked.read(name)).digest() != digest for name, digest in hashes.items()):
                    raise ValueError("图片或非目标 DOCX 内容保全校验失败")
            os.replace(temporary, docx)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
        message = "已新增登录页小节，图文归位，原标题及内容保留，编号已顺延" if plan["structure_changed"] else "已有独立登录小节，已对齐说明文字字体格式"
        if plan["typography"]["status"] == "READY":
            plan["typography"].update(status="APPLIED", message="登录说明字体格式已与正文对齐")
        return {**plan, "status": "APPLIED", "changed": True, "message": message, "verified": True}
    except NeedsReview as exc:
        return _result("NEEDS_MANUAL_REVIEW", str(exc), document=docx.name)
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.XMLSyntaxError) as exc:
        return _result("FAILED", f"登录模块整理失败：{exc}", document=docx.name)


def render_report(packages: list[dict[str, Any]], selected: bool) -> str:
    import html

    rows = []
    labels = {"READY": "待整理", "APPLIED": "已整理", "ALREADY_COMPLIANT": "已有独立小节", "NEEDS_MANUAL_REVIEW": "需人工处理", "FAILED": "处理失败"}
    for package in packages:
        result = package.get("login_module", {})
        if not result:
            continue
        state = result.get("status", "")
        if result.get("typography", {}).get("status") == "NEEDS_MANUAL_REVIEW" and state != "FAILED":
            state = "NEEDS_MANUAL_REVIEW"
        details = [str(result.get("message", ""))]
        font = result.get("typography", {})
        if font:
            details.append(font.get("message", ""))
            if font.get("reference_font"):
                details.append(f"正文参考：{font['reference_font']}，{font['reference_size']:g} 磅；{font['reference_count']} 段采用该格式")
        if result.get("target_chapter"):
            details.append(f"移入“{result['target_chapter']}”首个小节；{result.get('original_order', '')} → 文→图；{result.get('numbering', '')}")
        if result.get("source_paragraph"):
            details.append(f"原文字第 {result['source_paragraph']} 段，原图片第 {result.get('image_paragraph', '?')} 段")
        if result.get("heading_changes"):
            details.extend(f"{change['before']} → {change['after']}" for change in result["heading_changes"])
        color = "#b71c1c" if state in {"NEEDS_MANUAL_REVIEW", "FAILED"} else "#176b32"
        rows.append(f'<tr style="color:{color}"><td>{html.escape(str(package.get("name", result.get("document", ""))))}</td><td>{labels.get(state, html.escape(str(state)))}</td><td>{"<br>".join(html.escape(d) for d in details)}</td></tr>')
    content = '<table><thead><tr><th>说明书</th><th>登录模块</th><th>整理结果与位置</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>" if rows else ("<p>没有可处理的说明书。</p>" if selected else "<p>本次未选择此项整理。</p>")
    return "<section><h2>整理登录模块</h2>" + content + "</section>"

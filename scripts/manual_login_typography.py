"""Resolve inherited OOXML body typography and align only login instruction runs."""
from __future__ import annotations

from collections import Counter
from lxml import etree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = {"w": W[1:-1], "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")


def _chain(styles, style_id):
    result, seen = [], set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        style = styles.get(style_id)
        if style is None:
            break
        result.append(style)
        parent = style.find("w:basedOn", NS)
        style_id = parent.get(W + "val") if parent is not None else None
    return list(reversed(result))


class Typography:
    def __init__(self, contents):
        parser = ET.XMLParser(resolve_entities=False, no_network=True)
        root = ET.fromstring(contents.get("word/styles.xml", b'<w:styles xmlns:w="' + W[1:-1].encode() + b'"/>'), parser)
        self.styles = {s.get(W + "styleId"): s for s in root.findall("w:style", NS)}
        self.default = next((s.get(W + "styleId") for s in self.styles.values() if s.get(W + "type") == "paragraph" and s.get(W + "default") in {"1", "true", "on"}), "")
        self.defaults = root.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
        self.theme = ET.fromstring(contents["word/theme/theme1.xml"], parser) if "word/theme/theme1.xml" in contents else None

    def effective(self, p, run):
        pp = p.find("w:pPr", NS)
        ps = pp.find("w:pStyle", NS) if pp is not None else None
        rp = run.find("w:rPr", NS)
        rs = rp.find("w:rStyle", NS) if rp is not None else None
        layers = [(self.defaults, False)]
        for sid in (ps.get(W + "val") if ps is not None else self.default, rs.get(W + "val") if rs is not None else ""):
            layers.extend((style.find("w:rPr", NS), True) for style in _chain(self.styles, sid))
        layers.append((rp, False))
        languages = {}
        for props, _ in layers:
            lang = props.find("w:lang", NS) if props is not None else None
            if lang is not None:
                languages.update({key.split("}")[-1]: value for key, value in lang.attrib.items()})
        values = {"b": "0", "i": "0"}
        fonts = {}
        for props, is_style in layers:
            if props is None:
                continue
            rf = props.find("w:rFonts", NS)
            if rf is not None:
                for slot in SLOTS:
                    theme_key = "cstheme" if slot == "cs" else slot + "Theme"
                    theme_value = rf.get(W + theme_key)
                    literal = rf.get(W + slot)
                    if theme_value:
                        fonts[slot] = self.theme_font(theme_value, languages)
                    elif literal:
                        fonts[slot] = literal
            for tag in ("sz", "b", "i"):
                item = props.find("w:" + tag, NS)
                if item is None:
                    continue
                value = item.get(W + "val", "1")
                if tag in {"b", "i"}:
                    on = value not in {"0", "false", "off"}
                    values[tag] = str(int((values[tag] == "0") if is_style and on else (values[tag] == "1") if is_style else on))
                else:
                    values[tag] = value
        # Missing script slots follow the resolved Latin face in Word.
        latin = fonts.get("ascii") or fonts.get("hAnsi")
        east = fonts.get("eastAsia") or latin
        if not latin or not east or not values.get("sz") or any(v is None for v in fonts.values()):
            return None
        return tuple(fonts.get(slot) or (east if slot == "eastAsia" else latin) for slot in SLOTS) + (values["sz"], values["b"], values["i"])

    def theme_font(self, value, languages):
        if self.theme is None:
            return None
        family = "majorFont" if value.startswith("major") else "minorFont"
        root = self.theme.find("a:themeElements/a:fontScheme/a:" + family, NS)
        if root is None:
            return None
        kind = "ea" if value.endswith("EastAsia") else "cs" if value.endswith("Bidi") else "latin"
        font = root.find("a:" + kind, NS)
        face = font.get("typeface") if font is not None else ""
        if face:
            return face
        language = languages.get("eastAsia" if kind == "ea" else "bidi", "").lower()
        script = ""
        if kind == "ea":
            if language.startswith(("zh-cn", "zh-sg", "zh-hans")):
                script = "Hans"
            elif language.startswith(("zh-tw", "zh-hk", "zh-mo", "zh-hant")):
                script = "Hant"
            elif language.startswith("ja"):
                script = "Jpan"
            elif language.startswith("ko"):
                script = "Hang"
        elif kind == "cs":
            if language.startswith(("ar", "fa", "ur")):
                script = "Arab"
            elif language.startswith("he"):
                script = "Hebr"
        font = root.find('a:font[@script="' + script + '"]', NS) if script else None
        return font.get("typeface") if font is not None else None

    def signatures(self, p):
        return [self.effective(p, run) for run in p.findall(".//w:r", NS) if "".join(run.xpath(".//w:t/text()", namespaces=NS)).strip()]


def plan(manual, contents, login):
    if login.xpath(".//w:ins|.//w:del|.//w:fldChar|.//w:instrText|.//w:fldSimple", namespaces=NS):
        return {"status": "NEEDS_MANUAL_REVIEW", "message": "登录说明含修订或字段，字体格式需人工确认"}
    resolver = Typography(contents)
    try:
        heading_indices = {h.index for h in manual.headings()}
    except ValueError as exc:
        return {"status": "NEEDS_MANUAL_REVIEW", "message": "正文字体参考不明确：" + str(exc)}
    counts, examples = Counter(), {}
    first_heading = min(heading_indices, default=len(manual.nodes))
    for index, p in enumerate(manual.nodes):
        value = "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()
        if index <= first_heading or index in heading_indices or p is login or p.tag != W + "p" or len(value) < 15:
            continue
        if p.xpath(".//w:drawing|.//w:pict|.//w:fldChar|.//w:instrText", namespaces=NS) or manual.is_toc(p):
            continue
        if manual.property(p, "w:numPr/w:numId") not in {"", "0"}:
            continue
        signatures = resolver.signatures(p)
        if signatures and None not in signatures and len(set(signatures)) == 1:
            signature = signatures[0]
            counts[signature] += 1
            examples.setdefault(signature, index + 1)
    ranked = counts.most_common()
    if not ranked or ranked[0][1] < 2 or (len(ranked) > 1 and ranked[0][1] == ranked[1][1]):
        return {"status": "NEEDS_MANUAL_REVIEW", "message": "没有至少两段且占优的普通正文字体参考，请人工确认登录说明格式"}
    reference, count = ranked[0]
    current = resolver.signatures(login)
    changed = bool(current) and any(signature != reference for signature in current)
    return {"status": "READY" if changed else "ALREADY_COMPLIANT", "message": "登录说明字体格式将与正文对齐" if changed else "登录说明字体格式已与正文一致", "reference": list(reference), "reference_paragraph": examples[reference], "reference_count": count, "reference_font": reference[2], "reference_size": int(reference[4]) / 2, "changed": changed}


def apply(login, result):
    reference = result["reference"]
    for run in login.findall(".//w:r", NS):
        if not "".join(run.xpath(".//w:t/text()", namespaces=NS)).strip():
            continue
        props = run.find("w:rPr", NS)
        if props is None:
            props = ET.Element(W + "rPr")
            run.insert(0, props)
        # Retain color, highlighting, hyperlinks, bookmarks, spacing and content.
        for tag in ("rFonts", "sz", "szCs", "b", "bCs", "i", "iCs"):
            for old in props.findall("w:" + tag, NS):
                props.remove(old)
        rf = ET.SubElement(props, W + "rFonts")
        for slot, face in zip(SLOTS, reference[:4]):
            rf.set(W + slot, face)
        for tag, value in (("sz", reference[4]), ("szCs", reference[4]), ("b", reference[5]), ("bCs", reference[5]), ("i", reference[6]), ("iCs", reference[6])):
            ET.SubElement(props, W + tag).set(W + "val", value)

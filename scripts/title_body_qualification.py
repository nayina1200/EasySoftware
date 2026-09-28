"""Strict TXT/manual pairing and structure-preserving DOCX text extraction."""
from __future__ import annotations

import hashlib
import re
import unicodedata
import zipfile
from pathlib import Path
from typing import Any, Callable
from xml.etree import ElementTree as ET

from title_body_cache import TitleBodyCache, semantic_input_hash


STRATEGY_VERSION = "title-body-s1-v2-section-roles"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WORD_NS}}}"

_OVERVIEW_MARKERS = (
    "概述", "简介", "摘要", "引言", "overview", "summary",
)
_MAIN_FUNCTION_MARKERS = (
    "主要功能", "核心功能", "功能模块", "应用功能", "系统功能", "产品功能", "业务功能",
)
_FUNCTION_MARKERS = (
    *_MAIN_FUNCTION_MARKERS, "功能介绍", "功能说明", "功能特点", "功能展示", "操作说明", "操作指南",
)
_NON_FUNCTION_TOP_LEVEL_MARKERS = (
    "环境", "术语", "编写目的", "文档约定", "参考资料", "版本记录", "目录", "前言",
)


def _normalise_name(value: str) -> str:
    return unicodedata.normalize("NFKC", value).strip().casefold()


def _section_role(title: str, outline_level: int | None) -> tuple[str, bool]:
    normalized = re.sub(r"\s+", "", unicodedata.normalize("NFKC", title)).casefold()
    if any(marker in normalized for marker in _OVERVIEW_MARKERS):
        return "overview", False
    is_main_function = any(marker in normalized for marker in _MAIN_FUNCTION_MARKERS)
    if any(marker in normalized for marker in _FUNCTION_MARKERS):
        return ("main_function" if is_main_function else "function"), is_main_function
    if outline_level is not None and outline_level <= 1 and not any(
        marker in normalized for marker in _NON_FUNCTION_TOP_LEVEL_MARKERS
    ):
        # A top-level product workflow is useful representative material even
        # when its heading does not literally contain “功能”.
        return "function", False
    return "general", False


def _manual_title_stem(path: Path) -> str:
    stem = unicodedata.normalize("NFKC", path.stem)
    if stem.endswith("说明"):
        stem = stem[:-2]
    return stem


def _association_stem(path: Path) -> str:
    stem = unicodedata.normalize("NFKC", path.stem)
    for suffix in ("说明", "代码"):
        if stem.endswith(suffix):
            return stem[:-len(suffix)]
    return stem


def _project_id(relative_docx: str, relative_txt: str | None) -> str:
    identity = f"{relative_docx}\0{relative_txt or ''}".replace("\\", "/")
    return "project:" + hashlib.sha256(identity.encode("utf-8")).hexdigest()[:24]


def _style_metadata(style_id: str | None, styles: dict[str, dict[str, Any]]) -> tuple[str | None, int | None, bool]:
    if not style_id:
        return None, None, False
    current = style_id
    seen: set[str] = set()
    display_name = ""
    outline_level: int | None = None
    is_toc = False
    while current and current not in seen and current in styles:
        seen.add(current)
        item = styles[current]
        display_name = display_name or item["name"]
        if outline_level is None and item["outline_level"] is not None:
            try:
                outline_level = int(item["outline_level"])
            except (TypeError, ValueError):
                pass
        label = f"{current} {item['name']}".casefold()
        is_toc = is_toc or current.casefold().startswith("toc") or "目录" in label or "table of contents" in label
        current = item["based_on"]
    return display_name or style_id, outline_level, is_toc


def _paragraph_content_nodes(paragraph: ET.Element):
    stack = list(reversed(list(paragraph)))
    while stack:
        node = stack.pop()
        if node.tag == W + "p":
            continue
        yield node
        stack.extend(reversed(list(node)))


def _on_off_enabled(element: ET.Element) -> bool:
    return element.get(W + "val", "1").casefold() not in {"0", "false", "off", "no"}


def _inherited_on_off(style_id: str | None, property_name: str, styles: dict[str, dict[str, Any]]) -> bool | None:
    current = style_id
    seen: set[str] = set()
    while current and current not in seen and current in styles:
        seen.add(current)
        value = styles[current].get("run_properties", {}).get(property_name)
        if value is not None:
            return value
        current = styles[current]["based_on"]
    return None


def _paragraph_text(
    paragraph: ET.Element,
    styles: dict[str, dict[str, Any]],
    *,
    default_paragraph_style: str | None = None,
    default_character_style: str | None = None,
    paragraph_default_run_properties: dict[str, bool] | None = None,
    document_default_run_properties: dict[str, bool] | None = None,
) -> str:
    pieces: list[str] = []
    ppr = paragraph.find(W + "pPr")
    paragraph_style = ppr.find(W + "pStyle") if ppr is not None else None
    paragraph_style_id = paragraph_style.get(W + "val") if paragraph_style is not None else default_paragraph_style
    paragraph_run_properties = ppr.find(W + "rPr") if ppr is not None else None

    def hidden_run(node: ET.Element) -> bool:
        properties = node.find(W + "rPr")
        character_style = properties.find(W + "rStyle") if properties is not None else None
        character_style_id = character_style.get(W + "val") if character_style is not None else None
        for property_name in ("vanish", "webHidden"):
            direct = properties.find(W + property_name) if properties is not None else None
            if direct is not None:
                if _on_off_enabled(direct):
                    return True
                continue
            paragraph_direct = paragraph_run_properties.find(W + property_name) if paragraph_run_properties is not None else None
            if paragraph_direct is not None:
                if _on_off_enabled(paragraph_direct):
                    return True
                continue
            style_candidates = (
                character_style_id or default_character_style,
                paragraph_style_id,
            )
            for style_id in style_candidates:
                inherited = _inherited_on_off(style_id, property_name, styles)
                if inherited is not None:
                    if inherited:
                        return True
                    break
            else:
                default_value = (paragraph_default_run_properties or {}).get(property_name)
                if default_value is None:
                    default_value = (document_default_run_properties or {}).get(property_name)
                if default_value:
                    return True
        return False

    def visit(node: ET.Element) -> None:
        if node.tag == W + "r":
            if hidden_run(node):
                return
        for child in node:
            if child.tag == W + "p":
                continue
            if child.tag == W + "t":
                pieces.append(child.text or "")
            elif child.tag == W + "tab":
                pieces.append("\t")
            elif child.tag in {W + "br", W + "cr"}:
                pieces.append("\n")
            else:
                visit(child)

    visit(paragraph)
    return "".join(pieces)


def _paragraph_structure(
    root: ET.Element,
    paragraphs: list[ET.Element],
    styles: dict[str, dict[str, Any]],
) -> tuple[dict[int, bool], dict[int, tuple[bool, bool]]]:
    parents = {child: parent for parent in root.iter() for child in parent}
    toc_indexes: dict[int, bool] = {}
    context: dict[int, tuple[bool, bool]] = {}
    field_stack: list[dict[str, Any]] = []
    for index, paragraph in enumerate(paragraphs):
        in_table = False
        in_textbox = False
        parent = parents.get(paragraph)
        while parent is not None:
            in_table = in_table or parent.tag == W + "tbl"
            in_textbox = in_textbox or parent.tag == W + "txbxContent"
            parent = parents.get(parent)
        context[index] = (in_table, in_textbox)
        toc = any(bool(frame.get("is_toc")) for frame in field_stack)
        ppr = paragraph.find(W + "pPr")
        pstyle = ppr.find(W + "pStyle") if ppr is not None else None
        style_id = pstyle.get(W + "val") if pstyle is not None else ""
        _, _, style_is_toc = _style_metadata(style_id, styles)
        toc = toc or style_is_toc
        for node in _paragraph_content_nodes(paragraph):
            if node.tag == W + "fldSimple":
                instruction = node.get(W + "instr", "")
                if re.search(r"\bTOC\b", instruction, re.IGNORECASE):
                    toc = True
            elif node.tag == W + "fldChar":
                kind = node.get(W + "fldCharType", "")
                if kind == "begin":
                    field_stack.append({"instruction": "", "is_toc": False})
                elif kind == "separate" and field_stack and field_stack[-1]["is_toc"]:
                    toc = True
                elif kind == "end" and field_stack:
                    if field_stack[-1]["is_toc"]:
                        toc = True
                    field_stack.pop()
            elif node.tag == W + "instrText" and field_stack:
                field_stack[-1]["instruction"] += node.text or ""
                if re.search(r"\bTOC\b", field_stack[-1]["instruction"], re.IGNORECASE):
                    field_stack[-1]["is_toc"] = True
                    toc = True
        if any(bool(frame.get("is_toc")) for frame in field_stack):
            toc = True
        toc_indexes[index] = toc
    return toc_indexes, context


def extract_docx_paragraphs(path: str | Path, title: str = "") -> list[dict[str, Any]]:
    """Extract visible paragraph text and structural evidence directly from OOXML."""
    docx_path = Path(path)
    with zipfile.ZipFile(docx_path) as archive:
        document_root = ET.fromstring(archive.read("word/document.xml"))
        try:
            styles_root = ET.fromstring(archive.read("word/styles.xml"))
        except KeyError:
            styles_root = ET.Element(W + "styles")
    styles: dict[str, dict[str, Any]] = {}
    for style in styles_root.iter(W + "style"):
        style_id = style.get(W + "styleId", "")
        if not style_id:
            continue
        name = style.find(W + "name")
        based_on = style.find(W + "basedOn")
        ppr = style.find(W + "pPr")
        outline = ppr.find(W + "outlineLvl") if ppr is not None else None
        styles[style_id] = {
            "name": name.get(W + "val", "") if name is not None else "",
            "based_on": based_on.get(W + "val", "") if based_on is not None else "",
            "outline_level": outline.get(W + "val") if outline is not None else None,
            "run_properties": {
                property_name: _on_off_enabled(rpr.find(W + property_name))
                for property_name in ("vanish", "webHidden")
                if (rpr := style.find(W + "rPr")) is not None and rpr.find(W + property_name) is not None
            },
        }
    default_paragraph_style = next((
        style.get(W + "styleId") for style in styles_root.iter(W + "style")
        if style.get(W + "type") == "paragraph"
        and style.get(W + "default", "0").casefold() in {"1", "true", "on"}
    ), None)
    default_character_style = next((
        style.get(W + "styleId") for style in styles_root.iter(W + "style")
        if style.get(W + "type") == "character"
        and style.get(W + "default", "0").casefold() in {"1", "true", "on"}
    ), None)
    doc_defaults = styles_root.find(W + "docDefaults")
    run_defaults: dict[str, bool] = {}
    paragraph_run_defaults: dict[str, bool] = {}
    if doc_defaults is not None:
        run_default = doc_defaults.find(W + "rPrDefault")
        run_default_properties = run_default.find(W + "rPr") if run_default is not None else None
        para_default = doc_defaults.find(W + "pPrDefault")
        para_default_properties = para_default.find(W + "pPr") if para_default is not None else None
        para_run_default_properties = para_default_properties.find(W + "rPr") if para_default_properties is not None else None
        for property_name in ("vanish", "webHidden"):
            for properties, target in (
                (run_default_properties, run_defaults),
                (para_run_default_properties, paragraph_run_defaults),
            ):
                element = properties.find(W + property_name) if properties is not None else None
                if element is not None:
                    target[property_name] = _on_off_enabled(element)
    xml_paragraphs = list(document_root.iter(W + "p"))
    toc_flags, structure = _paragraph_structure(document_root, xml_paragraphs, styles)

    title_key = _normalise_name(title)
    extracted: list[dict[str, Any]] = []
    current_section = "section:front"
    current_section_title = ""
    current_section_role = "general"
    current_is_main_function = False
    for index, paragraph in enumerate(xml_paragraphs):
        text = _paragraph_text(
            paragraph,
            styles,
            default_paragraph_style=default_paragraph_style,
            default_character_style=default_character_style,
            paragraph_default_run_properties=paragraph_run_defaults,
            document_default_run_properties=run_defaults,
        )
        if not text.strip():
            continue
        ppr = paragraph.find(W + "pPr")
        pstyle = ppr.find(W + "pStyle") if ppr is not None else None
        style_id = pstyle.get(W + "val") if pstyle is not None else None
        style_name, style_outline, style_is_toc = _style_metadata(style_id, styles)
        outline_node = ppr.find(W + "outlineLvl") if ppr is not None else None
        outline_level = style_outline
        if outline_node is not None:
            try:
                outline_level = int(outline_node.get(W + "val", ""))
            except (TypeError, ValueError):
                outline_level = None
        paragraph_id = f"paragraph:{index + 1}"
        is_heading = (outline_level is not None and outline_level < 9) or (style_id or "").casefold().startswith("heading")
        is_toc = bool(toc_flags.get(index, False) or style_is_toc)
        if is_heading and not is_toc:
            current_section = paragraph_id
            current_section_title = text.strip()
            current_section_role, current_is_main_function = _section_role(text, outline_level)
        normalized_text = _normalise_name(text)
        is_title_repeat = bool(title_key and normalized_text == title_key)
        in_table, in_textbox = structure[index]
        extracted.append({
            "paragraph_id": paragraph_id,
            "text": text,
            "section_id": current_section,
            "section_title": current_section_title,
            "section_role": current_section_role,
            "is_overview": current_section_role == "overview",
            "is_functional": current_section_role in {"function", "main_function"},
            "is_main_function": current_is_main_function,
            "style_id": style_id,
            "style_name": style_name,
            "outline_level": outline_level,
            "is_heading": is_heading,
            "is_toc": is_toc,
            "in_table": in_table,
            "in_textbox": in_textbox,
            "is_title_repeat": is_title_repeat,
            "structure_uncertain": bool(style_id and style_id not in styles),
        })
    return extracted


def _not_evaluated_result(reason: str, code: str, **extra: Any) -> dict[str, Any]:
    return {
        "status": "NOT_EVALUATED", "reason": reason,
        "diagnostic_code": code, "failure": code,
        "core_claims": [], "evidence": [], **extra,
    }


def _status(value: Any) -> str:
    normalized = str(value or "").strip().casefold()
    if normalized in {"match", "matched", "符合"}:
        return "MATCH"
    if normalized in {"mismatch", "not_match", "不匹配", "错配"}:
        return "MISMATCH"
    return "NOT_EVALUATED"


def _invalid_evidence(reason: str, **extra: Any) -> dict[str, Any]:
    return _not_evaluated_result(
        reason, str(extra.pop("diagnostic_code", "EVIDENCE_INVALID")),
        evidence_valid=False, **extra,
    )


def normalize_claim_evidence_result(result: dict[str, Any]) -> dict[str, Any]:
    """Flatten evidence nested under model claims into stable report/cache IDs."""
    normalized = dict(result)
    claims = normalized.get("core_claims")
    if not isinstance(claims, list) or not any(isinstance(item, dict) and "evidence" in item for item in claims):
        return normalized
    flattened: list[dict[str, Any]] = []
    normalized_claims: list[dict[str, Any]] = []
    for claim_index, source_claim in enumerate(claims, 1):
        if not isinstance(source_claim, dict):
            normalized_claims.append(source_claim)
            continue
        claim_id = str(source_claim.get("id") or f"claim-{claim_index}")
        claim = {key: value for key, value in source_claim.items() if key != "evidence"}
        claim["id"] = claim_id
        evidence_ids: list[str] = []
        claim_evidence = source_claim.get("evidence")
        if isinstance(claim_evidence, list):
            for evidence_index, source_evidence in enumerate(claim_evidence, 1):
                if not isinstance(source_evidence, dict):
                    flattened.append(source_evidence)
                    continue
                evidence_id = str(source_evidence.get("id") or f"evidence-{claim_index}-{evidence_index}")
                evidence_ids.append(evidence_id)
                flattened.append({
                    **source_evidence,
                    "id": evidence_id,
                    "claim_id": claim_id,
                })
        claim["evidence_ids"] = evidence_ids
        normalized_claims.append(claim)
    normalized["core_claims"] = normalized_claims
    normalized["evidence"] = flattened
    return normalized


def validate_claim_evidence_structure(result: dict[str, Any]) -> str | None:
    """Require every decisive claim to map to typed evidence before accepting semantics."""
    status = _status(result.get("status"))
    if status == "NOT_EVALUATED":
        return None
    claims = result.get("core_claims")
    evidence = result.get("evidence")
    if not isinstance(claims, list) or not claims or not isinstance(evidence, list) or not evidence:
        return "MATCH 或 MISMATCH 缺少结构化主张或证据"

    claims_by_id: dict[str, dict[str, Any]] = {}
    for claim in claims:
        if not isinstance(claim, dict):
            return "核心主张结构无效"
        claim_id = claim.get("id")
        text = claim.get("claim")
        supported = claim.get("supported")
        evidence_ids = claim.get("evidence_ids")
        if (not isinstance(claim_id, str) or not claim_id.strip()
                or not isinstance(text, str) or not text.strip()
                or supported not in {"yes", "no", "unclear"}
                or not isinstance(evidence_ids, list) or not evidence_ids
                or not all(isinstance(value, str) and value for value in evidence_ids)
                or len(set(evidence_ids)) != len(evidence_ids)):
            return "核心主张缺少 ID、内容、支持状态或证据映射"
        if claim_id in claims_by_id:
            return "核心主张 ID 重复"
        claims_by_id[claim_id] = claim

    evidence_by_id: dict[str, dict[str, Any]] = {}
    for citation in evidence:
        if not isinstance(citation, dict):
            return "正文证据结构无效"
        evidence_id = citation.get("id")
        claim_id = citation.get("claim_id")
        paragraph_id = citation.get("paragraph_id")
        quote = citation.get("quote")
        relation = citation.get("relation")
        if (not isinstance(evidence_id, str) or not evidence_id.strip()
                or not isinstance(claim_id, str) or claim_id not in claims_by_id
                or not isinstance(paragraph_id, (str, int)) or isinstance(paragraph_id, bool)
                or not isinstance(quote, str) or not quote.strip()
                or relation not in {"supports", "contradicts", "context"}):
            return "正文证据缺少 ID、段落、引文、主张映射或关系"
        if evidence_id in evidence_by_id:
            return "正文证据 ID 重复"
        evidence_by_id[evidence_id] = citation

    for claim_id, claim in claims_by_id.items():
        linked_ids = claim["evidence_ids"]
        if any(evidence_id not in evidence_by_id for evidence_id in linked_ids):
            return "核心主张引用了不存在的证据 ID"
        linked = [evidence_by_id[evidence_id] for evidence_id in linked_ids]
        if any(item["claim_id"] != claim_id for item in linked):
            return "主张与证据的双向映射不一致"
        supported = claim["supported"]
        relations = {item["relation"] for item in linked}
        if supported == "yes" and "supports" not in relations:
            return "被判为支持的主张没有对应支持证据"
        if supported == "no" and "contradicts" not in relations:
            return "被判为不支持的主张没有明确矛盾证据"

    linked_ids = {evidence_id for claim in claims for evidence_id in claim["evidence_ids"]}
    if linked_ids != set(evidence_by_id):
        return "存在未关联到核心主张的正文证据"
    if status == "MATCH":
        if any(claim["supported"] != "yes" for claim in claims):
            return "MATCH 含有未获支持的标题主张"
        if any(item["relation"] == "contradicts" for item in evidence):
            return "MATCH 含有明确矛盾证据"
    elif not any(claim["supported"] == "no" for claim in claims):
        return "MISMATCH 缺少明确不符的标题主张"
    return None


def validate_evidence_result(
    result: dict[str, Any],
    paragraphs: list[dict[str, Any]],
    *,
    allow_rebind: bool = False,
) -> dict[str, Any]:
    """Validate cited original IDs/quotes; title repeats and TOC rows cannot be sole support."""
    if result.get("judgment_source") == "typesafe_jev_score":
        normalized = dict(result)
        score = normalized.get("score_0_100")
        threshold = normalized.get("decision_threshold")
        status = str(normalized.get("status", "")).upper()
        valid_score = (
            isinstance(score, (int, float)) and not isinstance(score, bool)
            and isinstance(threshold, (int, float)) and not isinstance(threshold, bool)
            and 0 <= float(score) <= 100 and 0 <= float(threshold) <= 100
            and status == ("MATCH" if float(score) >= float(threshold) else "MISMATCH")
            and normalized.get("evidence") == []
            and normalized.get("core_claims") == []
        )
        if not valid_score:
            return _not_evaluated_result("Jev Score 缓存或结果结构无效", "SCORE_RESULT_INVALID")
        return normalized
    status = _status(result.get("status"))
    normalized = normalize_claim_evidence_result(result)
    normalized["status"] = status
    claim_error = validate_claim_evidence_structure(normalized)
    if claim_error:
        return _invalid_evidence(claim_error, failure=True, diagnostic_code="EVIDENCE_INVALID")
    if result.get("evidence_valid") is False:
        return _invalid_evidence(
            str(result.get("reason") or "引文证据校验失败"),
            evidence_valid=False,
            diagnostic_code=result.get("diagnostic_code") or "EVIDENCE_INVALID",
        )
    # Nested model evidence is flattened by normalize_claim_evidence_result;
    # validate the same normalized representation used for the claim mapping.
    raw_evidence = normalized.get("evidence", [])
    if not isinstance(raw_evidence, list):
        return _invalid_evidence("引文结构无效", diagnostic_code="EVIDENCE_INVALID")
    if not raw_evidence:
        if status != "NOT_EVALUATED":
            return _invalid_evidence("语义结论缺少原文引文", diagnostic_code="EVIDENCE_INVALID")
        normalized["evidence"] = []
        normalized["evidence_valid"] = True
        return normalized
    by_id = {str(paragraph.get("paragraph_id", "")): paragraph for paragraph in paragraphs}
    checked: list[dict[str, Any]] = []
    substantive = 0
    for citation in raw_evidence:
        if not isinstance(citation, dict):
            return _invalid_evidence("引文结构无效", diagnostic_code="EVIDENCE_INVALID")
        paragraph_id = str(citation.get("paragraph_id", ""))
        quote = str(citation.get("quote", ""))
        if not quote:
            return _invalid_evidence("引文为空", diagnostic_code="EVIDENCE_INVALID")
        paragraph = by_id.get(paragraph_id)
        if paragraph is None or quote not in str(paragraph.get("text", "")):
            if not allow_rebind:
                return _invalid_evidence("引文段落 ID 不属于原文或引文不在该段落中", diagnostic_code="EVIDENCE_INVALID")
            matches = [item for item in paragraphs if quote in str(item.get("text", ""))]
            if len(matches) != 1:
                return _invalid_evidence("缓存引文无法唯一重绑定到原文段落", diagnostic_code="EVIDENCE_INVALID")
            paragraph = matches[0]
            paragraph_id = str(paragraph.get("paragraph_id", ""))
        eligible = not bool(paragraph.get("is_toc")) and not bool(paragraph.get("is_title_repeat"))
        if eligible:
            substantive += 1
        checked.append({**citation, "paragraph_id": paragraph_id, "quote": quote, "eligible": eligible})
    if substantive == 0:
        return _invalid_evidence("目录或重复标题不能单独作为判定证据", diagnostic_code="EVIDENCE_INVALID")
    normalized["evidence"] = checked
    normalized["evidence_valid"] = True
    return normalized


def scan_title_body_inputs(
    root: str | Path,
    *,
    strategy_version: str = STRATEGY_VERSION,
    model_id: str = "",
    prompt_version: str = "",
    policy_version: str = "1",
    extra_dependencies: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Build the frozen local pairing index and extract every *说明.docx once."""
    material_root = Path(root).resolve()
    if not material_root.is_dir():
        raise NotADirectoryError(str(material_root))
    manuals = sorted(material_root.rglob("*说明.docx"), key=lambda path: path.as_posix().casefold())
    rows: list[dict[str, Any]] = []
    by_directory: dict[Path, list[Path]] = {}
    for manual in manuals:
        by_directory.setdefault(manual.parent, []).append(manual)

    pairings: dict[Path, tuple[Path | None, str, str]] = {}
    for directory, directory_manuals in by_directory.items():
        txt_files = sorted(
            (path for path in directory.iterdir() if path.is_file() and path.suffix.casefold() == ".txt"),
            key=lambda path: path.name.casefold(),
        )
        if not txt_files:
            for manual in directory_manuals:
                pairings[manual] = (None, "NOT_EVALUATED", "TXT_MISSING")
            continue
        if len(directory_manuals) == 1 and len(txt_files) == 1:
            manual = directory_manuals[0]
            txt = txt_files[0]
            directory_key = _normalise_name(directory.name)
            known_project_directory = directory != material_root and directory_key in {
                _normalise_name(txt.stem), _normalise_name(_manual_title_stem(manual)),
            }
            single_root_pair_matches = _normalise_name(txt.stem) == _normalise_name(_manual_title_stem(manual))
            if known_project_directory or single_root_pair_matches:
                pairings[manual] = (txt, "PAIRED", "已识别项目目录中的唯一配对")
            else:
                pairings[manual] = (None, "NOT_EVALUATED", "PAIRING_AMBIGUOUS")
            continue
        candidates: dict[Path, list[Path]] = {}
        claimed: dict[Path, int] = {}
        for manual in directory_manuals:
            key = _normalise_name(_manual_title_stem(manual))
            matches = [txt for txt in txt_files if _normalise_name(txt.stem) == key]
            candidates[manual] = matches
            for txt in matches:
                claimed[txt] = claimed.get(txt, 0) + 1
        for manual in directory_manuals:
            matches = candidates[manual]
            if len(matches) == 1 and claimed[matches[0]] == 1:
                pairings[manual] = (matches[0], "PAIRED", "规范化文件名主干一一对应")
            else:
                pairings[manual] = (None, "NOT_EVALUATED", "PAIRING_AMBIGUOUS")

    for manual in manuals:
        txt_path, pairing_status, pairing_reason = pairings[manual]
        relative_docx = manual.relative_to(material_root).as_posix()
        relative_txt = txt_path.relative_to(material_root).as_posix() if txt_path else None
        project_id = _project_id(relative_docx, relative_txt)
        raw_title = txt_path.stem if txt_path is not None else ""
        title = raw_title.strip()
        associated_files: list[str] = [relative_docx]
        if txt_path is not None:
            title_keys = {
                _normalise_name(title),
                _normalise_name(_manual_title_stem(manual)),
            }
            associated_files = [
                candidate.relative_to(material_root).as_posix()
                for candidate in sorted(manual.parent.iterdir(), key=lambda path: path.name.casefold())
                if candidate.is_file()
                and candidate.suffix.casefold() in {".txt", ".docx", ".pdf"}
                and _normalise_name(_association_stem(candidate)) in title_keys
            ]
            for required in (manual, txt_path):
                relative_required = required.relative_to(material_root).as_posix()
                if relative_required not in associated_files:
                    associated_files.append(relative_required)
        base = {
            "project_id": project_id,
            "status": "READY" if pairing_status == "PAIRED" else "NOT_EVALUATED",
            "title": title,
            "relative_path": relative_docx,
            "txt_relative_path": relative_txt,
            "associated_files": associated_files,
            "txt_path": str(txt_path) if txt_path else None,
            "docx_path": str(manual),
            "semantic_input_hash": None,
            "paragraphs": [],
            "metadata": {
                "project_id": project_id,
                "relative_path": relative_docx,
                "txt_relative_path": relative_txt,
                "raw_title": raw_title,
                "title_source": "txt_filename_stem" if txt_path else None,
                "pairing_status": pairing_status,
                "pairing_reason": pairing_reason,
                "reason": pairing_reason if pairing_status != "PAIRED" else "",
                "diagnostic_code": pairing_reason if pairing_reason in {"TXT_MISSING", "PAIRING_AMBIGUOUS"} else None,
                "docx_path": str(manual),
                "txt_path": str(txt_path) if txt_path else None,
                "associated_files": associated_files,
                "strategy_version": strategy_version,
                "model_id": model_id,
                "prompt_version": prompt_version,
                "policy_version": policy_version,
            },
        }
        if pairing_status != "PAIRED":
            rows.append(base)
            continue
        try:
            paragraphs = extract_docx_paragraphs(manual, title)
            base["paragraphs"] = paragraphs
            if not paragraphs:
                base["status"] = "NOT_EVALUATED"
                base["metadata"]["reason"] = "说明书没有可发送的正文文本"
                base["metadata"]["diagnostic_code"] = "BODY_TEXT_MISSING"
                base["metadata"]["pairing_status"] = "PAIRED"
            else:
                input_hash = semantic_input_hash(
                    title,
                    paragraphs,
                    strategy_version=strategy_version,
                    model_id=model_id,
                    prompt_version=prompt_version,
                    policy_version=policy_version,
                    extra_dependencies=extra_dependencies,
                )
                base["semantic_input_hash"] = input_hash
                base["metadata"]["semantic_input_hash"] = input_hash
        except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError, RuntimeError) as exc:
            base["status"] = "NOT_EVALUATED"
            base["metadata"]["reason"] = f"DOCX_OOXML_READ_FAILED: {type(exc).__name__}: {exc}"
            base["metadata"]["diagnostic_code"] = "DOCX_OOXML_READ_FAILED"
            base["metadata"]["pairing_status"] = "PAIRED"
        rows.append(base)
    return rows


def qualify_materials(
    root: str | Path,
    classify: Callable[[str, list[dict[str, Any]], dict[str, Any]], dict[str, Any]] | None = None,
    cache: TitleBodyCache | str | Path | None = None,
    progress_callback: Callable[[str, int, int, dict[str, Any]], None] | None = None,
    **scan_options: Any,
) -> list[dict[str, Any]]:
    """Scan once, use only injected classifiers, and validate all semantic evidence."""
    items = scan_title_body_inputs(root, **scan_options)
    cache_obj = cache if isinstance(cache, TitleBodyCache) else TitleBodyCache(cache) if cache is not None else None
    total = len(items)
    for index, item in enumerate(items, 1):
        if progress_callback is not None:
            progress_callback("active", index, total, item)
        if item["status"] != "READY" or not item["semantic_input_hash"]:
            diagnostic_code = item["metadata"].get("diagnostic_code")
            item["result"] = _not_evaluated_result(
                str(item["metadata"].get("reason") or "输入不满足 TXT 配对或正文条件"),
                str(diagnostic_code or "INPUT_NOT_READY"),
            )
            if progress_callback is not None:
                progress_callback("complete", index, total, item)
            continue
        input_hash = str(item["semantic_input_hash"])
        cached = cache_obj.get(input_hash, item["paragraphs"]) if cache_obj else None
        if cached is not None:
            item["result"] = cached
            item["cache_hit"] = True
            if progress_callback is not None:
                progress_callback("complete", index, total, item)
            continue
        if classify is None:
            item["result"] = _not_evaluated_result("未提供语义分类器", "CLASSIFIER_NOT_PROVIDED")
            item["cache_hit"] = False
            if progress_callback is not None:
                progress_callback("complete", index, total, item)
            continue
        try:
            classified = classify(item["title"], item["paragraphs"], item["metadata"])
            if not isinstance(classified, dict):
                raise TypeError("classifier must return a dictionary")
            raw_status = str(classified.get("status", "")).upper()
            if raw_status in {"UNKNOWN", "NOT_EVALUATED"}:
                validated = _not_evaluated_result(
                    str(classified.get("reason") or "输入未能完成评估"),
                    str(classified.get("diagnostic_code") or classified.get("failure") or "RESULT_NOT_EVALUATED"),
                    **{key: value for key, value in classified.items() if key not in {"status", "reason", "diagnostic_code", "failure", "core_claims", "evidence"}},
                )
            else:
                validated = validate_evidence_result(classified, item["paragraphs"])
            if validated.get("status") not in {"MATCH", "MISMATCH", "NOT_EVALUATED"}:
                validated = _not_evaluated_result(
                    str(validated.get("reason") or "分类器未能完成评估"),
                    str(validated.get("diagnostic_code") or "RESULT_NOT_EVALUATED"),
                )
        except Exception as exc:
            validated = _not_evaluated_result(
                f"语义分类失败：{type(exc).__name__}: {exc}", "CLASSIFIER_EXCEPTION",
            )
        item["result"] = validated
        item["cache_hit"] = False
        cacheable = (
            validated.get("status") in {"MATCH", "MISMATCH"}
            and (validated.get("evidence_valid") is True or validated.get("judgment_source") == "typesafe_jev_score")
        )
        if cache_obj and cacheable:
            try:
                cache_obj.put(input_hash, validated)
            except (OSError, TypeError, ValueError) as exc:
                item["cache_warning"] = f"CACHE_WRITE_FAILED: {type(exc).__name__}: {exc}"
        if progress_callback is not None:
            progress_callback("complete", index, total, item)
    return items


def index_by_relative_path(items: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Index scan results by their stable path relative to the supplied material root."""
    return {str(item["relative_path"]): item for item in items if item.get("relative_path")}


__all__ = [
    "STRATEGY_VERSION", "extract_docx_paragraphs", "qualify_materials",
    "index_by_relative_path", "normalize_claim_evidence_result", "scan_title_body_inputs", "validate_claim_evidence_structure",
    "validate_evidence_result",
]

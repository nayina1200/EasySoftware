from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any


RULES_PATH = Path(__file__).with_name("cover_rules.json")


@lru_cache(maxsize=1)
def load_rules() -> dict[str, Any]:
    return json.loads(RULES_PATH.read_text(encoding="utf-8"))


def normalize_cover_text(value: str | None) -> str:
    text = re.sub(r"\s+", "", value or "")
    for opening, closing in load_rules()["wrappers"]:
        if text.startswith(opening) and text.endswith(closing) and len(text) >= len(opening) + len(closing):
            text = text[len(opening):-len(closing)]
            break
    return text


def strip_cover_version(value: str | None) -> str:
    pattern = load_rules()["cover_version_pattern"]
    return re.sub(pattern, "", value or "")


def has_cover_version(value: str | None) -> bool:
    normalized = normalize_cover_text(value)
    return bool(normalized and strip_cover_version(normalized) != normalized)


def classify_cover_text(value: str | None, title: str | None = None) -> dict[str, str]:
    normalized = normalize_cover_text(value)
    result = {"normalized": normalized, "kind": "other", "reason": "未匹配完整软件名称或说明书标签"}
    if not normalized:
        result["reason"] = "内容为空"
        return result

    versionless = strip_cover_version(normalized)
    expected_title = strip_cover_version(normalize_cover_text(title)) if title else ""
    had_ordinal = False
    if expected_title:
        ordinal_pattern = load_rules()["cover_ordinal_pattern"]
        candidate, replacements = re.subn(ordinal_pattern, "", versionless, count=1)
        had_ordinal = bool(replacements)
        versionless = candidate
    if expected_title and versionless == expected_title:
        result["kind"] = "cover_title"
        reasons = []
        if had_ordinal:
            reasons.append("封面编号可移除")
        if strip_cover_version(normalized) != normalized:
            reasons.append("末尾版本号可移除")
        result["reason"] = f"与软件全称完整匹配（{'；'.join(reasons)}）" if reasons else "与软件全称完整匹配"
        return result

    rules = load_rules()
    excludes = [term for term in rules["exclude_terms"] if term in normalized]
    if excludes:
        result["kind"] = "excluded"
        result["reason"] = f"命中排除词“{excludes[0]}”"
        return result
    if any(mark in normalized for mark in rules["sentence_punctuation"]):
        result["reason"] = "含句末标点或文档编号字符，不是完整标签"
        return result
    if len(normalized) > int(rules["maximum_label_length"]):
        result["reason"] = "文本过长，不是完整封面标签"
        return result
    if normalized in rules["full_labels"]:
        result["kind"] = "manual_label"
        result["reason"] = f"完整标签“{normalized}”"
        return result

    for document_type in sorted(rules["document_types"], key=len, reverse=True):
        if not normalized.endswith(document_type):
            continue
        purpose = normalized[:-len(document_type)]
        for prefix in sorted(rules["allowed_prefixes"], key=len, reverse=True):
            if purpose.startswith(prefix):
                purpose = purpose[len(prefix):]
                break
        matched = next((term for term in rules["purpose_terms"] if purpose == term), None)
        if matched:
            result["kind"] = "manual_label"
            result["reason"] = f"用途词“{matched}”与类型词“{document_type}”完整匹配"
            return result
    result["reason"] = "缺少明确文档用途词"
    return result


def is_manual_label(value: str | None) -> bool:
    return classify_cover_text(value)["kind"] == "manual_label"


def is_cover_title(value: str | None, title: str | None) -> bool:
    return classify_cover_text(value, title)["kind"] == "cover_title"

"""Local product-name warnings over already extracted visible paragraphs."""
from __future__ import annotations

import re
import time
import unicodedata
from typing import Any, Iterable


RULE_VERSION = "software-name-consistency-v2"
_NAME = re.compile(r"[\u4e00-\u9fffA-Za-z0-9][\u4e00-\u9fffA-Za-z0-9·•_\- ]{1,70}?(?:系统|平台|软件|客户端|SDK)", re.I)
_VERSION = re.compile(r"\s*(?:v(?:ersion)?\s*\d+(?:\.\d+)*|版本\s*\d+(?:\.\d+)*)\s*$", re.I)
_NUMBERING = re.compile(r"^\s*(?:第[一二三四五六七八九十百\d]+[章节]|[一二三四五六七八九十\d]+(?:[.．]\d+)*[、.．)）\s]+)")
_INTRO = re.compile(r"欢迎使用|欢迎登录|软件名称为?|系统名称为?|平台名称为?|本软件名为|本系统名为|本平台名为|以下简称|简称为?|名为|登录|进入|打开|使用|对接|集成|接入|连接|调用|兼容|第三方|外部|例如|比如|依赖|同步至")
_CLAUSE = re.compile(r"[，,。；;！？!?\n]")
_DESCRIPTION = re.compile(r"能够|可以|用于|通过|实现|提供|包括|包含|支持|功能|运行于|兼容|负责|采用|利用|建议|适用于|聚焦|基于|构建|是一款|是一个|主要针对")
_GRAMMAR = re.compile(r"^(?:这套|这个|这些|这时|这是|那套|那个|那些|(?:本|该|此|各|整个|所有|上述|其他|其它|我们的)(?:系统|平台|软件|客户端)|没有|无论|启动|调整|用户在|作为|是|而|并|在|由|会|图\s*\d|\d)|所开发的|的系统$|的平台$|的软件$|等多个|将其|即可|只需|只要|点击|选择|需要|进行|显示|确保|了解|给各个|方便|把好|展示|选定|导入")
_REFERENCE = re.compile(r"对接|集成|接入|连接|调用|同步至|兼容|第三方|外部|例如|比如|诸如|依赖|操作系统|运行环境|(?:^|[（(、\s])如$")
_REFERENCE_AFTER = re.compile(r"^\s*(?:是|为|属于).{0,24}(?:第三方|外部|合作方|提供商)")
_MODULE = re.compile(r"^(?:模块|子系统|组件|接口)|(?:模块|子系统|组件|接口)(?:中|内|里|的|配置|设置)$|(?:模块|子系统|组件|接口)(?:名称)?\s*[:：]\s*$")
_NAVIGATION = re.compile(r"点击|菜单|按钮|面板|板块|工作台|功能入口|选项卡")
_SELF_BEFORE = re.compile(r"欢迎使用|欢迎登录|(?:软件|系统|平台)名称(?:为|是)?\s*[:：]?|(?:本软件|本系统|本平台)?名为")
_SELF_AFTER = re.compile(r"^\s*(?:是一款|是一个|是一套|是专为|是面向|为一款|专为)")
_SUFFIX = re.compile(r"(?:系统|平台|软件|客户端|sdk)$", re.I)
_ENVIRONMENT = re.compile(r"^(?:安卓|android|ios|windows|linux|unix|macos|鸿蒙|harmonyos|操作)系统$", re.I)
_GENERIC = {
    "系统", "平台", "软件", "sdk", "客户端", "业务系统", "管理系统", "应用系统",
    "操作系统", "信息系统", "智能化管理平台", "可视化平台", "控制系统", "底层系统",
    "综合控制系统", "综合管控平台", "传播平台", "游戏平台", "系统软件", "工业系统软件",
}


def normalize_name(name: str) -> str:
    name = _VERSION.sub("", unicodedata.normalize("NFKC", name).strip())
    return "".join(character for character in name.casefold() if character.isalnum())


def _quote(text: str, start: int, end: int) -> str:
    before = max((text.rfind(mark, 0, start) for mark in "。！？!?\n"), default=-1)
    after = [position for mark in "。！？!?\n" if (position := text.find(mark, end)) >= 0]
    return text[before + 1:min(after) + 1 if after else len(text)].strip()


def _possible_short_name(key: str, accepted: set[str]) -> bool:
    core = _SUFFIX.sub("", key)
    # Similar names remain review hints; they are never silently accepted aliases.
    return len(core) >= 3 and any(core in _SUFFIX.sub("", value) or _SUFFIX.sub("", value) in core for value in accepted)


def check_software_names(title: str, paragraphs: list[dict[str, Any]], *, aliases: Iterable[str] = ()) -> dict[str, Any]:
    started = time.perf_counter()
    accepted = {normalize_name(title)} if title.strip() else set()
    if not isinstance(aliases, (str, bytes, dict)):
        accepted.update(normalize_name(value) for value in aliases if isinstance(value, str) and value.strip())
    accepted.discard("")
    audit: dict[str, Any] = {
        "rule_version": RULE_VERSION, "status": "NOT_CHECKED", "title": title,
        "scope": "visible_docx_body_including_tables_and_headings", "image_text_checked": False,
        "warning_count": 0, "findings": [], "review_count": 0, "review_findings": [],
        "ignored_reference_count": 0, "filtered_candidate_count": 0,
        "scanned_paragraph_count": 0, "extra_model_requests": 0, "action": "REVIEW_ONLY",
    }
    if not accepted or not paragraphs:
        audit["reason"] = "缺少唯一配对 TXT 名称或可读取的正文，未完成名称检查"
        audit["elapsed_ms"] = round((time.perf_counter() - started) * 1000, 3)
        return audit
    # Compile once per document. Mask the entire accepted name before suffix scanning,
    # so a name containing both 平台 and 系统 cannot turn into two false candidates.
    gap = r"[\s·•_\-《》“”\"'（）()]*"
    accepted_pattern = re.compile("|".join(gap.join(map(re.escape, value)) for value in sorted(accepted, key=len, reverse=True)), re.I)
    groups: dict[str, dict[str, Any]] = {}
    for paragraph_index, paragraph in enumerate(paragraphs):
        text = str(paragraph.get("text") or "")
        if not text or paragraph.get("is_toc"):
            continue
        audit["scanned_paragraph_count"] += 1
        searchable = accepted_pattern.sub(lambda match: " " * len(match.group()), text)
        for match in _NAME.finditer(searchable):
            raw = match.group()
            candidate = _NUMBERING.sub("", raw).strip()
            introductions = list(_INTRO.finditer(candidate))
            if introductions:
                candidate = candidate[introductions[-1].end():].strip()
            key = normalize_name(candidate)
            if (len(key) < 4 or key in accepted or key in _GENERIC or _ENVIRONMENT.fullmatch(key)
                    or _DESCRIPTION.search(candidate) or _GRAMMAR.search(candidate) or candidate.endswith("子系统")):
                audit["filtered_candidate_count"] += 1
                continue
            start = match.start() + raw.rfind(candidate)
            end = start + len(candidate)
            left = _CLAUSE.split(text[max(0, start - 80):start])[-1]
            right = text[end:end + 32]
            if (_REFERENCE.search(left) or _REFERENCE_AFTER.search(right)
                    or _MODULE.search(right) or _MODULE.search(left)):
                audit["ignored_reference_count"] += 1
                continue
            short_name = _possible_short_name(key, accepted)
            explicit_self = bool(_SELF_BEFORE.search(left) or _SELF_AFTER.search(right))
            standalone = normalize_name(text.strip(" \t\r\n《》“”\"'")) == key
            # An early, standalone software title or a versioned product heading is
            # identity evidence. An arbitrary 系统 heading can be a feature module.
            identity_heading = standalone and (paragraph_index < 12 and key.endswith(("软件", "客户端")) or bool(_VERSION.search(text)))
            module_context = bool(_NAVIGATION.search(left) or right.lstrip().startswith(("：", ":")))
            if module_context and not explicit_self and not short_name:
                audit["filtered_candidate_count"] += 1
                continue
            strong = (explicit_self or identity_heading) and not short_name and not module_context
            if not strong and not (short_name or standalone or explicit_self):
                audit["filtered_candidate_count"] += 1
                continue
            role, priority = ("self_name", "high") if strong else ("possible_alias" if short_name else "uncertain_name", "review")
            reason = "正文明确使用另一软件名称，与正式名称不一致" if strong else (
                "可能是正式名称的简称或名称变体，待确认" if short_name else "可能是功能模块或产品名称，证据不足，待确认"
            )
            group = groups.setdefault(key, {"name": candidate, "role": role, "priority": priority, "reason": reason, "occurrences": []})
            if strong:
                group.update(role=role, priority=priority, reason=reason)
            occurrence = {
                "paragraph_id": str(paragraph.get("paragraph_id", "")), "section_title": str(paragraph.get("section_title", "")),
                "in_table": bool(paragraph.get("in_table")), "quote": _quote(text, start, end), "source": "docx_body",
            }
            if occurrence not in group["occurrences"]:
                group["occurrences"].append(occurrence)
    for group in groups.values():
        group["occurrence_count"] = len(group["occurrences"])
        audit["findings" if group["priority"] == "high" else "review_findings"].append(group)
    audit["warning_count"] = len(audit["findings"])
    audit["review_count"] = len(audit["review_findings"])
    audit["status"] = "REVIEW" if audit["findings"] else "PENDING_REVIEW" if audit["review_findings"] else "NO_CONFLICT_FOUND"
    audit["reason"] = "发现有明确证据的名称冲突，请复核" if audit["findings"] else "没有明确名称冲突；待确认项可展开查看" if audit["review_findings"] else "本地规则未发现名称冲突；不代表完整语义验收"
    audit["elapsed_ms"] = round((time.perf_counter() - started) * 1000, 3)
    return audit


def render_name_warnings(projects: list[dict[str, Any]]) -> str:
    import html

    def row(project: dict[str, Any], audit: dict[str, Any], finding: dict[str, Any]) -> str:
        evidence = "<br>".join(html.escape("；".join(filter(None, [
            str(occurrence.get("section_title", "")), str(occurrence.get("paragraph_id", "")), str(occurrence.get("quote", "")),
        ]))) for occurrence in finding.get("occurrences", []))
        return "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
            html.escape(str(audit.get("title") or project.get("name", ""))), html.escape(str(finding.get("name", ""))),
            html.escape(str(finding.get("reason", ""))), evidence,
        )

    warnings, pending = [], []
    checked = 0
    for project in projects:
        audit = project.get("name_consistency") or {}
        checked += audit.get("status") in {"REVIEW", "PENDING_REVIEW", "NO_CONFLICT_FOUND"}
        for finding in audit.get("findings", []):
            # Older report data may still contain review hints in findings.
            (warnings if finding.get("priority") == "high" else pending).append(row(project, audit, finding))
        pending.extend(row(project, audit, finding) for finding in audit.get("review_findings", []))
    header = "<table><thead><tr><th>正式名称</th><th>正文中的名称</th><th>提示原因</th><th>原文与段落位置</th></tr></thead><tbody>"
    if not checked:
        return "<p>未完成软件名称检查：未选择名称检查或无法配对/读取正文的材料不计为已检查。</p>"
    summary = "<p>已检查 {} 份；未完成 {} 份。图片内名称未检查。</p>".format(checked, len(projects) - checked)
    if warnings:
        summary += '<div class="name-alert"><b>软件名称疑似冲突：{} 处有明确证据，请复核</b><p>本检查只报警，不修改或跳过材料。</p></div>'.format(len(warnings))
        summary += header + "".join(warnings) + "</tbody></table>"
    else:
        summary += "<p>未发现有明确证据的软件名称冲突。</p>"
    if pending:
        summary += "<details><summary>待确认名称 {} 项（可能为简称或模块，不计入报警）</summary>".format(len(pending))
        summary += header + "".join(pending) + "</tbody></table></details>"
    return summary

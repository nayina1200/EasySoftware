from __future__ import annotations

import ctypes
import difflib
import html
import io
import json
import os
import re
import shutil
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import archive_types
import code_comment_cleaner
import manual_content_stage
import process_batch
from software_name_consistency import check_software_names, render_name_warnings


VERSION = 1
PREVIEW_STATUS = "NEEDS_DOCUMENT_PROCESSING_REVIEW"
WORKFLOW_OPTION_ORDER = (
    "match_title_body", "match_software_name", "match_manual_ocr", "match_manual_vision", "match_code", "match_txt",
    "manual_body_years", "manual_image_dates", "manual_ai_hollow_lines", "manual_cover", "manual_header", "manual_page_numbers", "manual_pdf",
    "code_cleanup", "code_pdf", "txt_repair",
)
WORKFLOW_GROUPS = {
    "all": WORKFLOW_OPTION_ORDER,
    "checks": tuple(item for item in WORKFLOW_OPTION_ORDER if item.startswith("match_")),
    "manual": tuple(option for option in WORKFLOW_OPTION_ORDER if option in {
        "match_title_body", "match_software_name", "match_manual_ocr", "match_manual_vision"
    } or option.startswith("manual_")),
    "code": ("match_code", "code_cleanup", "code_pdf"),
    "txt": ("match_txt", "txt_repair"),
}
DEFAULT_WORKFLOW_OPTIONS = WORKFLOW_OPTION_ORDER
WORKFLOW_OPTION_LABELS = {
    "match_title_body": "标题与说明书正文主题匹配",
    "match_software_name": "软件名称一致性检查（本地）",
    "match_manual_ocr": "截图 OCR 与软件主题匹配",
    "match_manual_vision": "视觉模型复核截图",
    "match_code": "代码 DOCX/PDF 配套检查",
    "match_txt": "TXT 名称、字段与主题检查",
    "manual_body_years": "检查/修正说明书正文年份",
    "manual_image_dates": "检查/修正说明书图片日期",
    "manual_ai_hollow_lines": "检测空心线",
    "manual_cover": "整理说明书封面",
    "manual_header": "整理说明书页眉",
    "manual_page_numbers": "整理说明书页码",
    "manual_pdf": "导出说明书 PDF",
    "code_cleanup": "翻译代码中文注释并尽量保持页数",
    "code_pdf": "导出代码 PDF",
    "txt_repair": "修复 TXT 主要功能字段",
}


def _expand_workflow_options(stages: str | list[str] | tuple[str, ...] | None) -> tuple[str, ...]:
    if stages is None:
        return DEFAULT_WORKFLOW_OPTIONS
    requested = re.split(r"[,\s]+", stages.strip()) if isinstance(stages, str) else stages
    selected: set[str] = set()
    for raw in requested:
        stage = str(raw).strip()
        if stage in WORKFLOW_GROUPS:
            selected.update(WORKFLOW_GROUPS[stage])
        elif stage == "txt_placeholder":
            selected.update(WORKFLOW_GROUPS["txt"])
        elif stage:
            selected.add(stage)
    unknown = selected.difference(WORKFLOW_OPTION_ORDER)
    if unknown:
        raise ValueError(f"未知流程选项：{', '.join(sorted(unknown))}")
    return tuple(option for option in WORKFLOW_OPTION_ORDER if option in selected)


def _selected_stages(stages: str | list[str] | tuple[str, ...] | None) -> tuple[str, ...]:
    return _expand_workflow_options(stages)


def changes_html(changes: list[dict[str, Any]], heading: str) -> str:
    if not changes:
        return ""
    rows = []
    context_size = 24
    for item in changes:
        before = str(item.get("before", ""))
        after = str(item.get("after", ""))
        for tag, before_start, before_end, after_start, after_end in difflib.SequenceMatcher(None, before, after).get_opcodes():
            if tag == "equal":
                continue

            def snippet(value: str, start: int, end: int) -> str:
                left = max(0, start - context_size)
                right = min(len(value), end + context_size)
                prefix = "…" if left else ""
                suffix = "…" if right < len(value) else ""
                changed = value[start:end]
                if len(changed) > 120:
                    omitted = len(changed) - 80
                    changed = f"{changed[:40]}…{changed[-40:]}（修改段共 {len(value[start:end])} 字，中间省略 {omitted} 字）"
                changed_html = f"<mark>{html.escape(changed)}</mark>" if changed else "<mark>（无）</mark>"
                return prefix + html.escape(value[left:start]) + changed_html + html.escape(value[end:right]) + suffix

            rows.append(
                '<tr><td>{}</td><td>{}</td></tr>'.format(
                    snippet(before, before_start, before_end),
                    snippet(after, after_start, after_end),
                )
            )
    if not rows:
        return ""
    return (
        f'<section class="txt-diff"><h4>{html.escape(heading)}</h4><p>仅显示修改点及附近上下文，黄色部分为具体差异。</p>'
        f'<table><thead><tr><th>修改前</th><th>修改后</th></tr></thead><tbody>{"".join(rows)}</tbody></table></section>'
    )


def _publish_progress(
    progress: Any,
    phase: str,
    current: int,
    name: str = "",
    *,
    status: str = "RUNNING",
    failures: int = 0,
    error: str | None = None,
    active: bool | None = None,
) -> None:
    if progress is None:
        return
    details = {"status": status, "failures": failures}
    if error:
        details["error"] = error
    if active is not None:
        details["active"] = active
    target = progress.update if hasattr(progress, "update") else progress
    try:
        target(phase, current, name, **details)
    except TypeError:
        target(phase, current, name)


class _StageProgress:
    def __init__(self, progress: Any, offset: int, failures: int = 0, floor_current: int = 0):
        self.progress = progress
        self.offset = offset
        self.failures = failures
        self.floor_current = floor_current

    def update(self, phase: str, current: int = 0, name: str = "", **details: Any) -> None:
        _publish_progress(
            self.progress,
            phase,
            max(self.floor_current, self.offset + current),
            name,
            status=str(details.get("status", "RUNNING")),
            failures=max(self.failures, int(details.get("failures", 0))),
            error=details.get("error"),
            active=details.get("active"),
        )


def _material_root(source: Path, work: Path, app_root: Path, *, expand_order_downloads: bool = False) -> tuple[Path, str]:
    source = source.resolve()
    if source.is_file():
        if not archive_types.is_archive_file(source):
            raise ValueError("材料输入需要是压缩包（RAR/ZIP/7z 等）或已解压的材料文件夹")
        extracted = code_comment_cleaner.extract_cached(source, work, app_root)
        root = code_comment_cleaner.material_root(extracted)
        nested_state_path = work / "combined_nested_archives.json"
        try:
            prior = json.loads(nested_state_path.read_text(encoding="utf-8")) if nested_state_path.is_file() else []
        except (OSError, json.JSONDecodeError):
            prior = []
        hashes = process_batch.HashCache(work / "combined_hashes.json")
        nested = process_batch.expand_nested_archives(
            root,
            process_batch.find_7z(),
            max(1, min(4, os.cpu_count() or 1)),
            hashes,
            prior_results=prior,
        )
        hashes.save()
        nested_state_path.write_text(json.dumps(nested, ensure_ascii=False, indent=2), encoding="utf-8")
        return code_comment_cleaner.material_root(root), source.stem
    if source.is_dir():
        if "_已处理_" in source.name:
            raise ValueError("不能把已处理目录再次作为输入")
        if expand_order_downloads:
            nested_state_path = work / "combined_nested_archives.json"
            try:
                prior = json.loads(nested_state_path.read_text(encoding="utf-8")) if nested_state_path.is_file() else []
            except (OSError, json.JSONDecodeError):
                prior = []
            hashes = process_batch.HashCache(work / "combined_hashes.json")
            nested = process_batch.expand_nested_archives(
                source, process_batch.find_7z(), max(1, min(4, os.cpu_count() or 1)), hashes, prior_results=prior,
            )
            hashes.save()
            nested_state_path.write_text(json.dumps(nested, ensure_ascii=False, indent=2), encoding="utf-8")
        return source, source.name
    raise FileNotFoundError(f"输入不存在：{source}")


def _is_excluded_project(path: Path, root: Path, excluded: list[dict[str, Any]]) -> bool:
    relative = path.relative_to(root)
    normalized = relative.as_posix().casefold()
    stem = path.stem.removesuffix("说明").removesuffix("代码").casefold()
    for project in excluded:
        exact_files = project.get("associated_files") or project.get("files")
        if exact_files is not None:
            normalized_files = {
                str(item).replace("\\", "/").strip("/").casefold()
                for item in exact_files
            }
            if normalized in normalized_files:
                return True
            continue
        folder = project.get("folder", "").strip("./\\").replace("\\", "/").casefold()
        name = project.get("name", "").casefold()
        if name and stem == name:
            return True
        if folder and (normalized == folder or normalized.startswith(folder + "/")):
            return True
    return False


_MISMATCH_CLIPBOARD_HEADING = "其余材料正在处理，以下材料不匹配："


def _mismatch_names(projects: list[dict[str, Any]]) -> list[str]:
    names: list[str] = []
    seen: set[str] = set()
    for item in projects:
        if not isinstance(item, dict):
            continue
        statuses = {
            str(item.get("status") or "").strip().upper(),
            str(item.get("semantic_status") or "").strip().upper(),
        }
        if "MISMATCH" not in statuses:
            continue
        name = str(item.get("name") or "").strip()
        key = name.casefold()
        if name and key not in seen:
            names.append(name)
            seen.add(key)
    return names


def _format_mismatch_clipboard(names: list[str]) -> str:
    return f"{_MISMATCH_CLIPBOARD_HEADING}\n" + "\n".join(names) if names else ""


def _write_windows_clipboard(text: str) -> None:
    """Write Unicode text to the Windows clipboard using only the standard library."""
    if os.name != "nt":
        raise OSError("Windows 剪贴板仅在 Windows 上可用")

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    user32.CreateWindowExW.argtypes = [
        ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
    ]
    user32.CreateWindowExW.restype = ctypes.c_void_p
    user32.DestroyWindow.argtypes = [ctypes.c_void_p]
    user32.DestroyWindow.restype = ctypes.c_int
    user32.OpenClipboard.argtypes = [ctypes.c_void_p]
    user32.OpenClipboard.restype = ctypes.c_int
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = ctypes.c_int
    user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
    user32.SetClipboardData.restype = ctypes.c_void_p
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = ctypes.c_int
    kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.restype = ctypes.c_int
    kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
    kernel32.GlobalFree.restype = ctypes.c_void_p

    encoded = text.encode("utf-16-le") + b"\x00\x00"
    memory = kernel32.GlobalAlloc(0x0002 | 0x0040, len(encoded))  # GMEM_MOVEABLE | GMEM_ZEROINIT
    if not memory:
        raise ctypes.WinError(ctypes.get_last_error())
    transferred = False
    owner = None
    opened = False
    operation_error: Exception | None = None
    cleanup_errors: list[str] = []
    try:
        pointer = kernel32.GlobalLock(memory)
        if not pointer:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            ctypes.memmove(pointer, encoded, len(encoded))
        finally:
            ctypes.set_last_error(0)
            still_locked = kernel32.GlobalUnlock(memory)
            unlock_error = ctypes.get_last_error()
            if still_locked:
                raise OSError("剪贴板文本内存块解锁后仍处于锁定状态")
            if unlock_error:
                raise ctypes.WinError(unlock_error)

        # EmptyClipboard assigns ownership to the window associated with the
        # open clipboard. A NULL owner makes the following SetClipboardData fail.
        owner = user32.CreateWindowExW(
            0, "STATIC", "EasySoftware clipboard owner", 0x80000000,
            0, 0, 0, 0, None, None, None, None,
        )
        if not owner:
            raise ctypes.WinError(ctypes.get_last_error())
        for attempt in range(4):
            if user32.OpenClipboard(owner):
                opened = True
                break
            if attempt < 3:
                time.sleep(0.05 * (attempt + 1))
        if not opened:
            raise ctypes.WinError(ctypes.get_last_error())
        if not user32.EmptyClipboard():
            raise ctypes.WinError(ctypes.get_last_error())
        if not user32.SetClipboardData(13, memory):  # CF_UNICODETEXT
            raise ctypes.WinError(ctypes.get_last_error())
        transferred = True  # Windows owns the allocation after this call.
    except Exception as exc:
        operation_error = exc
    finally:
        if opened and not user32.CloseClipboard():
            cleanup_errors.append(f"CloseClipboard: {ctypes.WinError(ctypes.get_last_error())}")
        if owner and not user32.DestroyWindow(owner):
            cleanup_errors.append(f"DestroyWindow: {ctypes.WinError(ctypes.get_last_error())}")
        if not transferred:
            if kernel32.GlobalFree(memory):
                cleanup_errors.append(f"GlobalFree: {ctypes.WinError(ctypes.get_last_error())}")
    if operation_error is not None:
        if cleanup_errors:
            raise OSError(f"{operation_error}; 清理失败：{'；'.join(cleanup_errors)}") from operation_error
        raise operation_error
    if cleanup_errors:
        raise OSError("剪贴板操作完成后的清理失败：" + "；".join(cleanup_errors))


def _copy_mismatch_summary(
    names: list[str], *, apply: bool, clipboard_writer: Any = None,
) -> dict[str, Any]:
    text = _format_mismatch_clipboard(names)
    result: dict[str, Any] = {
        "names": list(names),
        "text": text,
        "attempted": False,
        "status": "no_mismatch" if not names else "preview_not_copied",
        "failure": False,
        "error": None,
    }
    if not names or not apply:
        return result
    result["attempted"] = True
    try:
        (clipboard_writer or _write_windows_clipboard)(text)
        result["status"] = "copied"
    except Exception as exc:
        result["status"] = "failed"
        result["failure"] = True
        result["error"] = f"{type(exc).__name__}: {exc}"[:400]
    return result


def _title_body_config(app_root: Path) -> tuple[Path, dict[str, Any]]:
    configured = app_root / "title_body_config.json"
    example_candidates = (
        app_root / "title_body_config.example.json",
        app_root / "scripts" / "title_body_config.example.json",
    )
    example = next((path for path in example_candidates if path.is_file()), example_candidates[-1])
    config_path = configured if configured.is_file() else example
    if not config_path.is_file():
        return config_path, {"mode": "enforced", "skip_mismatch_requested": True, "not_evaluated_action": "hold_changes"}
    try:
        value = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return config_path, {"mode": "enforced", "skip_mismatch_requested": True, "not_evaluated_action": "hold_changes", "config_error": "标题正文配置无法读取"}
    if not isinstance(value, dict):
        return config_path, {"mode": "enforced", "skip_mismatch_requested": True, "not_evaluated_action": "hold_changes", "config_error": "标题正文配置格式无效"}
    mode = str(value.get("mode", "enforced")).casefold()
    if mode not in {"off", "shadow", "enforced"}:
        mode = "enforced"
        value["config_error"] = "无效模式已安全回退到 enforced"
    value["mode"] = mode
    value["not_evaluated_action"] = value.get("not_evaluated_action", value.get("unknown_action", "hold_changes"))
    value.pop("unknown_action", None)
    if value["not_evaluated_action"] not in {"hold_changes", "continue_with_review"}:
        value["not_evaluated_action"] = "hold_changes"
        value["config_error"] = value.get("config_error") or "无效 not_evaluated_action 已安全回退到 hold_changes"
    if not isinstance(value.get("skip_mismatch_requested", True), bool):
        value["skip_mismatch_requested"] = True
        value["config_error"] = value.get("config_error") or "无效 skip_mismatch_requested 已安全回退为 true"
    value.setdefault("prompt_version", "title-body-jev-domain-fit-score-v3")
    value.setdefault("policy_version", "title-body-jev-threshold-policy-v1")
    provider = str(value.get("provider", "typesafe_jev")).casefold()
    if provider not in {"typesafe_jev", "local_llama"}:
        value["config_error"] = value.get("config_error") or "无效 provider 已安全回退到 typesafe_jev"
        provider = "typesafe_jev"
    value["provider"] = provider
    return config_path, value


def _source_name_audit(source: dict[str, Any], config: dict[str, Any], enabled: bool) -> dict[str, Any]:
    if not enabled:
        return {"status": "NOT_CHECKED", "warning_count": 0, "findings": [], "reason": "未选择软件名称一致性检查"}
    alias_map = config.get("software_name_aliases", {})
    aliases = alias_map.get(source.get("title", ""), []) if isinstance(alias_map, dict) else []
    return check_software_names(
        str(source.get("title") or ""), source.get("paragraphs", []),
        aliases=aliases if isinstance(aliases, list) else [],
    )


def _scan_title_body_qualification(
    root: Path,
    progress: Any = None,
    *,
    work_path: Path | None = None,
    app_root: Path | None = None,
    config_path: Path | None = None,
    progress_current: int | None = None,
    check_names: bool = True,
    local_only: bool = False,
) -> dict[str, Any]:
    """Qualify paired TXT/DOCX materials once, without loading images or Word."""
    root = root.resolve()
    app_root = (app_root or SCRIPT_DIR.parent).resolve()
    work_path = (work_path or app_root / ".title-body-work").resolve()
    if config_path is None:
        config_path, config = _title_body_config(app_root)
    else:
        _loaded_path, config = _title_body_config(app_root)
        config_path = config_path.resolve()
        try:
            raw_config = json.loads(config_path.read_text(encoding="utf-8"))
            if isinstance(raw_config, dict):
                config = raw_config
        except (OSError, UnicodeError, json.JSONDecodeError):
            config.setdefault("config_error", "标题正文配置无法读取")
    mode = str(config.get("mode", "enforced")).casefold()
    if mode not in {"off", "shadow", "enforced"}:
        mode = "enforced"
        config["config_error"] = config.get("config_error") or "无效模式已安全回退到 enforced"
    config["mode"] = mode
    config["not_evaluated_action"] = config.get("not_evaluated_action", config.get("unknown_action", "hold_changes"))
    if config["not_evaluated_action"] not in {"hold_changes", "continue_with_review"}:
        config["not_evaluated_action"] = "hold_changes"
        config["config_error"] = config.get("config_error") or "无效 not_evaluated_action 已安全回退到 hold_changes"
    if not isinstance(config.get("skip_mismatch_requested", True), bool):
        config["skip_mismatch_requested"] = True
        config["config_error"] = config.get("config_error") or "无效 skip_mismatch_requested 已安全回退为 true"
    config.setdefault("prompt_version", "title-body-jev-domain-fit-score-v3")
    config.setdefault("policy_version", "title-body-jev-threshold-policy-v1")
    provider = str(config.get("provider", "typesafe_jev")).casefold()
    if provider not in {"typesafe_jev", "local_llama"}:
        config["config_error"] = config.get("config_error") or "无效 provider 已安全回退到 typesafe_jev"
        provider = "typesafe_jev"
    config["provider"] = provider
    if local_only:
        from title_body_qualification import scan_title_body_inputs
        sources = scan_title_body_inputs(root)
        projects = []
        for source in sources:
            projects.append({
                "name": source.get("title") or Path(source["relative_path"]).stem.removesuffix("说明"),
                "relative_path": source["relative_path"],
                "status": "NOT_RUN", "semantic_status": "NOT_RUN",
                "classification_method": "local_name_rules",
                "name_consistency": _source_name_audit(source, config, check_names),
                "reason": "本次仅运行本地名称检查，未选择主题评分",
                "action": "CONTINUE_WITH_REVIEW", "cache_hit": False,
                "result": {}, "associated_files": source.get("associated_files", []),
            })
        return {
            "status": "COMPLETED", "mode": mode, "provider": "local_name_rules", "model_id": None,
            "scanned_count": len(projects), "cache_hit_count": 0, "model_request_count": 0,
            "failure_count": 0, "failures": [], "projects": projects,
            "name_warning_count": sum(item["name_consistency"]["warning_count"] for item in projects),
            "name_warning_names": [item["name"] for item in projects if item["name_consistency"]["warning_count"]],
            "disqualified_projects": [], "_inputs_by_relative_path": {},
        }
    if mode == "off":
        from title_body_qualification import scan_title_body_inputs
        local_inputs = {
            item["relative_path"]: item for item in scan_title_body_inputs(root)
        } if check_names else {}
        projects: list[dict[str, Any]] = []
        legacy_mismatches: list[dict[str, Any]] = []
        for index, manual in enumerate(sorted(root.rglob("*说明.docx"), key=lambda path: path.as_posix().casefold()), 1):
            name = manual.stem.removesuffix("说明")
            try:
                paragraphs = manual_content_stage.read_docx_paragraphs(manual)
                legacy = manual_content_stage.assess_body_relevance(name, paragraphs)
            except Exception as exc:
                legacy = {"status": "unknown", "reason": f"无法读取正文：{type(exc).__name__}: {exc}"}
            local_source = local_inputs.get(manual.relative_to(root).as_posix(), {})
            name_audit = _source_name_audit(local_source, config, check_names)
            item = {
                "name": name,
                "relative_path": manual.relative_to(root).as_posix(),
                "status": "NOT_RUN",
                "semantic_status": "NOT_RUN",
                "classification_method": "legacy_lexical",
                "name_consistency": name_audit,
                "legacy_status": str(legacy.get("status", "unknown")),
                "legacy_similarity_score": legacy.get("similarity_score"),
                "reason": str(legacy.get("reason", "")),
                "action": "CONTINUE_WITH_REVIEW",
                "associated_files": [],
                "cache_hit": False,
                "result": {"status": "NOT_EVALUATED", "reason": "标题正文评分已关闭"},
            }
            projects.append(item)
            if item["legacy_status"] == "review_topic_mismatch":
                legacy_mismatches.append(item)
            if progress is not None and progress_current is None:
                _publish_progress(progress, "标题正文旧规则预检（语义判定已关闭）", index, name, active=False)
        if progress is not None and progress_current is not None:
            _publish_progress(progress, "标题正文旧规则预检完成", progress_current, "已将结果交给后续阶段", active=False)
        return {
            "status": "COMPLETED", "mode": mode, "model_id": None,
            "config_path": str(config_path), "config_error": config.get("config_error"),
            "scanned_count": len(projects), "cache_hit_count": 0, "model_request_count": 0,
            "failure_count": 0, "failures": [], "projects": projects,
            "name_warning_count": sum(item["name_consistency"]["warning_count"] for item in projects),
            "name_warning_names": [item["name"] for item in projects if item["name_consistency"]["warning_count"]],
            "disqualified_projects": [], "legacy_conflict_projects": legacy_mismatches,
            "_inputs_by_relative_path": {},
        }

    from title_body_cache import TitleBodyCache
    from title_body_qualification import STRATEGY_VERSION, qualify_materials

    provider = str(config.get("provider", "typesafe_jev")).casefold()
    if provider == "local_llama":
        from local_model_runtime import LocalModelRuntime as Runtime, SELECTION_POLICY_VERSION
    else:
        from typesafe_jev_runtime import TypeSafeJevRuntime as Runtime, SELECTION_POLICY_VERSION
        provider = "typesafe_jev"

    runtime = Runtime(config_path, app_root=app_root)
    identity_started = time.monotonic()
    identity_provider = getattr(runtime, "cache_identity", None)
    model_identity = identity_provider() if callable(identity_provider) else None
    cache_identity_elapsed_ms = int((time.monotonic() - identity_started) * 1000)
    cache = TitleBodyCache(work_path / "title_body_cache" / "v1.json") if model_identity else None
    runtime_start_attempted = False
    service_failures = 0
    paused_diagnostic: str | None = None
    service_failure_codes = {"MODEL_UNAVAILABLE", "MODEL_INTERFACE_ERROR", "REQUEST_TIMEOUT", "CANCEL_UNCONFIRMED"}

    def classify_once(title: str, paragraphs: list[dict[str, Any]], metadata: dict[str, Any]) -> dict[str, Any]:
        nonlocal runtime_start_attempted, service_failures, paused_diagnostic
        if paused_diagnostic:
            return {
                "status": "NOT_EVALUATED", "diagnostic_code": paused_diagnostic,
                "failure": paused_diagnostic, "reason": "连续服务级故障后已暂停本批后续语义请求",
                "core_claims": [], "evidence": [],
                "diagnostics": {"generation_requests": 0, "circuit_breaker": "OPEN"},
            }
        if not runtime_start_attempted:
            start_status = runtime.start()
            runtime_start_attempted = True
            if start_status.state != "ready":
                paused_diagnostic = start_status.diagnostic_code or "MODEL_UNAVAILABLE"
                return {
                    "status": "NOT_EVALUATED", "diagnostic_code": paused_diagnostic,
                    "failure": paused_diagnostic, "reason": start_status.reason or f"{provider} 运行时不可用",
                    "core_claims": [], "evidence": [],
                    "diagnostics": {"generation_requests": 0, "circuit_breaker": "OPEN"},
                }
        result = runtime.classify(title, paragraphs, metadata)
        diagnostic = str(result.get("diagnostic_code") or "") if isinstance(result, dict) else ""
        if diagnostic in service_failure_codes:
            service_failures += 1
            if service_failures >= 3:
                paused_diagnostic = diagnostic
        else:
            service_failures = 0
        return result

    if progress is not None:
        _publish_progress(
            progress, f"标题正文资格预检：准备 {provider}", 0, "读取 TXT 配对与说明书正文",
            active=True,
        )

    progress_failure_count = 0

    def qualification_progress(phase: str, index: int, total_count: int, source_item: dict[str, Any]) -> None:
        nonlocal progress_failure_count
        if progress is None:
            return
        name = str(source_item.get("title") or Path(str(source_item.get("relative_path", ""))).stem)
        if phase == "complete":
            result = source_item.get("result", {})
            if isinstance(result, dict) and (result.get("failure") or result.get("diagnostic_code")):
                progress_failure_count += 1
        current = progress_current if progress_current is not None else index - (phase == "active")
        active = phase == "active" or progress_current is not None
        label = f"{index}/{total_count} {name}"
        _publish_progress(
            progress, "标题正文资格预检", int(current), label,
            failures=progress_failure_count, active=active,
        )

    try:
        items = qualify_materials(
            root,
            classify=classify_once,
            cache=cache,
            strategy_version=STRATEGY_VERSION,
            model_id=runtime.model_id,
            prompt_version=str(config.get("prompt_version", "title-body-jev-domain-fit-score-v3")),
            policy_version=str(config.get("policy_version", "title-body-jev-threshold-policy-v1")),
            progress_callback=qualification_progress,
            extra_dependencies={
                "mode": mode,
                "runtime": config.get("runtime", {}) if provider == "local_llama" else {
                    key: config.get("typesafe_jev", {}).get(key)
                    for key in ("model_id", "request_timeout_seconds", "max_retries", "max_retry_after_seconds",
                                "max_request_characters")
                    if isinstance(config.get("typesafe_jev", {}), dict) and key in config.get("typesafe_jev", {})
                },
                "provider": provider,
                "selection_policy_version": SELECTION_POLICY_VERSION,
                "coverage_policy": config.get("coverage_policy", {}),
                "model_identity": model_identity,
            },
        )
        runtime_status = runtime.status
    finally:
        runtime.close()

    projects: list[dict[str, Any]] = []
    mismatches: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    inputs_by_relative_path: dict[str, dict[str, Any]] = {}
    for index, source_item in enumerate(items, 1):
        result = source_item.get("result") if isinstance(source_item.get("result"), dict) else {}
        status = str(result.get("status", "NOT_EVALUATED")).upper()
        relative_path = str(source_item.get("relative_path", ""))
        name = str(source_item.get("title") or Path(relative_path).stem.removesuffix("说明"))
        diagnostic = str(result.get("diagnostic_code") or "")
        associated_files = list(source_item.get("associated_files", []))
        source_item["name_consistency"] = _source_name_audit(source_item, config, check_names)
        if mode != "off" and not associated_files and relative_path:
            associated_files = [relative_path]
        item = {
            "project_id": source_item.get("project_id"),
            "name": name,
            "status": status if status in {"MATCH", "MISMATCH", "NOT_EVALUATED"} else "NOT_EVALUATED",
            "semantic_status": status if status in {"MATCH", "MISMATCH", "NOT_EVALUATED"} else "NOT_EVALUATED",
            "classification_method": provider,
            "reason": str(result.get("reason", source_item.get("metadata", {}).get("reason", ""))),
            "relative_path": relative_path,
            "txt_relative_path": source_item.get("txt_relative_path"),
            "title_source": source_item.get("metadata", {}).get("title_source"),
            "pairing_status": source_item.get("metadata", {}).get("pairing_status"),
            "pairing_reason": source_item.get("metadata", {}).get("pairing_reason"),
            "associated_files": associated_files,
            "semantic_input_hash": source_item.get("semantic_input_hash"),
            "cache_hit": bool(source_item.get("cache_hit")),
            "cache_warning": source_item.get("cache_warning"),
            "diagnostic_code": diagnostic or result.get("diagnostic_code"),
            "result": result,
            "name_consistency": source_item["name_consistency"],
            "action": "CONTINUE_WITH_REVIEW",
        }
        if relative_path:
            inputs_by_relative_path[relative_path] = source_item
        projects.append(item)
        if item["status"] == "MISMATCH":
            mismatches.append(item)
        source_diagnostic = str(source_item.get("metadata", {}).get("diagnostic_code") or "")
        if result.get("failure") or diagnostic or source_diagnostic:
            failures.append({"name": name, "relative_path": relative_path, "diagnostic_code": diagnostic, "error": str(result.get("failure") or item["reason"])})
    if progress is not None and progress_current is not None:
        _publish_progress(progress, "标题正文资格预检完成", progress_current, "已将结果交给后续阶段", failures=len(failures), active=False)
    return {
        "status": "COMPLETED_WITH_WARNINGS" if failures else "COMPLETED",
        "mode": mode,
        "provider": provider,
        "model_id": runtime.model_id,
        "cache_enabled": cache is not None,
        "cache_disabled_reason": None if cache is not None else (
            "TYPESAFE_API_KEY_UNAVAILABLE" if provider == "typesafe_jev" and not os.environ.get("TYPESAFE_API_KEY", "").strip()
            else "ACTUAL_MODEL_IDENTITY_UNAVAILABLE"
        ),
        "cache_identity_elapsed_ms": cache_identity_elapsed_ms,
        "runtime_status": {
            "state": runtime_status.state,
            "diagnostic_code": runtime_status.diagnostic_code,
            "reason": runtime_status.reason,
            "server_version": runtime_status.server_version,
        },
        "prompt_version": str(config.get("prompt_version", "title-body-jev-domain-fit-score-v3")),
        "policy_version": str(config.get("policy_version", "title-body-jev-threshold-policy-v1")),
        "config_path": str(config_path),
        "config_error": config.get("config_error"),
        "scanned_count": len(items),
        "cache_hit_count": sum(bool(item.get("cache_hit")) for item in items),
        "cache_failure_count": sum(bool(item.get("cache_warning")) for item in items),
        "model_request_count": sum(
            int(item.get("result", {}).get("diagnostics", {}).get("generation_requests", 0) or 0)
            for item in items if not item.get("cache_hit")
        ),
        "failure_count": len(failures),
        "failures": failures,
        "name_warning_count": sum(item["name_consistency"]["warning_count"] for item in projects),
        "name_warning_names": [item["name"] for item in projects if item["name_consistency"]["warning_count"]],
        "projects": projects,
        "disqualified_projects": mismatches,
        "_inputs_by_relative_path": inputs_by_relative_path,
    }


def _code_inventory(root: Path, *, inspect: bool, excluded: list[dict[str, str]] | None = None) -> list[dict[str, Any]]:
    excluded = excluded or []
    all_docx_files, all_pdf_files = code_comment_cleaner._code_files(root)
    excluded_docx = [path for path in all_docx_files if _is_excluded_project(path, root, excluded)]
    excluded_pdf = [path for path in all_pdf_files if _is_excluded_project(path, root, excluded)]
    docx_files = [path for path in all_docx_files if path not in excluded_docx]
    pdf_files = [path for path in all_pdf_files if path not in excluded_pdf]
    pdf_by_key = {code_comment_cleaner._pair_key(path): path for path in pdf_files}
    docx_keys = {code_comment_cleaner._pair_key(path) for path in docx_files}
    excluded_pdf_by_key = {code_comment_cleaner._pair_key(path): path for path in excluded_pdf}
    rows: list[dict[str, Any]] = [
        {
            "docx": str(docx),
            "pdf": str(excluded_pdf_by_key.get(code_comment_cleaner._pair_key(docx), "")),
            "status": "SKIPPED_TOPIC_MISMATCH",
            "deletions": 0,
            "error": "项目标题与正文主题不匹配，按流程设置跳过",
            "findings": [],
        }
        for docx in excluded_docx
    ]
    for docx in docx_files:
        pdf = pdf_by_key.get(code_comment_cleaner._pair_key(docx))
        if pdf is None:
            rows.append({
                "docx": str(docx), "pdf": "", "status": "SKIPPED_MISSING_PDF",
                "deletions": 0, "error": "缺少同名代码 PDF，未处理", "findings": [],
            })
            continue
        row = code_comment_cleaner.inspect_docx(docx) if inspect else {
            "docx": str(docx), "status": "READY", "deletions": 0, "findings": [],
        }
        row["pdf"] = str(pdf)
        rows.append(row)
    for pdf in excluded_pdf:
        if code_comment_cleaner._pair_key(pdf) not in {code_comment_cleaner._pair_key(path) for path in excluded_docx}:
            rows.append({
                "docx": "", "pdf": str(pdf), "status": "SKIPPED_TOPIC_MISMATCH",
                "deletions": 0, "error": "项目标题与正文主题不匹配，按流程设置跳过", "findings": [],
            })
    for pdf in pdf_files:
        if code_comment_cleaner._pair_key(pdf) not in docx_keys:
            rows.append({
                "docx": "", "pdf": str(pdf), "status": "SKIPPED_MISSING_DOCX",
                "deletions": 0, "error": "缺少同名代码 DOCX，未处理", "findings": [],
            })
    if not all_docx_files and not all_pdf_files:
        rows.append({
            "docx": "", "pdf": "", "status": "MISSING_CODE_DOCX",
            "deletions": 0, "error": "未找到代码 DOCX，未处理", "findings": [],
        })
    return rows


def _run_code_stage(
    root: Path,
    app_root: Path,
    work: Path,
    *,
    apply: bool,
    progress: Any,
    offset: int,
    failures: int,
    excluded: list[dict[str, str]] | None = None,
    inspect: bool = True,
    cleanup: bool = True,
    export_pdf: bool = True,
) -> tuple[dict[str, Any], int]:
    rows = _code_inventory(root, inspect=inspect or cleanup, excluded=excluded)
    results: list[dict[str, Any]] = []
    stage_failures = 0
    for index, discovered in enumerate(rows, 1):
        item_name = Path(discovered.get("docx") or discovered.get("pdf") or "代码材料").stem
        current = offset + index
        transaction_snapshot: dict[Path, bytes | None] = {}
        try:
            if discovered["status"] != "READY":
                row = discovered
            else:
                docx_path = Path(discovered["docx"])
                row = code_comment_cleaner.inspect_docx(docx_path) if (inspect or cleanup) else {
                    "docx": str(docx_path), "status": "READY", "deletions": 0, "findings": [],
                }
                if apply and cleanup:
                    original_pdf = Path(discovered["pdf"]) if discovered.get("pdf") else None
                    transaction_paths = [docx_path, *([original_pdf] if original_pdf else [])]
                    transaction_snapshot = {
                        path: path.read_bytes() if path.is_file() else None
                        for path in transaction_paths
                    }
                    target_pages = code_comment_cleaner.pdf_page_count(original_pdf)
                    row = code_comment_cleaner.translate_and_fit_docx(docx_path, app_root, work, target_pages)
                row["pdf"] = discovered["pdf"]
                if apply and export_pdf:
                    _publish_progress(progress, "导出代码PDF", current - 1, item_name, failures=failures + stage_failures)
                    with process_batch.word_automation_lock(progress, "代码处理"):
                        exported, detail = code_comment_cleaner._export_pdf(docx_path, app_root, work, row.get("target_pages"))
                    row["pdf_status"] = "REPLACED" if exported else "PRESERVED"
                    row["pdf_detail"] = detail
                    if not exported:
                        row["status"] = "CLEANED_PDF_FAILED"
                elif export_pdf:
                    row["pdf_status"] = "READY_TO_EXPORT"
                elif apply and cleanup and row.get("deletions"):
                    row["pdf_status"] = "STALE_NOT_EXPORTED"
                    row["pdf_detail"] = "代码 DOCX 已清理，但未选择导出代码 PDF；配套 PDF 可能过期"
                else:
                    row["pdf_status"] = "NOT_SELECTED"
        except Exception as exc:
            row = {
                "docx": discovered.get("docx", ""), "pdf": discovered.get("pdf", ""),
                "status": "FAILED", "deletions": 0, "error": str(exc), "findings": [],
            }
        if row.get("page_60_lines") is not None and int(row["page_60_lines"]) < 51:
            row["page_match"] = False
        if row.get("page_match") is False:
            target_pages = row.get("target_pages")
            actual_pages = row.get("actual_pages")
            row["page_tolerance_match"] = False
            row["acceptance_error"] = (
                f"页数不足：至少 {target_pages} 页，实际 {actual_pages} 页；"
                f"第 60 页需至少 51 行，实际 {row.get('page_60_lines', '未读取')} 行"
            )
            if row.get("status") == "CLEANED":
                row["status"] = "CLEANED_PAGE_MISMATCH"
        item_failed = (
            row.get("status") not in {"READY", "CLEANED", "SKIPPED_TOPIC_MISMATCH"}
            or row.get("page_tolerance_match") is False
        )
        if transaction_snapshot and item_failed:
            try:
                for path, original_bytes in transaction_snapshot.items():
                    if original_bytes is None:
                        path.unlink(missing_ok=True)
                    else:
                        path.write_bytes(original_bytes)
                row["transaction_rollback"] = "DOCX/PDF 已恢复至处理前字节"
            except OSError as rollback_error:
                row["transaction_rollback"] = "恢复快照失败"
                row["error"] = "；".join(
                    value for value in (str(row.get("error")), f"事务快照恢复失败：{rollback_error}") if value
                )
                item_failed = True
        if item_failed:
            stage_failures += 1
        results.append(row)
        event_error = "；".join(
            str(value) for value in (
                row.get("error"), row.get("acceptance_error"),
                row.get("pdf_detail") if item_failed else None,
            ) if value
        ) or None
        _publish_progress(
            progress,
            "代码处理完成" if apply else "代码预检完成",
            current,
            item_name,
            status="FAILED" if item_failed else "RUNNING",
            failures=failures + stage_failures,
            error=event_error,
            active=False,
        )
    return {
        "stage": "code",
        "status": "PARTIAL" if stage_failures else ("APPLIED" if apply else "READY_FOR_REVIEW"),
        "options": {"inspect": inspect, "cleanup": cleanup, "export_pdf": export_pdf},
        "documents": results,
        "deletions": sum(int(row.get("deletions", 0)) for row in results),
        "failures": stage_failures,
    }, stage_failures


def _run_txt_stage(
    root: Path,
    *,
    apply: bool,
    progress: Any,
    offset: int,
    failures: int,
    check: bool = True,
    repair: bool = True,
    excluded: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    files = sorted((path for path in root.rglob("*") if path.is_file() and path.suffix.casefold() == ".txt"), key=lambda path: str(path).casefold())
    items: list[dict[str, Any]] = []
    failures_count = 0
    totals: dict[str, int] = {}
    field_re = re.compile(r"^(\s*(?:主要功能|软件主要功能)\s*[:：]\s*)(.*)$")
    for path in files:
        relative = path.relative_to(root).as_posix()
        row: dict[str, Any] = {
            "name": path.stem,
            "folder": path.parent.relative_to(root).as_posix(),
            "file": path.name,
            "path": relative,
            "main_function_changes": [],
            "repairs": {},
        }
        if _is_excluded_project(path, root, excluded or []):
            row["status"] = "SKIPPED_TOPIC_MISMATCH"
            row["findings"] = []
            items.append(row)
            continue
        try:
            inspection = process_batch.inspect_txt(path, path.stem)
            row["encoding"] = inspection.get("encoding")
            row["line_count"] = inspection.get("line_count", 0)
            row["findings"] = [
                finding for finding in inspection.get("clear_errors", [])
                if finding.get("message") != "主要功能包含序号或序数词"
            ]
            row["candidates"] = inspection.get("candidates", [])
            row["original_valid"] = not bool(row["findings"])
            if row["encoding"] == "unknown":
                raise ValueError("无法可靠识别原编码，已停止该文件的自动修复")
            original = path.read_bytes()
            updated_lines: list[str] = []
            changed = False
            for line in (inspection.get("text", "").splitlines(keepends=True) if repair else []):
                content = line.rstrip("\r\n")
                match = field_re.match(content)
                if not match:
                    updated_lines.append(line)
                    continue
                clean_value, counts = process_batch.repair_main_function_text(
                    match.group(2), remove_ordinal_markers=False
                )
                for key, value in counts.items():
                    row["repairs"][key] = row["repairs"].get(key, 0) + value
                    totals[key] = totals.get(key, 0) + value
                if clean_value != match.group(2):
                    row["main_function_changes"].append({"before": match.group(2), "after": clean_value})
                    updated_lines.append(match.group(1) + clean_value + line[len(content):])
                    changed = True
                else:
                    updated_lines.append(line)
            row["changed"] = changed
            row["original_valid"] = row["original_valid"] and not bool(row["repairs"])
            if changed and apply and repair:
                process_batch.write_txt_with_original_encoding(path, original, "".join(updated_lines), str(row["encoding"] or "unknown"))
                final_inspection = process_batch.inspect_txt(path, path.stem)
                row["final_findings"] = [
                    finding for finding in final_inspection.get("clear_errors", [])
                    if finding.get("message") != "主要功能包含序号或序数词"
                ]
                row["final_valid"] = not bool(row["final_findings"])
                row["status"] = "REPAIRED"
            elif changed and repair:
                row["status"] = "CHANGES_PREVIEWED"
            else:
                row["final_findings"] = row["findings"]
                row["final_valid"] = row["original_valid"]
                row["status"] = "CHECKED" if check and not repair else "NO_CHANGE"
        except Exception as exc:
            failures_count += 1
            row["status"] = "FAILED"
            row["error"] = str(exc)
        items.append(row)
    work_items = files or [None]
    for index, path in enumerate(work_items, 1):
        skipped = bool(path and _is_excluded_project(path, root, excluded or []))
        _publish_progress(
            progress,
            "跳过标题主题不匹配项目" if skipped else "TXT处理",
            offset + index,
            path.name if path else "未找到TXT",
            failures=failures + failures_count,
            status="FAILED" if path and next((item.get("status") == "FAILED" for item in items if item.get("path") == path.relative_to(root).as_posix()), False) else "RUNNING",
            active=False,
        )
    return {
        "stage": "txt",
        "status": "PARTIAL" if failures_count else ("APPLIED" if apply else "READY_FOR_REVIEW"),
        "message": "TXT 先检查，再按选项修复主要功能字段；原编码和其他字段保留。" if apply else "TXT 检查与修复预览，不修改原文件。",
        "options": {"check": check or repair, "repair": repair},
        "files": items,
        "file_count": len(items),
        "changed_count": sum(bool(item.get("changed")) and item.get("status") != "SKIPPED_TOPIC_MISMATCH" for item in items),
        "repair_totals": totals,
        "failures": failures_count,
    }


def _manual_failure_count(report: dict[str, Any], apply: bool) -> int:
    if not apply:
        return 0
    failed_statuses = {"BATCH_FUSED", "BLOCKED_IRRELEVANT", "DATE_REWRITER_UNAVAILABLE"}
    return sum(project.get("status") in failed_statuses for project in report.get("projects", []))


def _build_manual_cover_plan(
    root: Path,
    path: Path,
    excluded: list[dict[str, str]] | None = None,
    selected_options: tuple[str, ...] = (),
) -> dict[str, Any]:
    packages: list[dict[str, Any]] = []
    audit_errors: list[dict[str, str]] = []
    unchanged: list[dict[str, str]] = []
    selected = set(selected_options)
    for docx in sorted(root.rglob("*说明.docx"), key=lambda item: str(item).casefold()):
        if _is_excluded_project(docx, root, excluded or []):
            continue
        txt_files = sorted(
            (item for item in docx.parent.glob("*.txt") if not item.name.startswith("~$")),
            key=lambda item: item.name.casefold(),
        )
        if len(txt_files) == 1:
            title = txt_files[0].stem
            name_source = "unique_txt_filename"
        else:
            title = docx.parent.name
            name_source = "project_folder"
        docx_title = docx.stem.removesuffix("说明")
        audit: dict[str, Any] = {}
        audit_error: str | None = None
        try:
            audit = process_batch.inspect_docx(docx, title, False)
        except Exception as exc:
            audit_error = str(exc)
            audit_errors.append({"docx": str(docx), "error": audit_error})

        actions: list[str] = []
        if "manual_cover" in selected:
            # Word pagination details are deferred to a targeted repair only
            # when the exported PDF screen finds a page anomaly. The PDF screen
            # scans all pages without asking Word COM to map every paragraph.
            actions.extend(["ENSURE_COVER_FIRST_PAGE", "AUDIT_PDF_ALL_PAGES", "NORMALIZE_COVER_STYLE", "EXPORT_PDF"])
            if audit.get("duplicate_cover_blocks") or audit.get("residual_manual_labels"):
                actions.append("REMOVE_DUPLICATE_COVER_TEXT")
            if audit.get("cover_title_has_version"):
                actions.append("REMOVE_COVER_VERSION")
        if "manual_header" in selected or "manual_page_numbers" in selected:
            try:
                header = process_batch.inspect_docx_header_only(docx, title)
                if "manual_header" in selected and not (header.get("header_title") and header.get("header_version")):
                    actions.append("NORMALIZE_HEADER")
                active_counts = list(header.get("header_page_field_counts", []))
                footer_counts = list(header.get("footer_page_field_counts", []))
                active_count = max(len(active_counts), len(footer_counts))
                active_counts.extend([0] * max(0, active_count - len(active_counts)))
                footer_counts.extend([0] * max(0, active_count - len(footer_counts)))
                page_compliant = bool(active_count) and all(
                    header_count + footer_count == 1
                    for header_count, footer_count in zip(active_counts, footer_counts)
                )
                if "manual_page_numbers" in selected and not page_compliant:
                    actions.append("NORMALIZE_PAGE_NUMBERS")
            except Exception as exc:
                audit_errors.append({"docx": str(docx), "error": f"页眉页码预检失败：{exc}"})
        if "manual_pdf" in selected:
            # The post-export PDF audit is Python-only. Keep this marker out of
            # the Word action set so a PDF-only request does not trigger cover
            # layout edits in Word.
            actions.append("AUDIT_PAGE_LABELS")
            if "EXPORT_PDF" not in actions:
                actions.append("EXPORT_PDF")
        pdf = docx.with_suffix(".pdf")
        package = {
            "name": title,
            "name_source": name_source,
            "docx_title": docx_title,
            "name_conflicts": ["DOCX 文件名与软件全称不一致"] if docx_title != title else [],
            "cover_text_conflicts": [],
            "static_audit_error": audit_error,
            "static_audit": audit,
            "folder": str(docx.parent.relative_to(root)),
            "docx": docx.name,
            "pdf": pdf.name,
            "actions": list(dict.fromkeys(actions)),
        }
        if package["actions"]:
            packages.append(package)
        else:
            unchanged.append({"name": title, "folder": package["folder"], "docx": docx.name, "reason": "所选整理项检查通过，无需修改"})
    plan = {
        "version": VERSION,
        "base": str(root),
        "created_at": datetime.now().astimezone().isoformat(),
        "packages": packages,
        "unchanged": unchanged,
        "audit_errors": audit_errors,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    return plan


def _cover_result_failures(report: dict[str, Any]) -> int:
    return sum(
        row.get("status") not in {"OK"} or bool(row.get("failures"))
        for row in report.get("packages", [])
    )


def _workflow_dependency_warnings(options: tuple[str, ...]) -> list[dict[str, str]]:
    selected = set(options)
    warnings: list[dict[str, str]] = []
    if selected.intersection({"manual_body_years", "manual_image_dates", "manual_cover", "manual_header", "manual_page_numbers"}) and "manual_pdf" not in selected:
        warnings.append({"option": "manual_pdf", "reason": "说明书 DOCX 有修改风险，但未选择重新导出 PDF；DOCX/PDF 可能不一致"})
    if "code_cleanup" in selected and "code_pdf" not in selected:
        warnings.append({"option": "code_pdf", "reason": "代码 DOCX 有修改风险，但未选择重新导出代码 PDF；DOCX/PDF 可能不一致"})
    return warnings


def _year_image_comparisons(path: Path, report: dict[str, Any]) -> str:
    """Build offline image evidence from the original and processed DOCX packages."""
    if not report.get("applied"):
        return '<p>预览阶段尚未修改图片；执行后显示原稿与输出的图片对比。</p>'
    source_root = Path(report["material_root"]).resolve()
    output_root = Path(report["output"]).resolve()
    assets = path.with_name(path.stem + "_assets")
    cards: list[str] = []
    try:
        from PIL import Image, ImageChops, ImageOps
    except ImportError:
        return '<p>缺少 Pillow，无法核对原稿与输出图片的像素差异。</p>'
    supported = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
    for project in report["stages"]["manual"].get("projects", []):
        result = project.get("date_rewrite") or {}
        if not result.get("changed"):
            continue
        recorded_media = {
            item.get("media_path") for item in result.get("changes", [])
            if item.get("changed_image") and item.get("media_path")
        }
        relative = Path(str(project.get("relative_path", "")))
        original = (source_root / relative).resolve()
        processed = (output_root / relative).resolve()
        if not original.is_relative_to(source_root) or not processed.is_relative_to(output_root):
            continue
        if not original.is_file() or not processed.is_file():
            continue
        try:
            with zipfile.ZipFile(original) as before_doc, zipfile.ZipFile(processed) as after_doc:
                before_names = {
                    name for name in before_doc.namelist()
                    if name.lower().startswith("word/media/") and Path(name).suffix.lower() in supported
                }
                after_names = set(after_doc.namelist())
                changed = []
                for media_name in sorted(before_names & after_names & recorded_media):
                    before_data = before_doc.read(media_name)
                    after_data = after_doc.read(media_name)
                    if before_data == after_data:
                        continue
                    try:
                        before_image = ImageOps.exif_transpose(Image.open(io.BytesIO(before_data))).convert("RGB")
                        after_image = ImageOps.exif_transpose(Image.open(io.BytesIO(after_data))).convert("RGB")
                        # Word's later PDF/cover pass can resize other pictures.
                        # The year rewriter preserves dimensions, so those are not
                        # evidence of its date edits.
                        if before_image.size != after_image.size:
                            continue
                        if ImageChops.difference(before_image, after_image).getbbox() is None:
                            continue
                    except (OSError, ValueError):
                        continue
                    changed.append((media_name, before_data, after_data, before_image, after_image))
                for index, (media_name, before_data, after_data, before_image, after_image) in enumerate(changed, 1):
                    assets.mkdir(parents=True, exist_ok=True)
                    prefix = f"{len(cards) + 1:03d}"
                    files = []
                    for label, picture in (("before", before_image), ("after", after_image)):
                        suffix = ".png"
                        filename = f"{prefix}-{label}{suffix}"
                        target = assets / filename
                        picture.save(target, format="PNG")
                        files.append(f"{assets.name}/{filename}")
                    zoom = ""
                    if before_image is not None and after_image is not None and before_image.size == after_image.size:
                        # JPEG recompression causes low-level noise across the image.
                        # A high threshold isolates the visually changed year glyphs.
                        difference = ImageChops.difference(before_image, after_image).convert("L")
                        strong = difference.point(lambda value: 255 if value >= 45 else 0)
                        bounds = strong.getbbox()
                        if bounds:
                            width, height = before_image.size
                            left, top, right, bottom = bounds
                            if (right - left) * (bottom - top) < width * height * 0.5:
                                margin = max(20, min(width, height) // 30)
                                crop_box = (max(0, left - margin), max(0, top - margin),
                                            min(width, right + margin), min(height, bottom + margin))
                                crop_files = []
                                for label, picture in (("before", before_image), ("after", after_image)):
                                    filename = f"{prefix}-{label}-detail.png"
                                    picture.crop(crop_box).save(assets / filename, format="PNG")
                                    crop_files.append(f"{assets.name}/{filename}")
                                zoom = ("<div class='pair'><figure><img src='{}' alt='修改前局部'><figcaption>修改前局部</figcaption></figure>"
                                        "<figure><img src='{}' alt='修改后局部'><figcaption>修改后局部</figcaption></figure></div>").format(
                                            *(html.escape(value, quote=True) for value in crop_files))
                    title = html.escape(str(relative))
                    name = html.escape(media_name)
                    recorded_fields = [
                        f"{item.get('source_year')} → {item.get('target_year')}（{item.get('ocr_text')}）"
                        for item in result.get("changes", [])
                        if item.get("media_path") == media_name and item.get("changed_image")
                    ]
                    field_detail = (
                        "<p>识别到的旧日期年份：" + html.escape("；".join(recorded_fields)) + "</p>"
                        if recorded_fields else ""
                    )
                    cards.append(
                        f"<article class='comparison'><h3>{title} · {name}</h3>"
                        + field_detail +
                        "<p class='review'>待核验：这张图片的像素发生变化，但仅凭像素差异无法确认修改的是年份。"
                        "请对照原图检查日期语义和处理后字符，尤其留意时间、编号被误识别为年份的情况。</p>"
                        + zoom +
                        "<div class='pair'>" + "".join(
                            f"<figure><a href='{html.escape(src, quote=True)}' target='_blank'>"
                            f"<img src='{html.escape(src, quote=True)}' alt='{caption}'></a>"
                            f"<figcaption>{caption}（点击查看原尺寸）</figcaption></figure>"
                            for src, caption in zip(files, ("原稿", "处理后"))
                        ) + "</div></article>"
                    )
        except (OSError, zipfile.BadZipFile):
            continue
    if not cards:
        return '<p>没有可展示的图片年份修改对比。</p>'
    return (f"<p>以下 {len(cards)} 张图片从原稿与处理后 DOCX 提取，确认存在像素变化。"
            "自动检测结果尚需逐张核验，不能仅据此认定年份改写正确。</p>" + "".join(cards))


def _company_ocr_alerts(manual_report: dict[str, Any], effective_stages: list[str]) -> dict[str, Any]:
    """Find a literal company marker in OCR text already produced for sampled images."""
    if not {"match_manual_ocr", "match_manual_vision"}.intersection(effective_stages):
        return {"status": "NOT_SELECTED", "scanned_image_count": 0, "warning_count": 0, "matches": []}
    matches: list[dict[str, Any]] = []
    scanned = 0
    for project in manual_report.get("projects", []):
        if not isinstance(project, dict):
            continue
        for sample in project.get("samples", []):
            if not isinstance(sample, dict):
                continue
            scanned += 1
            text = str(sample.get("ocr_text") or "")
            count = text.count("公司")
            if not count:
                continue
            lines = [line.strip() for line in text.splitlines() if "公司" in line]
            matches.append({
                "project": str(project.get("title") or Path(str(project.get("relative_path", ""))).stem.removesuffix("说明")),
                "relative_path": str(project.get("relative_path") or ""),
                "image": str(sample.get("image") or ""),
                "sample_number": int(sample.get("index", -1)) + 1,
                "occurrence_count": count,
                "matched_lines": lines[:8],
            })
    return {
        "status": "FOUND" if matches else "CLEAR",
        "scanned_image_count": scanned,
        "warning_count": sum(item["occurrence_count"] for item in matches),
        "matching_image_count": len(matches),
        "matches": matches,
    }


def _company_ocr_alert_html(alert: dict[str, Any]) -> str:
    if alert.get("status") == "NOT_SELECTED":
        return "<section><h2>图片 OCR 公司字样检查</h2><p>本次未选择截图 OCR 检查，未执行此项判断。</p></section>"
    if alert.get("status") != "FOUND":
        return (
            '<section><h2>图片 OCR 公司字样检查</h2><p>已检查 {} 张抽样图片，OCR 文字中未发现“公司”。'
            '此项只检查当前 OCR 抽样结果。</p></section>'
        ).format(int(alert.get("scanned_image_count", 0)))
    rows = []
    for item in alert.get("matches", []):
        evidence = "<br>".join(html.escape(line) for line in item.get("matched_lines", [])) or "OCR 识别到“公司”"
        rows.append("<tr><td>{}</td><td>第 {} 张（{}）</td><td>{}</td><td>{}</td></tr>".format(
            html.escape(item.get("project", "")), int(item.get("sample_number", 0)),
            html.escape(item.get("image", "")), int(item.get("occurrence_count", 0)), evidence,
        ))
    return (
        '<section class="company-word-alert"><h2>⚠ OCR 图片红色警告：发现“公司”字样</h2>'
        '<p>在 {} 张抽样图片的 OCR 结果中命中 {} 次，涉及 {} 张图片。请核对截图是否包含公司名称或标识。</p>'
        '<table><thead><tr><th>项目</th><th>图片</th><th>命中次数</th><th>OCR 原文</th></tr></thead><tbody>{}</tbody></table>'
        '<p>这是人工复核提示，不会自动跳过或修改材料。</p></section>'
    ).format(int(alert.get("scanned_image_count", 0)), int(alert.get("warning_count", 0)),
             int(alert.get("matching_image_count", 0)), "".join(rows))


def _write_report(path: Path, report: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)

    stage_rows = []
    for key in ("manual", "code", "txt"):
        stage = report["stages"][key]
        count_text = int(stage.get("failures", 0))
        if key == "txt" and not report.get("applied"):
            count_text = f"{int(stage.get('changed_count', 0))} 项待修改；{count_text} 项失败"
        stage_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape({"manual": "说明书", "code": "代码", "txt": "TXT"}[key]),
                html.escape(str(stage.get("status", ""))),
                html.escape(str(count_text)),
            )
        )
    code_rows = []
    for item in report["stages"]["code"].get("documents", []):
        name = item.get("docx") or item.get("pdf") or "代码材料"
        code_rows.append(
            f"<li><b>{html.escape(str(name))}</b>：{html.escape(str(item.get('status', '')))}；"
            f"预计/实际删除 {int(item.get('deletions', 0))} 处"
            f"；代码 PDF：{html.escape(str(item.get('pdf_status', '未选择导出')))}"
            f"{('；' + html.escape(str(item.get('pdf_detail'))) ) if item.get('pdf_detail') else ''}"
            f"{('；' + html.escape(str(item.get('acceptance_error'))) ) if item.get('acceptance_error') else ''}"
            f"{('；' + html.escape(str(item.get('error')))) if item.get('error') else ''}</li>"
        )
    manual_rows = []
    visual_rows = []
    disqualified_rows = []
    similarity_rows = []
    for item in report.get("qualification", {}).get("projects", []):
        topic_status = str(item.get("semantic_status", item.get("status", "NOT_EVALUATED"))).upper()
        result = item.get("result", {}) if isinstance(item.get("result"), dict) else {}
        score_value = result.get("score_0_100")
        threshold = result.get("decision_threshold")
        score_label = (
            f"{float(score_value):.1f}/100；阈值 {float(threshold):.1f}；"
            f"置信度 {float(result.get('confidence', 0)):.2f}"
            if isinstance(score_value, (int, float)) and isinstance(threshold, (int, float))
            else "旧规则估算" if item.get("legacy_similarity_score", item.get("similarity_score")) is not None else "—"
        )
        decision = {"MATCH": "匹配", "MISMATCH": "不匹配", "NOT_EVALUATED": "未评估", "NOT_RUN": "未运行"}.get(topic_status, topic_status)
        action = str(item.get("action", "CONTINUE_WITH_REVIEW"))
        skipped = action in {"SKIP_TITLE_BODY_MISMATCH", "HOLD_NOT_EVALUATED"}
        evidence = result.get("evidence", []) if isinstance(result.get("evidence"), list) else []
        evidence_text = "；".join(str(row.get("quote", "")) for row in evidence if isinstance(row, dict) and row.get("quote"))
        probability_text = result.get("score_probabilities")
        usage = result.get("usage", {}) if isinstance(result.get("usage"), dict) else {}
        result_diagnostics = result.get("diagnostics", {}) if isinstance(result.get("diagnostics"), dict) else {}
        metrics = []
        if isinstance(probability_text, dict):
            metrics.append("评分等级概率 " + json.dumps(probability_text, ensure_ascii=False, sort_keys=True))
        if usage.get("input_tokens") is not None or usage.get("output_tokens") is not None:
            metrics.append(f"token 输入/输出 {usage.get('input_tokens', '—')}/{usage.get('output_tokens', '—')}")
        if result_diagnostics.get("elapsed_ms") is not None:
            metrics.append(f"耗时 {int(result_diagnostics['elapsed_ms'])} ms")
        reason_text = str(item.get("reason", ""))
        detail_text = "；".join(filter(None, [reason_text, "；".join(metrics), evidence_text]))
        similarity_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape(str(item.get("name", ""))),
                html.escape(score_label),
                html.escape(decision),
                html.escape(action if skipped else "继续处理，建议复核"),
                html.escape(detail_text),
            )
        )
    selected_names = "、".join(WORKFLOW_OPTION_LABELS.get(item, item) for item in report.get("selected_stages", [])) or "无"
    dependency_rows = "、".join(
        f"{WORKFLOW_OPTION_LABELS.get(item.get('option', ''), item.get('option', ''))}（{item.get('reason', '')}）"
        for item in report.get("auto_dependencies", [])
    ) or "无"
    warning_rows = "、".join(item.get("reason", "") for item in report.get("dependency_warnings", [])) or "无"
    execution_order = " → ".join(report.get("stage_order", []))
    for item in report.get("qualification", {}).get("disqualified_projects", []):
        result = item.get("result", {}) if isinstance(item.get("result"), dict) else {}
        evidence = result.get("evidence", []) if isinstance(result.get("evidence"), list) else []
        evidence_text = "；".join(str(row.get("quote", "")) for row in evidence if isinstance(row, dict) and row.get("quote"))
        disqualified_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape(str(item.get("name", ""))),
                html.escape(str(item.get("action", "CONTINUE_WITH_REVIEW"))),
                html.escape("；".join(filter(None, [str(item.get("reason", "")), evidence_text]))),
            )
        )
    for project in report["stages"]["manual"].get("projects", []):
        samples = project.get("samples", [])
        counts = {
            state: sum(item.get("status") == state for item in samples)
            for state in ("relevant", "irrelevant", "unknown")
        }
        year_rows = project.get("year_changes", [])
        updated_years = sum(bool(item.get("changed")) for item in year_rows)
        review_years = sum(item.get("decision") == "REVIEW" for item in year_rows)
        date_result = project.get("date_rewrite", {})
        body_relevance = project.get("body_relevance", {})
        title_body = project.get("title_body_result", {}) if isinstance(project.get("title_body_result"), dict) else {}
        body_status = str(title_body.get("status", project.get("qualification_status", body_relevance.get("status", "尚未审核"))))
        body_score = "旧规则估算" if body_relevance.get("similarity_score") is not None else "—"
        body_detail = str(title_body.get("reason", body_relevance.get("reason", "尚未审核")))
        visual_samples = [item for item in samples if item.get("visual_review", {}).get("status") != "not_requested" and item.get("visual_review")]
        visual_counts = {
            state: sum(item["visual_review"].get("status") == state for item in visual_samples)
            for state in ("relevant", "irrelevant", "unknown")
        }
        visual_summary = (
            f"抽看 {int(project.get('visual_call_count', 0))}/{int(project.get('visual_candidate_count', len(visual_samples)))}；"
            f"相关 {visual_counts['relevant']} / 不相关 {visual_counts['irrelevant']} / 无法判断 {visual_counts['unknown']}；"
            f"接口失败 {int(project.get('visual_failure_count', 0))}"
            if visual_samples else ("资格不合格，跳过识图" if project.get("status") == "SKIPPED_TOPIC_MISMATCH" else "未请求")
        )
        manual_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}（{}）；{}</td><td>相关 {} / 不相关 {} / 无法判断 {}</td><td>{}</td>"
            "<td>更新 {} / 人工复核 {}</td><td>{}</td></tr>".format(
                html.escape(str(project.get("relative_path", ""))),
                html.escape(str(project.get("status", ""))),
                body_status, body_score, html.escape(body_detail),
                counts["relevant"], counts["irrelevant"], counts["unknown"],
                visual_summary,
                updated_years, review_years,
                "检测到变化（待核验）" if date_result.get("changed") else html.escape(str(date_result.get("status", "预览阶段未修改"))),
            )
        )
        if visual_samples:
            items = "".join(
                "<li>{}：{}；{}</li>".format(
                    html.escape(str(item.get("image", ""))),
                    html.escape(str(item["visual_review"].get("status", "unknown"))),
                    html.escape("；".join(filter(None, [
                        str(item["visual_review"].get("reason", "")),
                        (f"图片 {int(item['visual_review'].get('image_original_bytes', 0))}→{int(item['visual_review'].get('image_sent_bytes', 0))} 字节（上限 {int(item['visual_review'].get('image_max_bytes', 0))}）") if item["visual_review"].get("image_original_bytes") is not None else "",
                        ("接口错误：" + str(item["visual_review"].get("error"))) if item["visual_review"].get("error") else "",
                    ])) or "无可用证据"),
                )
                for item in visual_samples
            )
            visual_rows.append(
                f"<details><summary>{html.escape(str(project.get('relative_path', '')))}：{visual_summary}</summary><ul>{items}</ul></details>"
            )
    cover_rows = []
    cover_evidence = []
    year_comparisons = _year_image_comparisons(path, report)
    cover_result = report["stages"]["manual"].get("cover_result", {})
    for item in cover_result.get("packages", []):
        requested = "、".join(map(str, item.get("actions_requested", item.get("actions", [])))) or "无"
        applied = "、".join(map(str, item.get("actions_applied", []))) or ("执行阶段生成" if not report["applied"] else "无")
        audit = item.get("pdf_audit", {})
        name_source = {"unique_txt_filename": "唯一 TXT 文件名", "project_folder": "项目文件夹名"}.get(
            str(item.get("name_source", "")), str(item.get("name_source", "未记录"))
        )
        conflicts = "、".join(map(str, item.get("name_conflicts", []))) or "无"
        cover_conflicts = "、".join(map(str, item.get("cover_text_conflicts", []))) or "无"
        original_cover_text = "、".join(map(str, item.get("original_cover_text", []))) or "无"
        blank_pages = "、".join(map(str, item.get("confirmed_blank_pages", []))) or "无"
        unresolved_pages = "、".join(map(str, item.get("unresolved_blank_pages", []))) or "无"
        bookmark_errors = "；".join(
            f"第 {int(error.get('page', 0))} 页（{error.get('type', '书签错误')}）：{error.get('snippet', '')}"
            for error in audit.get("bookmark_errors", [])
        ) or "无"
        blank_causes = "；".join(
            f"第 {int(candidate.get('page', 0))} 页：{'、'.join(map(str, candidate.get('suspected_causes', []))) or '无明确成因'}"
            for candidate in item.get("blank_page_diagnostics", [])
        ) or "无疑似空白页"
        attempt_count = max((int(attempt.get("attempt", 0)) for attempt in item.get("attempts", [])), default=0)
        restored = "；已恢复项目快照" if item.get("snapshot_restored") else ""
        broken_bookmark_fields = "；".join(
            f"{field.get('type', '未知')} {field.get('target', '未知')} → {field.get('cached_result', '')}"
            for field in item.get("broken_bookmark_fields_unlinked", [])
        ) or "无"
        failures_text = "；".join(map(str, item.get("failures", []))) or "无"
        details = (
            f"状态：{html.escape(str(item.get('status', '')))}{restored}；名称来源：{html.escape(name_source)}；"
            f"DOCX 名称：{html.escape(str(item.get('docx_title', '未记录')))}；名称冲突：{html.escape(conflicts)}；"
            f"原封面文字：{html.escape(original_cover_text)}；封面文字冲突/待分类：{html.escape(cover_conflicts)}；"
            f"Word/PDF 页数：{int(item.get('word_pages', 0))}/{int(item.get('pdf_page_count', audit.get('pages', 0)))}；"
            f"确认空白页：{html.escape(blank_pages)}；未能复核页：{html.escape(unresolved_pages)}；"
            f"书签错误文本：{html.escape(bookmark_errors)}；"
            f"空白页诊断：{html.escape(blank_causes)}；尝试次数：{attempt_count}；"
            f"缺失书签引用已保留缓存结果：{html.escape(broken_bookmark_fields)}；"
            f"文件清单哈希：{len(item.get('snapshot_manifest', []))} 项。"
        )
        cover_rows.append(
            f"<tr><td>{html.escape(str(item.get('name', '')))}<br>{details}</td><td>{html.escape(requested)}</td>"
            f"<td>{html.escape(applied)}</td><td>{html.escape(failures_text)}</td></tr>"
        )
        for candidate in audit.get("blank_page_candidates", []):
            image_path = candidate.get("render", {}).get("image")
            if not image_path:
                continue
            try:
                relative_image = Path(image_path).resolve().relative_to(path.parent.resolve()).as_posix()
            except (OSError, ValueError):
                continue
            page_number = int(candidate.get("page", 0))
            state = "确认空白" if candidate.get("confirmed_blank") else "渲染复核未判为空白"
            cover_evidence.append(
                f"<figure><a href='{html.escape(relative_image, quote=True)}' target='_blank'>"
                f"<img src='{html.escape(relative_image, quote=True)}' alt='第 {page_number} 页 PDF 渲染'></a>"
                f"<figcaption>{html.escape(str(item.get('name', '')))}：第 {page_number} 页，{state}</figcaption></figure>"
            )
        if item.get("attempts"):
            attempt_list = "".join(
                f"<li>第 {int(attempt.get('attempt', index))} 次：{html.escape(str(attempt.get('phase', '')))}；"
                f"{html.escape(str(attempt.get('status', '')))}；{html.escape('；'.join(map(str, attempt.get('failures', []))) or '通过')}</li>"
                for index, attempt in enumerate(item.get("attempts", []), 1)
            )
            cover_evidence.append(f"<details><summary>{html.escape(str(item.get('name', '')))}：验收尝试记录</summary><ol>{attempt_list}</ol></details>")
    for item in cover_result.get("unchanged", report.get("stages", {}).get("manual", {}).get("cover_plan", {}).get("unchanged", [])):
        cover_rows.append(
            f"<tr><td>{html.escape(str(item.get('name', '')))}</td><td>检查已选整理项</td>"
            f"<td>无需修改</td><td>{html.escape(str(item.get('reason', '')))}</td></tr>"
        )
    manual_report = report["stages"]["manual"]
    txt_rows = []
    for item in report.get("stages", {}).get("txt", {}).get("files", []):
        findings = item.get("final_findings", item.get("findings", []))
        candidates = item.get("candidates", [])
        details = [str(entry.get("message", "")) for entry in findings + candidates if entry.get("message")]
        labels = {
            "parentheses_to_commas": "圆括号转中文逗号",
            "quotes_removed": "删除双引号",
            "duplicate_punctuation_removed": "清理连续标点",
        }
        repair_summary = [
            f"{labels.get(key, key)} {int(value)} 处"
            for key, value in item.get("repairs", {}).items() if int(value)
        ]
        details.extend(repair_summary)
        if item.get("error"):
            details.append(str(item["error"]))
        state = {
            "SKIPPED_TOPIC_MISMATCH": "标题主题不匹配，已跳过",
            "CHANGES_PREVIEWED": "预览有待修复内容",
            "REPAIRED": "已修复",
            "NO_CHANGE": "无需修改",
            "FAILED": "处理失败",
        }.get(str(item.get("status", "")), str(item.get("status", "")))
        changes = changes_html(item.get("main_function_changes", []), "主要功能字段修改对比") or "无修改"
        txt_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape(str(item.get("path", ""))),
                html.escape(str(item.get("encoding", "未读取"))),
                html.escape(state),
                html.escape("；".join(details) if details else "无"),
                changes,
            )
        )
    vision_details = "视觉复核未运行"
    if manual_report.get("visual_backend"):
        vision_details = "视觉接口：{}；模型：{}；地址：{}；最大图片体积：{} 字节".format(
            html.escape(str(manual_report.get("visual_backend", ""))),
            html.escape(str(manual_report.get("visual_model", "未配置"))),
            html.escape(str(manual_report.get("visual_endpoint", "未配置"))),
            int(manual_report.get("visual_max_image_bytes", 0) or 0),
        )
    qualification_metrics = report.get("qualification", {})
    if not isinstance(qualification_metrics, dict):
        qualification_metrics = {}
    semantic_projects = qualification_metrics.get("projects", [])
    name_warnings_html = render_name_warnings(semantic_projects if isinstance(semantic_projects, list) else [])
    company_ocr_alert_html = _company_ocr_alert_html(report.get("company_ocr_alerts", {}))
    not_evaluated_count = sum(
        str(item.get("semantic_status", item.get("status", "NOT_EVALUATED"))).upper() == "NOT_EVALUATED"
        for item in semantic_projects if isinstance(item, dict)
    ) if isinstance(semantic_projects, list) else 0
    qualification_summary = (
        f"Provider {qualification_metrics.get('provider') or '未选择'}；模型 {qualification_metrics.get('model_id') or '未启动'}；"
        f"MATCH 纳入 {int(qualification_metrics.get('qualified_project_count', 0) or 0)}/"
        f"{int(qualification_metrics.get('input_project_count', 0) or 0)} 件；"
        f"缓存命中 {int(qualification_metrics.get('cache_hit_count', 0) or 0)}；"
        f"生成请求 {int(qualification_metrics.get('model_request_count', 0) or 0)}；"
        f"故障 {int(qualification_metrics.get('failure_count', 0) or 0)}；"
        f"未评估 {not_evaluated_count}"
    )
    all_semantic_results = [
        item.get("result", {}) for item in semantic_projects
        if isinstance(item, dict) and isinstance(item.get("result"), dict)
    ] if isinstance(semantic_projects, list) else []
    input_tokens = sum(
        int(result.get("usage", {}).get("input_tokens", 0) or 0)
        for result in all_semantic_results if isinstance(result.get("usage"), dict)
    )
    output_tokens = sum(
        int(result.get("usage", {}).get("output_tokens", 0) or 0)
        for result in all_semantic_results if isinstance(result.get("usage"), dict)
    )
    total_elapsed_ms = sum(
        int(result.get("diagnostics", {}).get("elapsed_ms", 0) or 0)
        for result in all_semantic_results if isinstance(result.get("diagnostics"), dict)
    )
    qualification_summary += f"；token 输入/输出 {input_tokens}/{output_tokens}；逐件耗时合计 {total_elapsed_ms} ms"
    mismatch_clipboard = qualification_metrics.get("mismatch_clipboard", {})
    if not isinstance(mismatch_clipboard, dict):
        mismatch_clipboard = {}
    clipboard_status_label = {
        "copied": "已复制到系统剪贴板",
        "preview_not_copied": "预览阶段未写入剪贴板",
        "no_mismatch": "无不匹配项目，未操作剪贴板",
        "failed": "剪贴板复制失败，处理继续",
    }.get(str(mismatch_clipboard.get("status", "")), "未尝试")
    mismatch_names_html = "<br>".join(
        html.escape(str(name)) for name in qualification_metrics.get("mismatch_names", [])
    ) or "无"
    mismatch_clipboard_error = str(mismatch_clipboard.get("error") or "")
    mismatch_clipboard_html = (
        f"<p><b>不匹配清单：</b>{mismatch_names_html}<br>"
        f"<b>剪贴板：</b>{html.escape(clipboard_status_label)}"
        f"{('<br><b>错误：</b>' + html.escape(mismatch_clipboard_error)) if mismatch_clipboard_error else ''}</p>"
    )
    if mismatch_clipboard.get("text"):
        mismatch_clipboard_html += (
            "<pre class=note>" + html.escape(str(mismatch_clipboard["text"])) + "</pre>"
        )
    hollow_line_alerts = report.get("ai_hollow_line_alerts", [])
    hollow_line_check_selected = "manual_ai_hollow_lines" in report.get("effective_stages", [])
    hollow_line_alert_html = (
        '<section class="name-alert"><h2>空心线检测报告</h2><p>发现参考图所示的横向空心线，请将以下说明书发给其他人：</p><ul>'
        + "".join(f"<li><b>{html.escape(str(item.get('name', '')))}</b>：{html.escape(str(item.get('message', '')))}</li>" for item in hollow_line_alerts)
        + "</ul></section>"
        if hollow_line_alerts else (
            "<section><h2>空心线检测报告</h2><p>未发现参考图所示的横向空心线。</p></section>"
            if hollow_line_check_selected else "<section><h2>空心线检测报告</h2><p>本次未选择此项检测。</p></section>"
        )
    )
    document = f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>材料统一处理报告</title>
<style>body{{font:15px/1.6 'Microsoft YaHei',sans-serif;max-width:1100px;margin:30px auto;color:#17202a}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #d9e2ec;padding:9px;text-align:left}}th{{background:#eef4ff}}section{{margin:24px 0}}.note,.review{{background:#fff7e6;padding:12px}}.name-alert{{background:#fff0f0;color:#a31515;border:2px solid #c62828;padding:12px}}.company-word-alert{{background:#ffe5e5;color:#8c1010;border:4px solid #c00000;padding:16px;box-shadow:0 0 0 3px #f4b4b4}}.company-word-alert h2{{margin-top:0}}.comparison{{border:1px solid #d9e2ec;padding:14px;margin:18px 0}}.comparison h3{{overflow-wrap:anywhere}}.pair{{display:flex;gap:16px;flex-wrap:wrap}}figure{{margin:0;flex:1 1 300px;min-width:0}}figure img{{max-width:100%;max-height:460px;object-fit:contain;border:1px solid #ddd}}figcaption{{font-weight:bold}}</style></head><body>
<h1>材料统一处理{'执行结果' if report['applied'] else '执行前预览'}</h1>
<p class="note">本次选择：{html.escape(selected_names)}。执行顺序：{html.escape(execution_order)}。未选步骤跳过；必要依赖和未同步导出提醒见下方。</p>
<p><b>自动依赖：</b>{html.escape(dependency_rows)}<br><b>一致性提醒：</b>{html.escape(warning_rows)}</p>
<p>状态：{html.escape(report['status'])}；失败/待处理项：{report['failures']}；输出目录：{html.escape(str(report.get('output') or '预览未创建'))}；输出名称：{html.escape(str(report.get('planned_output_label') or '未生成'))}（后接“_已处理_时间戳”）</p>
<table><thead><tr><th>阶段</th><th>状态</th><th>失败/待处理</th></tr></thead><tbody>{''.join(stage_rows)}</tbody></table>
<section><h2>软件名称一致性检查</h2>{name_warnings_html}</section>
<section><h2>标题与正文语义资格（全部项目）</h2><p>资格模式：{html.escape(str(qualification_metrics.get('mode', 'unknown')))}；Jev 成功返回 Score 后按唯一配置阈值二分为 MATCH/MISMATCH，置信度只用于监控。API、输入或配对故障显示为 NOT_EVALUATED。只有 MATCH 项目进入本次处理副本；MISMATCH 与 NOT_EVALUATED 留在原输入。Jev 的有序等级分数归一为 0–100，阈值初值尚未校准。{html.escape(qualification_summary)}</p>{mismatch_clipboard_html}<table><thead><tr><th>项目名称</th><th>Jev匹配分与置信度</th><th>检测结果</th><th>流程决定</th><th>理由、概率、token、耗时</th></tr></thead><tbody>{''.join(similarity_rows) or '<tr><td colspan="5">未扫描说明书</td></tr>'}</tbody></table></section>
<section><h2>未纳入项目</h2><p>MISMATCH 与 NOT_EVALUATED 项目不会复制到处理后文件夹，也不会进入后续修改；原始输入保持完整。</p><table><thead><tr><th>项目名称</th><th>流程决定</th><th>理由与正文证据</th></tr></thead><tbody>{''.join(disqualified_rows) or '<tr><td colspan="3">无</td></tr>'}</tbody></table></section>
{hollow_line_alert_html}{company_ocr_alert_html}<section><h2>说明书相关性、正文年份与图片年份</h2><p>标题正文资格使用配对 TXT 文件名与说明书正文证据；视觉复核仅提供人工审核证据，不参与现有 OCR 拦截及批量熔断。</p><p>{vision_details}</p><table><thead><tr><th>说明书</th><th>处理状态</th><th>标题正文资格</th><th>抽样图片 OCR</th><th>视觉复核</th><th>正文年份</th><th>图片年份</th></tr></thead><tbody>{''.join(manual_rows)}</tbody></table>{''.join(visual_rows)}</section>
<section><h2>TXT 明细（逐件）</h2><p>自动修改仅限主要功能字段；编码、字段名、功能字段以外的文字均保留。</p><table><thead><tr><th>文件</th><th>原编码</th><th>处理结果</th><th>检查结果</th><th>主要功能修改前后</th></tr></thead><tbody>{''.join(txt_rows) or '<tr><td colspan="5">未处理 TXT</td></tr>'}</tbody></table></section>
<section><h2>图片变化前后对比（待核验）</h2>{year_comparisons}</section>
<section><h2>封面、页眉、页码与说明书 PDF</h2><p>逐本显示名称来源、名称冲突、Word/PDF 页数、全书疑似空白页复核、修复尝试及项目快照回退情况。点击异常页截图可查看原尺寸。</p><table><thead><tr><th>项目与验收</th><th>计划动作</th><th>实际动作</th><th>仍需处理</th></tr></thead><tbody>{''.join(cover_rows) or '<tr><td colspan="4">未选择页面整理或 PDF 导出</td></tr>'}</tbody></table><div class="pair">{''.join(cover_evidence)}</div></section>
<section><h2>代码文件明细</h2><ul>{''.join(code_rows)}</ul></section>
</body></html>"""
    html_path = path.with_suffix(".html")
    html_temporary = html_path.with_suffix(html_path.suffix + ".tmp")
    html_temporary.write_text(document, encoding="utf-8")
    os.replace(html_temporary, html_path)


def _unique_output(source: Path, label: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = source.parent / f"{label}_已处理_{stamp}"
    suffix = 2
    while output.exists():
        output = source.parent / f"{label}_已处理_{stamp}_{suffix}"
        suffix += 1
    return output


def _output_customer_label(label: str, qualified_count: int, total_count: int) -> str:
    """Add the qualified/total project counts while preserving the customer name."""
    customer = re.sub(r"\d+\s*件$", "", label).rstrip(" _-")
    if not customer:
        customer = label.rstrip(" _-")
    return f"{customer}{qualified_count}件（共{total_count}件）"


def _relative_material_path(root: Path, value: Any) -> Path | None:
    raw = str(value or "").strip().replace("\\", "/")
    if not raw:
        return None
    posix = PurePosixPath(raw)
    if posix.is_absolute() or any(part in {"..", ""} for part in posix.parts):
        return None
    relative = Path(*posix.parts)
    try:
        (root / relative).resolve().relative_to(root.resolve())
    except (OSError, ValueError):
        return None
    return relative


def _qualified_material_plan(
    source_root: Path,
    projects: list[dict[str, Any]],
) -> tuple[list[Path], set[str], set[str], int, int]:
    """Plan a safe MATCH-only copy, retaining complete isolated project folders."""
    rows: list[tuple[dict[str, Any], str, Path | None]] = []
    for item in projects:
        if not isinstance(item, dict):
            continue
        status = str(item.get("semantic_status", item.get("status", "NOT_EVALUATED"))).upper()
        relative = _relative_material_path(source_root, item.get("relative_path"))
        rows.append((item, status, relative))

    matched = [row for row in rows if row[1] == "MATCH"]
    excluded_files: set[str] = set()
    matched_files: set[str] = set()
    for item, status, _ in rows:
        raw_files = item.get("associated_files")
        if not isinstance(raw_files, (list, tuple)) or not raw_files:
            raw_files = [item.get("relative_path")]
        target = matched_files if status == "MATCH" else excluded_files
        for raw in raw_files:
            relative = _relative_material_path(source_root, raw)
            if relative is not None:
                target.add(relative.as_posix())

    # A project directory is a safe copy unit only when every evaluated manual
    # beneath it passed. Shared/flat directories fall back to the exact paired
    # file set so neighboring rejected projects cannot leak into the output.
    candidate_dirs: set[Path] = set()
    for _item, status, relative in matched:
        if status != "MATCH" or relative is None or relative.parent == Path("."):
            continue
        project_dir = relative.parent
        descendants = [
            other for other in rows
            if other[2] is not None and other[2].is_relative_to(project_dir)
        ]
        if descendants and all(other[1] == "MATCH" for other in descendants):
            candidate_dirs.add(project_dir)

    copy_dirs: list[Path] = []
    for directory in sorted(candidate_dirs, key=lambda path: (len(path.parts), path.as_posix().casefold())):
        if not any(parent == directory or parent in directory.parents for parent in copy_dirs):
            copy_dirs.append(directory)

    qualified_count = 0
    for item, status, manual_path in matched:
        raw_files = item.get("associated_files")
        if not isinstance(raw_files, (list, tuple)) or not raw_files:
            raw_files = [item.get("relative_path")]
        candidates = [manual_path] if manual_path is not None else []
        candidates.extend(
            relative for raw in raw_files
            if (relative := _relative_material_path(source_root, raw)) is not None
        )
        if any(
            relative.as_posix() not in excluded_files
            and (
                relative.as_posix() in matched_files
                or any(relative.is_relative_to(directory) for directory in copy_dirs)
            )
            and (source_root / relative).is_file()
            for relative in candidates
        ):
            qualified_count += 1

    return copy_dirs, matched_files, excluded_files, qualified_count, len(rows)


def _path_is_qualified(
    path: Path,
    source_root: Path,
    copy_dirs: list[Path],
    matched_files: set[str],
    excluded_files: set[str],
) -> bool:
    try:
        relative = path.resolve().relative_to(source_root.resolve())
    except (OSError, ValueError):
        return False
    relative_name = relative.as_posix()
    if relative_name in excluded_files:
        return False
    return relative_name in matched_files or any(relative.is_relative_to(directory) for directory in copy_dirs)


def _copy_qualified_materials(
    source_root: Path,
    output_root: Path,
    plan: tuple[list[Path], set[str], set[str], int, int],
) -> list[str]:
    """Copy only the files selected by a validated MATCH-only project plan."""
    copy_dirs, matched_files, excluded_files, _qualified_count, _total_count = plan

    output_root.mkdir(parents=True, exist_ok=False)
    copied: set[str] = set()
    for relative_dir in copy_dirs:
        source_dir = (source_root / relative_dir).resolve()
        try:
            source_dir.relative_to(source_root.resolve())
        except ValueError:
            continue
        if not source_dir.is_dir():
            continue

        def ignore_excluded(directory: str, names: list[str]) -> list[str]:
            ignored = []
            base = Path(directory)
            for name in names:
                candidate = base / name
                if candidate.is_file():
                    relative = candidate.relative_to(source_root).as_posix()
                    if relative in excluded_files:
                        ignored.append(name)
                    else:
                        copied.add(relative)
            return ignored

        destination_dir = output_root / relative_dir
        destination_dir.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(source_dir, destination_dir, dirs_exist_ok=True, ignore=ignore_excluded)

    for raw in sorted(matched_files - excluded_files):
        if any(Path(raw).is_relative_to(directory) for directory in copy_dirs):
            continue
        relative = Path(raw)
        source_file = (source_root / relative).resolve()
        try:
            source_file.relative_to(source_root.resolve())
        except ValueError:
            continue
        if not source_file.is_file():
            continue
        destination_file = output_root / relative
        destination_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_file, destination_file)
        copied.add(relative.as_posix())
    return sorted(copied)


def _recover_processed_outputs(
    source_path: Path,
    legacy_label: str,
    customer_label: str,
    total_count: int,
) -> list[dict[str, Any]]:
    """Recover pending cover writes in both legacy and counted output folders."""
    from validation_isolation import require_validation_path

    output_prefixes = {f"{legacy_label}_已处理_"}
    output_prefixes.update(
        f"{customer_label}{count}件（共{total_count}件）_已处理_"
        for count in range(max(0, total_count) + 1)
    )
    recovered = process_batch.recover_pending_cover_transactions(source_path)
    if not source_path.parent.is_dir():
        return recovered
    for sibling in source_path.parent.iterdir():
        if (
            not sibling.is_dir()
            or not any(sibling.name.startswith(prefix) for prefix in output_prefixes)
            or not (sibling / process_batch.COVER_TRANSACTION_FILENAME).is_file()
        ):
            continue
        checked = require_validation_path(sibling, kind="material", same_run_as=source_path)
        recovered.extend(process_batch.recover_pending_cover_transactions(checked))
    return recovered


def run(
    source: str | Path,
    work: str | Path,
    app_root: str | Path,
    apply: bool = False,
    progress: Any = None,
    *,
    manual_options: dict[str, Any] | None = None,
    document_stages: str | list[str] | tuple[str, ...] | None = None,
    skip_topic_mismatch: bool = True,
    expand_order_downloads: bool = False,
    clipboard_writer: Any = None,
) -> dict[str, Any]:
    from validation_isolation import require_validation_path

    requested_options = _selected_stages(document_stages)
    effective = set(requested_options)
    auto_dependencies: list[dict[str, str]] = []
    if "match_manual_vision" in effective and "match_manual_ocr" not in effective:
        effective.add("match_manual_ocr")
        auto_dependencies.append({"option": "match_manual_ocr", "reason": "视觉复核依赖 OCR 筛选候选图片"})
    if "manual_image_dates" in effective and "match_manual_ocr" not in effective:
        effective.add("match_manual_ocr")
        auto_dependencies.append({"option": "match_manual_ocr", "reason": "图片日期处理前需要运行截图相关性预检"})
    if "manual_cover" in effective and "manual_pdf" not in effective:
        effective.add("manual_pdf")
        auto_dependencies.append({"option": "manual_pdf", "reason": "封面修复需要导出并逐页验收与 DOCX 同步的 PDF"})
    needs_qualification_gate = skip_topic_mismatch and bool(effective - {"match_software_name"})
    if needs_qualification_gate and "match_title_body" not in effective:
        effective.add("match_title_body")
        auto_dependencies.append({"option": "match_title_body", "reason": "默认资格门禁需要筛出标题正文主题冲突项目"})
    selected_stages = tuple(option for option in WORKFLOW_OPTION_ORDER if option in effective)
    manual_content_options = {"match_title_body", "match_manual_ocr", "match_manual_vision", "manual_body_years", "manual_image_dates", "manual_ai_hollow_lines"}
    manual_format_options = {"manual_cover", "manual_header", "manual_page_numbers", "manual_pdf"}
    code_options = {"match_code", "code_cleanup", "code_pdf"}
    txt_options = {"match_txt", "txt_repair"}
    manual_content_selected = bool(effective.intersection(manual_content_options - {"match_title_body"}))
    manual_format_selected = bool(effective.intersection(manual_format_options))
    manual_selected = manual_content_selected or manual_format_selected
    code_selected = bool(effective.intersection(code_options))
    txt_selected = bool(effective.intersection(txt_options))
    source_path = Path(source).resolve()
    work_path = Path(work).resolve()
    app_root_path = Path(app_root).resolve()
    source_path = require_validation_path(source, kind="material")
    require_validation_path(work_path, kind="workspace")
    work_path.mkdir(parents=True, exist_ok=True)
    material_root, label = _material_root(source_path, work_path, app_root_path, expand_order_downloads=expand_order_downloads)
    all_manuals = list(material_root.rglob("*说明.docx"))
    qualification_requested = "match_title_body" in effective or needs_qualification_gate
    name_check_requested = "match_software_name" in effective
    scan_requested = qualification_requested or name_check_requested
    source_manual_count = len(all_manuals)
    manual_count = source_manual_count if manual_selected or scan_requested else 0
    qualification_count = source_manual_count if scan_requested else 0
    code_count = len(_code_inventory(material_root, inspect=False)) if code_selected else 0
    txt_count = sum(path.is_file() and path.suffix.casefold() == ".txt" for path in material_root.rglob("*")) if txt_selected else 0
    total = manual_count + code_count + (max(1, txt_count) if txt_selected else 0)
    if progress is not None and hasattr(progress, "set_total"):
        progress.set_total(total)
    _publish_progress(progress, "材料清单已就绪", 0, failures=0, active=False)

    manual_stage: dict[str, Any] = {"stage": "manual", "status": "SKIPPED", "failures": 0, "options": []}
    manual_report: dict[str, Any] = {"stage": "manual", "status": "SKIPPED", "projects": [], "failures": 0}
    manual_failures = 0
    excluded_projects: list[dict[str, Any]] = []
    title_body_inputs: dict[str, dict[str, Any]] = {}
    title_body_path, title_body_config = _title_body_config(app_root_path)
    title_body_mode = str(title_body_config.get("mode", "enforced")).casefold()
    skip_mismatch_requested = bool(title_body_config.get("skip_mismatch_requested", True)) and bool(skip_topic_mismatch)
    not_evaluated_action = str(title_body_config.get("not_evaluated_action", "hold_changes")).casefold()
    qualification_report: dict[str, Any] = {
        "status": "NOT_RUN", "mode": title_body_mode, "scanned_count": 0,
        "failure_count": 0, "failures": [], "disqualified_projects": [],
        "cache_hit_count": 0, "model_request_count": 0,
    }
    output_copy_plan: tuple[list[Path], set[str], set[str], int, int] | None = None
    if scan_requested:
        qualification_report = _scan_title_body_qualification(
            material_root,
            progress=progress if not manual_content_selected else None,
            work_path=work_path,
            app_root=app_root_path,
            config_path=title_body_path,
            progress_current=0 if manual_content_selected else None,
            check_names=name_check_requested,
            local_only=not qualification_requested,
        )
        title_body_inputs = qualification_report.pop("_inputs_by_relative_path", {})
        qualification_report["skip_mismatch_requested"] = skip_mismatch_requested
        qualification_report["not_evaluated_action"] = not_evaluated_action
        projects = qualification_report.get("projects", [])
        if not isinstance(projects, list):
            projects = []
            qualification_report["projects"] = projects
        match_only_output = qualification_requested and bool(skip_topic_mismatch)
        if match_only_output:
            output_copy_plan = _qualified_material_plan(material_root, projects)
            qualified_count = output_copy_plan[3]
            total_project_count = output_copy_plan[4]
        else:
            qualified_count = sum(
                str(item.get("semantic_status", item.get("status", "NOT_EVALUATED"))).upper() == "MATCH"
                for item in projects if isinstance(item, dict)
            )
            total_project_count = len([item for item in projects if isinstance(item, dict)])
        qualification_report["qualified_project_count"] = qualified_count
        qualification_report["input_project_count"] = total_project_count
        qualification_report["output_selection"] = "MATCH_ONLY" if match_only_output else "ALL_BY_REQUEST"
        qualification_report["skip_mismatch_effective"] = match_only_output
        for item in qualification_report.get("projects", []):
            if not qualification_requested:
                continue
            semantic_status = str(item.get("semantic_status", item.get("status", "NOT_EVALUATED"))).upper()
            if match_only_output and semantic_status == "MISMATCH":
                action = "SKIP_TITLE_BODY_MISMATCH"
            elif match_only_output and semantic_status != "MATCH":
                action = "HOLD_NOT_EVALUATED"
            elif title_body_mode == "off":
                action = "SKIP_TITLE_BODY_MISMATCH" if item.get("legacy_status") == "review_topic_mismatch" and skip_topic_mismatch else "CONTINUE_WITH_REVIEW"
            elif title_body_mode == "enforced" and semantic_status == "MISMATCH" and skip_mismatch_requested:
                action = "SKIP_TITLE_BODY_MISMATCH"
            elif title_body_mode == "enforced" and semantic_status == "NOT_EVALUATED" and not_evaluated_action == "hold_changes":
                action = "HOLD_NOT_EVALUATED"
            else:
                action = "CONTINUE_WITH_REVIEW"
            item["action"] = action
            if action in {"SKIP_TITLE_BODY_MISMATCH", "HOLD_NOT_EVALUATED"}:
                excluded_item: dict[str, Any] = {
                    "name": str(item.get("name", "")),
                    "folder": Path(str(item.get("relative_path", ""))).parent.as_posix(),
                    "action": action,
                }
                associated_files = list(item.get("associated_files", []))
                if associated_files:
                    excluded_item["associated_files"] = associated_files
                excluded_projects.append(excluded_item)
        if not qualification_report.get("disqualified_projects"):
            qualification_report["disqualified_projects"] = [
                item for item in qualification_report.get("projects", [])
                if item.get("semantic_status", item.get("status")) == "MISMATCH"
            ]
        all_selected_code_rows = _code_inventory(material_root, inspect=False) if code_selected else []
        if output_copy_plan is not None:
            copy_dirs, matched_files, excluded_files, _qualified_count, _total_count = output_copy_plan
            manual_count = qualified_count if manual_selected or qualification_requested else 0
            code_count = sum(
                _path_is_qualified(
                    Path(str(row.get("docx") or row.get("pdf") or "")), material_root,
                    copy_dirs, matched_files, excluded_files,
                )
                for row in all_selected_code_rows
                if row.get("docx") or row.get("pdf")
            )
            txt_count = sum(
                path.is_file() and path.suffix.casefold() == ".txt"
                and _path_is_qualified(path, material_root, copy_dirs, matched_files, excluded_files)
                for path in material_root.rglob("*")
            ) if txt_selected else 0
            total = qualification_count + (code_count if code_selected else 0) + (max(1, txt_count) if txt_selected else 0)
            if progress is not None and hasattr(progress, "set_total"):
                progress.set_total(total)
        else:
            total_project_count = source_manual_count
            qualified_count = source_manual_count
            manual_count = source_manual_count if manual_selected or scan_requested else 0
            total = manual_count + code_count + (max(1, txt_count) if txt_selected else 0)
            if progress is not None and hasattr(progress, "set_total"):
                progress.set_total(total)

    else:
        total_project_count = source_manual_count
        qualified_count = source_manual_count

    output_label = _output_customer_label(label, qualified_count, total_project_count)
    customer_label = re.sub(r"\d+\s*件$", "", label).rstrip(" _-") or label.rstrip(" _-")
    recovered_cover_transactions = _recover_processed_outputs(
        source_path, label, customer_label, total_project_count,
    )
    planned_output = _unique_output(source_path, output_label) if apply else None
    if planned_output is not None:
        require_validation_path(planned_output, kind="material", same_run_as=source_path)

    mismatch_names = _mismatch_names(qualification_report.get("projects", []))
    mismatch_clipboard = _copy_mismatch_summary(
        mismatch_names, apply=apply, clipboard_writer=clipboard_writer,
    )
    qualification_report["mismatch_names"] = mismatch_names
    qualification_report["mismatch_clipboard"] = mismatch_clipboard
    progress_name = mismatch_clipboard["text"] or "未发现 MISMATCH 项目"
    name_warning_names = qualification_report.get("name_warning_names", [])
    if name_warning_names:
        progress_name += "\n软件名称不一致，请复核：\n" + "\n".join(dict.fromkeys(name_warning_names))
    if mismatch_names and not apply:
        progress_name = f"预览（未复制剪贴板）：\n{progress_name}"
    if scan_requested:
        _publish_progress(
            progress,
            "标题正文资格预检完成" if qualification_requested else "软件名称一致性检查完成",
            qualification_count,
            progress_name,
            failures=int(qualification_report.get("failure_count", 0) or 0),
            error=mismatch_clipboard.get("error"),
            active=False,
        )

    output: Path | None = None
    processing_root = material_root
    if apply:
        output = planned_output
        assert output is not None
        if output_copy_plan is not None:
            _copy_qualified_materials(material_root, output, output_copy_plan)
        else:
            shutil.copytree(material_root, output)
        processing_root = output

    if manual_content_selected:
        options = dict(manual_options or {})
        for reserved in ("input_root", "output_root", "report_path", "progress", "apply"):
            options.pop(reserved, None)
        options["skip_topic_mismatch"] = skip_topic_mismatch
        options["check_body_relevance"] = "match_title_body" in effective
        options["check_image_relevance"] = "match_manual_ocr" in effective
        options["check_image_vision"] = "match_manual_vision" in effective
        options["rewrite_body_years"] = "manual_body_years" in effective
        options["rewrite_image_dates"] = "manual_image_dates" in effective
        options["check_ai_hollow_lines"] = "manual_ai_hollow_lines" in effective
        options["title_body_inputs_by_relative_path"] = title_body_inputs if qualification_requested else None
        options["title_body_mode"] = title_body_mode
        options["skip_mismatch"] = skip_mismatch_requested
        options["not_evaluated_action"] = not_evaluated_action
        manual_report_path = work_path / ("combined_manual_result.json" if apply else "combined_manual_preview.json")
        manual_report = manual_content_stage.run_manual_content_stage(
            processing_root, processing_root, manual_report_path,
            progress=_StageProgress(progress, 0, floor_current=qualification_count), apply=apply, **options,
        )
        for project in manual_report.get("projects", []):
            if project.get("qualification_action") in {"SKIP_MISMATCH", "HOLD_NOT_EVALUATED"}:
                project["status"] = "SKIPPED_TOPIC_MISMATCH" if project.get("qualification_action") == "SKIP_MISMATCH" else "HELD_QUALIFICATION_NOT_EVALUATED"
    manual_failures = (
        _manual_failure_count(manual_report, apply) if manual_content_selected else 0
    ) + int(qualification_report.get("failure_count", 0) or 0)

    if manual_format_selected:
        cover_plan_path = work_path / ("combined_cover_plan_result.json" if apply else "combined_cover_plan_preview.json")
        cover_plan = _build_manual_cover_plan(processing_root, cover_plan_path, excluded_projects, selected_stages)
        cover_report: dict[str, Any] = {
            "status": "READY_FOR_REVIEW", "packages": cover_plan.get("packages", []),
            "audit_errors": cover_plan.get("audit_errors", []),
        }
        cover_failures = len(cover_plan.get("audit_errors", []))
        if apply and cover_plan.get("packages"):
            _publish_progress(progress, "说明书Word封面与PDF处理", 0, failures=cover_failures, active=True)
            cover_report = process_batch.invoke_word_pipeline(
                processing_root, cover_plan_path, work_path / "combined_cover_result.json",
                SCRIPT_DIR / "word_manual_pipeline_stable.ps1", progress=progress, module="说明书处理",
            )
            cover_failures += _cover_result_failures(cover_report)
            _publish_progress(progress, "说明书Word封面与PDF处理完成", manual_count, failures=cover_failures, active=False)
        manual_failures = (
            (_manual_failure_count(manual_report, apply) if manual_content_selected else 0)
            + int(qualification_report.get("failure_count", 0) or 0)
            + cover_failures
        )
        manual_stage = {
            **(manual_report if manual_content_selected else {"stage": "manual", "projects": [], "manual_count": manual_count}),
            "status": "PARTIAL" if manual_failures else ("APPLIED" if apply else "READY_FOR_REVIEW"),
            "failures": manual_failures, "cover_plan": cover_plan, "cover_result": cover_report,
        }
        manual_stage["options"] = [key for key in selected_stages if key in manual_format_options]
    elif manual_content_selected:
        manual_stage = {
            **manual_report,
            "status": "PARTIAL" if manual_failures else ("APPLIED" if apply else "READY_FOR_REVIEW"),
            "failures": manual_failures,
            "options": [key for key in selected_stages if key in manual_content_options],
        }

    code_stage: dict[str, Any] = {"stage": "code", "status": "SKIPPED", "failures": 0, "options": []}
    code_failures = 0
    if code_selected:
        code_stage, code_failures = _run_code_stage(
            processing_root, app_root_path, work_path, apply=apply, progress=progress,
            offset=qualification_count, failures=manual_failures, excluded=excluded_projects,
            inspect="match_code" in effective, cleanup="code_cleanup" in effective,
            export_pdf="code_pdf" in effective,
        )
        code_stage["options"] = [key for key in selected_stages if key in code_options]
    txt_stage: dict[str, Any] = {"stage": "txt", "status": "SKIPPED", "failures": 0, "options": []}
    if txt_selected:
        txt_stage = _run_txt_stage(
            processing_root, apply=apply, progress=progress, offset=qualification_count + code_count,
            failures=manual_failures + code_failures,
            check="match_txt" in effective, repair="txt_repair" in effective, excluded=excluded_projects,
        )
        txt_stage["options"] = [key for key in selected_stages if key in txt_options]
    failures = manual_failures + code_failures + int(txt_stage.get("failures", 0))
    status = PREVIEW_STATUS if not apply else ("DOCUMENT_PROCESSING_PARTIAL" if failures else "DOCUMENT_PROCESSING_OK")
    company_ocr_alerts = _company_ocr_alerts(manual_stage, selected_stages)
    report = {
        "version": VERSION,
        "created_at": datetime.now().astimezone().isoformat(),
        "source": str(source_path),
        "material_root": str(material_root),
        "recovered_cover_transactions": recovered_cover_transactions,
        "applied": apply,
        "status": status,
        "output": str(output) if output else None,
        "planned_output_label": output_label,
        "output_counts": {
            "qualified_projects": qualified_count,
            "total_projects": total_project_count,
            "selection": qualification_report.get("output_selection", "ALL_BY_REQUEST"),
        },
        "failures": failures,
        "stage_order": ["标题与说明书匹配资格预检", "说明书内容检测及所选修改", "说明书封面/页眉/页码整理及 PDF 导出", "代码检测/注释清理及 PDF 导出", "TXT 检查与修复", "统一报告"],
        "selected_stages": list(requested_options),
        "effective_stages": list(selected_stages),
        "auto_dependencies": auto_dependencies,
        "dependency_warnings": _workflow_dependency_warnings(selected_stages),
        "qualification": {
            "skip_topic_mismatch_enabled": bool(qualification_report.get("skip_mismatch_effective", False)),
            "skip_mismatch_requested": bool(qualification_report.get("skip_mismatch_requested", False)),
            "skip_mismatch_effective": bool(qualification_report.get("skip_mismatch_effective", False)),
            "not_evaluated_action": not_evaluated_action,
            "disqualified_count": len(qualification_report.get("disqualified_projects", [])),
            "skipped_count": sum(item.get("action") == "SKIP_TITLE_BODY_MISMATCH" for item in excluded_projects),
            "held_count": sum(item.get("action") == "HOLD_NOT_EVALUATED" for item in excluded_projects),
            **qualification_report,
        },
        "company_ocr_alerts": company_ocr_alerts,
        "ai_hollow_line_alerts": [
            {"name": project.get("title") or Path(str(project.get("relative_path", "")).replace("\\", "/")).stem.removesuffix("说明"), "relative_path": project.get("relative_path"), "message": project.get("ai_hollow_line_alert")}
            for project in manual_stage.get("projects", []) if project.get("ai_hollow_line_findings")
        ],
        "stages": {"manual": manual_stage, "code": code_stage, "txt": txt_stage},
    }
    report_path = work_path / ("combined_document_result.json" if apply else "combined_document_preview.json")
    _write_report(report_path, report)
    _publish_progress(
        progress,
        "材料统一处理完成" if apply else "材料统一预检完成",
        total,
        "软件名称不一致，请复核：" + "、".join(name_warning_names) if name_warning_names else "",
        status="FAILED" if failures else "DONE",
        failures=failures,
        active=False,
    )
    return {
        "status": status,
        "report": str(report_path.with_suffix(".html")),
        "report_json": str(report_path),
        "output": str(output) if output else None,
        "outputs": [str(output)] if output else [],
        "failures": failures,
        "work_dir": str(work_path),
        "name_warning_count": int(qualification_report.get("name_warning_count", 0)),
        "name_warning_names": name_warning_names,
    }


def preview(
    source: str | Path,
    work: str | Path,
    app_root: str | Path,
    progress: Any = None,
    *,
    manual_options: dict[str, Any] | None = None,
    document_stages: str | list[str] | tuple[str, ...] | None = None,
    skip_topic_mismatch: bool = True,
) -> dict[str, Any]:
    return run(source, work, app_root, apply=False, progress=progress,
               manual_options=manual_options, document_stages=document_stages,
               skip_topic_mismatch=skip_topic_mismatch)


def execute(
    source: str | Path,
    work: str | Path,
    app_root: str | Path,
    progress: Any = None,
    *,
    manual_options: dict[str, Any] | None = None,
    document_stages: str | list[str] | tuple[str, ...] | None = None,
    skip_topic_mismatch: bool = True,
) -> dict[str, Any]:
    return run(source, work, app_root, apply=True, progress=progress,
               manual_options=manual_options, document_stages=document_stages,
               skip_topic_mismatch=skip_topic_mismatch)


__all__ = ["PREVIEW_STATUS", "execute", "preview", "run"]

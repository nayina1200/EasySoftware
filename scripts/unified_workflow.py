from __future__ import annotations

import argparse
import contextlib
import ctypes
import difflib
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import contextlib
from pathlib import Path

import archive_types

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")

VERSION = "3.6.0"
HERE = Path(__file__).resolve().parent
PROCESS = HERE / "process_batch.py"
PACKAGE = HERE / "package_approved_modified.py"
EXPORT_PDFS = HERE / "export_modified_pdfs.ps1"
WORD_AUTOMATION_MUTEX = "Local\\EasySoftware.WordAutomation.v1"


def decode_process_output(value: str | bytes | None) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return value.decode(encoding)
        except UnicodeDecodeError:
            continue
    return value.decode("gb18030", errors="replace")


def process_is_running(pid: object) -> bool:
    try:
        os.kill(int(pid), 0)
    except (TypeError, ValueError, OSError):
        return False
    return True


@contextlib.contextmanager
def word_automation_lock(progress: ProgressReporter | None, module: str, timeout_seconds: int = 180):
    """Serialize Word COM work across multiple EasySoftware windows."""
    if os.name != "nt":
        yield
        return
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.WaitForSingleObject.argtypes = (ctypes.c_void_p, ctypes.c_uint32)
    kernel32.ReleaseMutex.argtypes = (ctypes.c_void_p,)
    kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
    handle = kernel32.CreateMutexW(None, False, WORD_AUTOMATION_MUTEX)
    if not handle:
        raise RuntimeError("无法创建 Word 处理队列")
    acquired = False
    deadline = time.monotonic() + timeout_seconds
    try:
        while True:
            result = kernel32.WaitForSingleObject(handle, 5_000)
            if result in (0, 0x80):
                acquired = True
                break
            if result == 0x102:
                if progress:
                    progress.update("等待其他任务完成 Word 导出", 0, module or "Word处理")
                if time.monotonic() >= deadline:
                    raise TimeoutError(f"等待 Word 队列超过 {timeout_seconds} 秒；请关闭重复的 EasySoftware 任务后重试")
                continue
            raise RuntimeError("等待 Word 处理队列时发生异常")
        yield
    finally:
        if acquired:
            kernel32.ReleaseMutex(handle)
        kernel32.CloseHandle(handle)


class ProgressReporter:
    """Write the small progress contract consumed by the desktop GUI."""

    def __init__(self, work: Path, total: int = 0):
        self.path = work / "progress.json"
        self.total = total
        self.started = time.perf_counter()
        self.completed = 0
        self.failures = 0

    def set_total(self, total: int) -> None:
        """Publish a stable denominator after a workflow prepares its task list."""
        self.total = max(0, total)

    def update(
        self, phase: str, current: int = 0, name: str = "",
        status: str = "RUNNING", error: str | None = None, *,
        active: bool | None = None, failures: int | None = None,
        archive: str = "",
    ) -> None:
        self.completed = max(self.completed, current)
        if failures is not None:
            self.failures = max(self.failures, failures)
        current = self.completed
        elapsed = time.perf_counter() - self.started
        rate = elapsed / self.completed if self.completed else 0
        row = {
            "version": VERSION,
            "status": status,
            "phase": phase,
            "current": current,
            "total": self.total,
            "project_name": name,
            "archive_name": archive,
            "elapsed_seconds": round(elapsed, 1),
            "completed_projects": min(self.total, self.completed),
            "remaining_projects": max(0, self.total - self.completed),
            "estimated_remaining_seconds": round(rate * max(0, self.total - self.completed), 1) if rate else None,
            "current_module": phase,
            "current_project_active": status == "RUNNING" if active is None else active,
            "failures": self.failures,
            "updated_at": __import__("datetime").datetime.now().astimezone().isoformat(),
        }
        if error:
            row["error"] = error
        atomic_json(self.path, row)


def emit(value: dict, code: int = 0) -> int:
    print(json.dumps(value, ensure_ascii=False, separators=(",", ":")))
    return code


def load_json(path: Path, default=None):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return default


def atomic_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def resolve_modified_directory(batch_root: Path, source: Path | None = None) -> Path:
    """定位修改版目录：优先返回真实存在的目录，不把已解压目录命名为
    “解压版”作为前提；名称约定只作为最后的兜底猜测。"""
    candidates: list[Path] = []
    if source is not None and source.is_dir():
        if source.name.endswith("修改版"):
            candidates.append(source)
        name = source.name
        if name.endswith("_解压版"):
            candidates.append(batch_root / f"{name[:-4]}_修改版")
        elif name.endswith("解压版"):
            candidates.append(batch_root / f"{name[:-3]}修改版")
        candidates.append(batch_root / f"{source.stem}_修改版")
        candidates.append(source / "修改版")
    candidates.append(batch_root / "修改版")
    candidates.append(batch_root.parent / f"{batch_root.stem}_修改版")
    candidates.append(batch_root.parent / "修改版")
    seen: set[Path] = set()
    ordered: list[Path] = []
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved not in seen:
            seen.add(resolved)
            ordered.append(candidate)
    for candidate in ordered:
        if candidate.is_dir() and any(candidate.iterdir()):
            return candidate
    return ordered[0]


def default_work_dir(batch_root: Path) -> Path:
    digest = hashlib.sha1(str(batch_root.resolve()).lower().encode("utf-8")).hexdigest()[:16]
    return Path(tempfile.gettempdir()) / "codex-softcopyright" / digest


def python_has_dependencies(executable: Path) -> bool:
    result = subprocess.run(
        [str(executable), "-c", "import lxml,PIL,pypdf"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return result.returncode == 0


def find_python(explicit: str | None) -> Path | None:
    home = Path.home()
    candidates = [
        Path(explicit) if explicit else None,
        Path(os.environ["SOFTCOPYRIGHT_PYTHON"]) if os.environ.get("SOFTCOPYRIGHT_PYTHON") else None,
        Path(sys.executable),
        home / ".cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe",
        home / ".cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python",
    ]
    seen = set()
    for candidate in candidates:
        if candidate is None:
            continue
        candidate = candidate.resolve()
        if candidate in seen or not candidate.is_file():
            continue
        seen.add(candidate)
        if python_has_dependencies(candidate):
            return candidate
    return None


def parse_result(result: subprocess.CompletedProcess[str]) -> dict:
    for stream in (result.stdout, result.stderr):
        for line in reversed((stream or "").splitlines()):
            try:
                value = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(value, dict):
                value.setdefault("returncode", result.returncode)
                return value
    return {
        "status": "FAILED",
        "returncode": result.returncode,
        "error": (result.stderr or result.stdout or "subprocess failed")[-3000:],
    }


def run_json(command: list[str]) -> dict:
    result = subprocess.run(
        command,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return parse_result(result)


def detect_format(batch_root: Path) -> str | None:
    # 交付格式只有 ZIP/RAR；7z 等只读输入格式不参与交付格式判断，
    # 否则纯 7z 批次会推导出无法交付的格式。
    archives = sorted(
        (path for path in batch_root.iterdir()
         if path.is_file() and path.suffix.lower() in archive_types.DELIVERY_SUFFIXES),
        key=lambda path: (path.stem.casefold() != batch_root.name.casefold(), -path.stat().st_mtime_ns),
    )
    formats = {path.suffix.lower().lstrip(".") for path in archives}
    if archives and archives[0].stem.casefold() == batch_root.name.casefold():
        return archives[0].suffix.lower().lstrip(".")
    return next(iter(formats)) if len(formats) == 1 else None


def prepare_archive_workspace(archive: Path) -> Path:
    """Place a selected archive in its own same-named folder before processing."""
    archive = archive.resolve()
    # Recover archives created by the 3.3.0 folder-comparison bug:
    # <name>/<name>/<name>.rar -> <name>/<name>.rar.
    while archive.parent.name.casefold() == archive.stem.casefold() and archive.parent.parent.name.casefold() == archive.stem.casefold():
        target = archive.parent.parent / archive.name
        if target.exists():
            if hashlib.sha256(target.read_bytes()).digest() != hashlib.sha256(archive.read_bytes()).digest():
                raise ValueError(f"同名文件夹中已有不同的压缩包：{target}")
            archive.unlink()
        else:
            shutil.move(str(archive), str(target))
        empty_folder = archive.parent
        if not any(empty_folder.iterdir()):
            empty_folder.rmdir()
        archive = target.resolve()
    target_folder = archive.parent / archive.stem
    if archive.parent.name.casefold() == archive.stem.casefold():
        return archive
    target_folder.mkdir(exist_ok=True)
    target = target_folder / archive.name
    if target.exists():
        if hashlib.sha256(target.read_bytes()).digest() != hashlib.sha256(archive.read_bytes()).digest():
            raise ValueError(f"同名文件夹中已有不同的压缩包：{target}")
        archive.unlink()
    else:
        shutil.move(str(archive), str(target))
    return target.resolve()


def parse_indexes(value: str, maximum: int) -> list[int]:
    text = value.strip().lower()
    if text in {"all", "全部", "全选"}:
        return list(range(1, maximum + 1))
    indexes: set[int] = set()
    for token in re.split(r"[,，、;；\s]+", text):
        if not token:
            continue
        match = re.fullmatch(r"(\d+)\s*[-~—至]\s*(\d+)", token)
        if match:
            start, end = map(int, match.groups())
            indexes.update(range(min(start, end), max(start, end) + 1))
        elif token.isdigit():
            indexes.add(int(token))
        else:
            raise ValueError(f"无法识别项目序号：{token}")
    if not indexes or min(indexes) < 1 or max(indexes) > maximum:
        raise ValueError(f"项目序号应在 1-{maximum} 范围内")
    return sorted(indexes)


def markdown_link(path: Path, label: str | None = None) -> str:
    path = path.resolve()
    return f"[{label or path.name}](<{path}>)"


def file_uri(value: str | None) -> str | None:
    if not value:
        return None
    path = Path(value).resolve()
    return path.as_uri() if path.is_file() else None


def code_header_repairs(preview: dict) -> list[str]:
    repairs = []
    for row in preview.get("packages", []):
        code = row.get("code_pdf") or {}
        # PDF text extraction often omits positioned PAGE fields. Do not rewrite a
        # code document solely because an extracted page number is absent.
        checks = (
            code.get("first_has_title"), code.get("last_has_title"),
            code.get("first_has_version"), code.get("last_has_version"),
            code.get("first_has_page_number"), code.get("last_has_page_number"),
        )
        if code and not all(checks):
            repairs.append(str(row["folder"]))
    return repairs


def findings_html(items: list[dict], kind: str) -> str:
    if not items:
        return f'<p class="ok">{html.escape(kind)}：程序未发现异常</p>'
    rows = []
    for item in items:
        where = " / ".join(str(item.get(key, "")) for key in ("file", "location") if item.get(key))
        rows.append(f"<li><b>{html.escape(where)}</b>：{html.escape(str(item.get('message', '')))}</li>")
    return f'<div class="findings"><h4>{html.escape(kind)}（{len(items)}）</h4><ul>{"".join(rows)}</ul></div>'


def image_panel(label: str, value: str | None) -> str:
    uri = file_uri(value)
    if not uri:
        return f'<section class="visual missing"><h4>{html.escape(label)}</h4><p>未生成审核图</p></section>'
    filename = html.escape(Path(value).name)
    escaped_uri = html.escape(uri, quote=True)
    return (
        f'<section class="visual"><h4>{html.escape(label)}</h4>'
        f'<a href="{escaped_uri}" target="_blank">打开原图：{filename}</a>'
        f'<a href="{escaped_uri}" target="_blank"><img loading="lazy" src="{escaped_uri}" alt="{html.escape(label)}"></a></section>'
    )


def result_cell(issue: str, actions: str, final: str) -> str:
    passed = final.strip() in {
        "原始TXT检查通过",
        "原代码检查通过",
        "原说明书检查通过",
        "最终检查通过",
        "封面、页眉和页码检查通过",
    }
    status = "通过" if passed else "未通过"
    return (
        f'<div class="result-cell"><span class="status {"pass" if passed else "fail"}">{status}</span><br><b>发现：</b>{html.escape(issue)}<br>'
        f'<b>修复：</b>{html.escape(actions)}<br><b>验收：</b>{html.escape(final)}</div>'
    )


def text_difference(before: str, after: str) -> tuple[str, str]:
    before_parts, after_parts = [], []
    for tag, first_start, first_end, second_start, second_end in difflib.SequenceMatcher(None, before, after).get_opcodes():
        left, right = html.escape(before[first_start:first_end]), html.escape(after[second_start:second_end])
        before_parts.append(f"<mark>{left}</mark>" if tag in {"delete", "replace"} else left)
        after_parts.append(f"<mark>{right}</mark>" if tag in {"insert", "replace"} else right)
    return "".join(before_parts), "".join(after_parts)


def changes_html(changes: list[dict], heading: str) -> str:
    if not changes:
        return ""
    rows = []
    for change in changes:
        before, after = text_difference(str(change.get("before", "")), str(change.get("after", "")))
        rows.append(f'<div class="txt-change"><b>修改前</b><p>{before}</p><b>修改后</b><p>{after}</p></div>')
    return f'<section class="txt-changes"><h4>{html.escape(heading)}</h4>' + "".join(rows) + "</section>"


def build_image_review(work: Path, input_target: Path, batch_root: Path, python: Path, args) -> tuple[Path, dict]:
    preview = load_json(work / "preview" / "preview.json", {}) or {}
    cards = []
    projects = []
    for index, row in enumerate(preview.get("packages", []), 1):
        folder = str(row.get("folder", ""))
        visual = next((item.get("interface_sheet") for item in preview.get("visuals", {}).get("pre_review_projects", []) if item.get("folder") == folder), None)
        login = row.get("manual_docx", {}).get("login_screen_review", {})
        login_status = login.get("status", "missing")
        login_messages = {
            "missing": "警报：未识别到登录界面，请确认说明书是否漏图。",
            "multiple": "警报：按顺序识别到两个登录界面，请退回检查。",
            "single": "登录界面检查：识别到 1 个。",
            "unavailable": "警报：当前 OCR 组件不可用，登录界面尚未完成检查。",
        }
        login_html = f'<p class="login-{"bad" if login_status in {"missing", "multiple", "unavailable"} else "ok"}">{html.escape(login_messages.get(login_status, "登录界面未检查"))}</p>'
        style_items = row.get("manual_docx", {}).get("interface_style_candidates", [])
        style_html = (f'<p class="login-bad">有{len(style_items)}个图片与项目主流截图的配色和大体均明显不一致</p>' if style_items else '<p class="login-ok">图片主配色和大体未发现明显不一致</p>')
        manual_name = str(row.get("files", {}).get("manual_pdf", {}).get("name", ""))
        manual_docx_name = str(row.get("files", {}).get("docx", {}).get("name", ""))
        source_base = Path(str(preview.get("base", batch_root)))
        manual_path = source_base / folder / manual_name if manual_name else None
        manual_docx_path = source_base / folder / manual_docx_name if manual_docx_name else None
        manual_link = file_uri(str(manual_path)) if manual_path else None
        manual_docx_link = file_uri(str(manual_docx_path)) if manual_docx_path else None
        source_links = "<p>" + (
            f'<a class="pdf-link" href="{html.escape(manual_link, quote=True)}" target="_blank">说明 PDF</a>'
            if manual_link else '<span class="missing">未找到说明 PDF</span>'
        ) + " " + (
            f'<a class="word-link" href="{html.escape(manual_docx_link, quote=True)}" target="_blank">打开说明 Word</a>'
            if manual_docx_link else '<span class="missing">未找到说明 Word</span>'
        ) + "</p>"
        projects.append({"index": index, "name": str(row.get("name", "")), "folder": folder, "manual_pdf": str(manual_path or ""), "manual_docx": str(manual_docx_path or "")})
        cards.append(f'<article class="project"><label><input class="rejection" type="checkbox" value="{index}"> '
            f'{source_links}'
            f'<b>{index}. {html.escape(str(row.get("name", "")))}</b><span class="reject-label">勾选退回</span></label>{style_html}{login_html}{image_panel("说明页中的界面和登录图片", visual)}</article>')
    review = work / "image_review.html"
    command = [str(python), str(Path(__file__).resolve()), str(input_target), "--work-dir", str(work)]
    if getattr(args, "from_extracted", False):
        command.append("--from-extracted")
    command.extend(["--reject-images", "未通过序号"])
    actions = '''<section class="review-actions"><h2>审核完成后继续</h2><p>确认已勾选所有需要退回的项目后，再复制序号或返回程序。</p><button onclick="copy()">复制未通过序号</button><button onclick="returnToApp()">返回程序继续处理</button><code id="out">未勾选任何项目：全部继续处理</code></section>'''
    document = f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>首轮图片审核</title>
<style>body{{font:15px/1.5 "Microsoft YaHei",sans-serif;margin:0;background:#f3f5f7;color:#182230}}main{{max-width:1200px;margin:auto;padding:24px}}.flow{{display:flex;gap:8px;flex-wrap:wrap;margin:14px 0 20px}}.flow span{{padding:6px 10px;border:1px solid #c7d7ff;background:#fff;color:#516078;font-size:13px}}.flow .current{{background:#165dff;border-color:#165dff;color:#fff;font-weight:700}}.project{{background:#fff;border:1px solid #dce2e8;margin:16px 0;padding:18px}}.visual img{{display:block;max-width:100%;max-height:800px;margin-top:10px;border:1px solid #dce2e8}}.visual a{{color:#165dff}}.word-link{{margin-left:10px;color:#087443;font-weight:700}}button{{padding:9px 14px;background:#165dff;color:#fff;border:0}}.reject-label{{margin-left:10px;color:#b42318;font-size:13px}}.login-bad{{padding:8px 10px;background:#fff0ef;color:#b42318;border-left:4px solid #d92d20}}.login-ok{{color:#087443}}.review-actions{{margin:24px 0 0;padding:20px;background:#fff;border:2px solid #165dff}}.review-actions h2{{margin:0 0 6px}}.review-actions button{{margin:8px 10px 8px 0}}.review-actions code{{display:block;padding:9px;background:#eef4ff;color:#163a70}}</style></head>
<body><main><h1>首轮图片审核</h1><div class="flow"><span>1 选择材料</span><span class="current">2 看图并微调 Word</span><span>3 自动二次处理</span><span>4 终审</span><span>5 导出 PDF</span><span>6 压缩交付</span></div><p>图片无明显问题时，可先点项目旁的“打开说明 Word”微调；完成后回到程序继续二次处理。只勾选图片不合格、需要退回的项目。</p>{''.join(cards)}{actions}</main>
<script>function copy(){{let v=[...document.querySelectorAll('.rejection:checked')].map(x=>x.value).join(',');out.textContent=v||'none';navigator.clipboard&&navigator.clipboard.writeText(v||'none')}}function returnToApp(){{copy();window.close()}}</script></body></html>'''
    review.write_text(document, encoding="utf-8")
    manifest = {"version": VERSION, "input_target": str(input_target.resolve()), "batch_root": str(batch_root), "review": str(review), "projects": projects, "next": subprocess.list2cmdline(command)}
    atomic_json(work / "image_review.json", manifest)
    return review, manifest


def build_review(work: Path, input_target: Path, batch_root: Path, modified: Path, python: Path, args) -> tuple[Path, dict]:
    preview = load_json(work / "preview" / "preview.json", {}) or {}
    audit = load_json(work / "audit" / "audit.json", {}) or {}
    summary = load_json(work / "summary.json", {}) or {}
    qa = load_json(work / "qa.json", {}) or {}
    txt_report = load_json(work / "txt_repair_report.json", {}) or {}
    word_report = load_json(work / "word_report.json", {}) or {}
    errors_by_folder: dict[str, list[dict]] = {}
    candidates_by_folder: dict[str, list[dict]] = {}
    folders_by_name: dict[str, list[str]] = {}
    for row in audit.get("packages", []):
        folders_by_name.setdefault(str(row.get("name", "")), []).append(str(row.get("folder", "")))
    for item in summary.get("clear_errors", []):
        folder = str(item.get("folder", ""))
        if not folder:
            matches = folders_by_name.get(str(item.get("program", "")), [])
            folder = matches[0] if len(matches) == 1 else ""
        errors_by_folder.setdefault(folder, []).append(item)

    projects = []
    cards = []
    result_by_folder = {str(item.get("folder", "")): item for item in summary.get("project_results", [])}
    txt_by_folder = {str(item.get("folder", "")): item for item in txt_report.get("packages", [])}
    manual_by_folder = {str(item.get("folder", "")): item for item in word_report.get("packages", [])}
    summary_rows = []
    for index, row in enumerate(audit.get("packages", []), 1):
        name = str(row.get("name", f"项目{index}"))
        folder = str(row.get("folder", ""))
        errors = errors_by_folder.get(folder, [])
        candidates = []
        projects.append({"index": index, "name": name, "folder": folder})
        result = result_by_folder.get(folder, {})
        txt_changes = txt_by_folder.get(folder, {}).get("main_function_changes", [])
        manual_changes = manual_by_folder.get(folder, {}).get("text_changes", [])
        # The first audit intentionally flags missing extracted PDF page numbers.
        # Once Word has re-exported a passing code PDF, that historical warning is
        # stale and would make the final review contradict the repair result.
        if result.get("code_final") == "最终检查通过":
            candidates = [
                item for item in candidates
                if str(item.get("message", "")) != "未同时提取到首尾页码，需结合拼图确认"
            ]
        repairs = result.get("txt_repairs", {})
        repair_labels = {
            "parentheses_to_commas": "括号改为逗号",
            "quotes_removed": "删除双引号",
            "duplicate_punctuation_removed": "删除重复标点",
            "ordinal_markers_removed": "删除序号或序数词",
        }
        repair_text = "；".join(f"{repair_labels.get(key, key)} {value} 处" for key, value in repairs.items()) or "无需修改"
        has_txt = bool(row.get("files", {}).get("txt"))
        txt_original = "无问题" if result.get("txt_original_valid", has_txt) else ("主要功能文字已发现问题" if repairs else "未提供 TXT")
        txt_final = str(result.get("txt_final", "本次未处理"))
        summary_rows.append(
            f"<tr><td>{index}</td><td>{html.escape(name)}</td>"
            f"<td>{result_cell(txt_original, repair_text, txt_final)}</td>"
            f"<td>{result_cell(str(result.get('code_original', '未提供代码 PDF')), str(result.get('code_actions', '无需修改')), str(result.get('code_final', '未生成最终检查结果')))}</td>"
            f"<td>{result_cell(str(result.get('manual_original', '未提供说明书')), str(result.get('manual_actions', '无需修改')), str(result.get('manual_final', '未生成最终检查结果')))}</td></tr>"
        )
        cards.append(
            f'<article class="project" id="project-{index}">'
            f'<header><label><input class="rejection" type="checkbox" value="{index}"> '
            f'<span class="index">{index}</span> {html.escape(name)}</label>'
            f'<span class="folder">{html.escape(folder)}</span></header>'
            f'{findings_html(errors, "需要处理")}'
            f'{changes_html(txt_changes, "主要功能模块修改对比")}'
            f'{changes_html(manual_changes, "说明书正文修改对比")}'
            f'</article>'
        )

    review_path = work / "unified_review.html"
    resume = [str(python), str(Path(__file__).resolve()), str(input_target)]
    for name, flag in (("extracted", "--extracted"), ("modified", "--modified"), ("work_dir", "--work-dir"), ("python", "--python")):
        value = getattr(args, name, None)
        if value:
            resume.extend([flag, str(value)])
    if getattr(args, "from_extracted", False):
        resume.append("--from-extracted")
    if args.workers:
        resume.extend(["--workers", str(args.workers)])
    if getattr(args, "manual_only", False):
        resume.append("--manual-only")
    resume.extend(["--approve", "通过序号"])
    command_prefix = subprocess.list2cmdline(resume)
    code_source = audit.get("visuals", {}).get("code_edges")
    code_modified = qa.get("modified_code_edges_qa")
    manual_modified = qa.get("modified_manual_qa")
    document = f'''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>软著材料统一审核</title>
<style>
:root{{--bg:#f4f6f8;--card:#fff;--ink:#172033;--muted:#657085;--line:#dfe4ea;--accent:#165dff;--bad:#b42318;--ok:#087443}}
*{{box-sizing:border-box}} body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 "Microsoft YaHei",sans-serif}}
main{{max-width:1500px;margin:auto;padding:28px}} h1{{margin:0 0 6px}} .lead{{color:var(--muted);margin:0 0 22px}}
.toolbar{{position:sticky;top:0;z-index:5;background:#ffffffee;border:1px solid var(--line);border-radius:12px;padding:12px 16px;margin-bottom:18px;backdrop-filter:blur(8px)}}
button{{border:0;border-radius:8px;padding:9px 14px;margin-right:8px;cursor:pointer}} .primary{{background:var(--accent);color:white}} code{{display:block;margin-top:10px;padding:10px;background:#eef2f7;border-radius:8px;overflow:auto}}
.summary-table{{width:100%;border-collapse:collapse;font-size:14px}} .summary-table th,.summary-table td{{border:1px solid var(--line);padding:10px;text-align:left;vertical-align:top}} .summary-table th{{background:#eef2f7}} .result-cell{{min-width:250px}} .result-cell b{{color:#34445e}} .status{{display:inline-block;padding:2px 10px;font-weight:700;border-radius:4px;margin-bottom:6px}} .status.pass{{color:#087443;background:#e7f6ee}} .status.fail{{color:#b42318;background:#ffebe9}} .txt-changes{{margin:14px 0 0;border-top:1px solid var(--line)}} .txt-changes h4{{margin:12px 0 6px}} .txt-change{{display:grid;grid-template-columns:90px 1fr;column-gap:12px;border-top:1px solid #edf0f3;padding:8px 0}} .txt-change p{{margin:0}} mark{{background:#ffe08a;color:#7a4300;padding:0 2px}}
.flow{{display:flex;gap:8px;flex-wrap:wrap;margin:14px 0 20px}}.flow span{{padding:6px 10px;border:1px solid #c7d7ff;background:#fff;color:var(--muted);font-size:13px;border-radius:6px}}.flow .current{{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:700}}.project,.global{{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px;margin:16px 0;box-shadow:0 2px 10px #1720330d}}
.project header{{display:flex;justify-content:space-between;gap:16px;align-items:center;font-size:20px;font-weight:700}} .index{{display:inline-grid;place-items:center;width:32px;height:32px;border-radius:50%;background:var(--accent);color:#fff}}
.folder{{font-size:12px;color:var(--muted);font-weight:400}} .findings h4{{margin-bottom:4px}} .findings li{{color:var(--bad)}} .ok{{color:var(--ok)}}
.visual-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:14px}} .visual{{border-top:1px solid var(--line);padding-top:10px}} .visual img{{display:block;max-width:100%;max-height:720px;margin-top:8px;border:1px solid var(--line);border-radius:8px;background:#eee}} .missing{{color:var(--muted)}}
@media(max-width:700px){{main{{padding:12px}}.visual-grid{{grid-template-columns:1fr}}.project header{{align-items:flex-start;flex-direction:column}}}}
</style></head><body><main>
<h1>项目整改结果</h1><div class="flow"><span>1 选择材料</span><span>2 首轮图审与 Word 微调</span><span>3 自动二次处理</span><span class="current">4 终审与人工微调</span><span>5 导出 PDF</span><span>6 压缩交付</span></div><p class="lead">终审发现问题时，先修改修改版 Word，再回程序选择“重新验收修改版”或“直接导出 PDF”。</p>
<div class="toolbar"><button onclick="setAll(true)">全部标为不通过</button><button onclick="setAll(false)">全部通过</button><button class="primary" onclick="copyCommand()">复制打包命令</button><code id="command">未勾选任何项目：全部打包</code></div>
<section class="global"><h2>项目总成果</h2><p class="lead">每一格只回答三件事：发现了什么、修了什么、现在是否合格。</p><table class="summary-table"><thead><tr><th>序号</th><th>名称</th><th>TXT</th><th>代码 PDF</th><th>说明书</th></tr></thead><tbody>{''.join(summary_rows)}</tbody></table></section>
<section class="global"><h2>修复后说明书封面前两页</h2><p class="lead">用于确认封面、软件名称、版本和页眉页码的修复成果；不重复展示首轮已经审核过的界面图片。</p>{image_panel("查看说明书封面前两页组图", manual_modified)}</section>
{("" if getattr(args, "manual_only", False) else f'<section class="global"><h2>最终代码 PDF 首尾页</h2><p class="lead">每个继续处理的项目都有一组首尾页。没有修复的项目显示原文件；已修复的项目显示修复后的最终文件。</p>{image_panel("查看代码首尾页组图", code_modified or code_source)}</section>')}
{''.join(cards)}
</main><script>
const prefix={json.dumps(command_prefix, ensure_ascii=False)};
function rejected(){{return [...document.querySelectorAll('.rejection:checked')].map(x=>x.value)}}
function approved(){{const bad=new Set(rejected());return [...document.querySelectorAll('.rejection')].map(x=>x.value).filter(x=>!bad.has(x))}}
function refresh(){{const v=approved();document.getElementById('command').textContent=v.length?prefix.replace('通过序号',v.join(',')):'没有可打包项目'}}
function setAll(v){{document.querySelectorAll('.rejection').forEach(x=>x.checked=v);refresh()}}
async function copyCommand(){{refresh();const text=document.getElementById('command').textContent;try{{await navigator.clipboard.writeText(text)}}catch(error){{const area=document.createElement('textarea');area.value=text;document.body.appendChild(area);area.select();document.execCommand('copy');area.remove()}}}}
document.querySelectorAll('.rejection').forEach(x=>x.addEventListener('change',refresh));
</script></body></html>'''
    review_path.write_text(document, encoding="utf-8")
    manifest = {
        "version": VERSION,
        "input_target": str(input_target),
        "batch_root": str(batch_root),
        "modified": str(modified),
        "review": str(review_path),
        "projects": projects,
        "summary": str(work / "summary.json"),
        "next": command_prefix,
        "review_fingerprint": hashlib.sha256(document.encode("utf-8")).hexdigest(),
    }
    atomic_json(work / "unified_review.json", manifest)
    return review_path, manifest


def process_all(args, python: Path, input_target: Path, batch_root: Path, extracted: Path, modified: Path, work: Path, selected_archive: Path | None) -> dict:
    common = [str(python), str(PROCESS), str(batch_root)]
    common.extend(["--extracted", str(extracted), "--modified", str(modified), "--work-dir", str(work)])
    if getattr(args, "from_extracted", False):
        common.append("--use-existing-extracted")
    if selected_archive is not None:
        common.extend(["--input-archive", str(selected_archive)])
    if args.workers:
        common.extend(["--workers", str(args.workers)])
    if args.force:
        common.append("--force")
    if args.manual_only:
        common.append("--manual-only")

    preview = run_json([*common, "--stage", "preview"])
    if preview.get("status") != "NEEDS_PROJECT_SELECTION":
        return preview
    preview_json = load_json(work / "preview" / "preview.json", {}) or {}
    packages = preview_json.get("packages", [])
    if not packages:
        return {"status": "FAILED", "error": "未发现可处理项目", "work_dir": str(work)}
    if args.approve_images is None and args.reject_images is None:
        review, manifest = build_image_review(work, input_target, batch_root, python, args)
        return {
            "version": VERSION,
            "status": "NEEDS_IMAGE_REVIEW",
            "projects": manifest["projects"],
            "review": str(review),
            "review_link": markdown_link(review, "打开首轮图片审核"),
            "question": "只勾选图片不通过、需要退回的项目；未勾选项目会继续处理。",
            "next": manifest["next"],
            "work_dir": str(work),
        }
    if args.reject_images is not None:
        rejected = [] if args.reject_images.strip().lower() in {"", "none", "无"} else parse_indexes(args.reject_images, len(packages))
        indexes = [index for index in range(1, len(packages) + 1) if index not in rejected]
    else:
        indexes = parse_indexes(args.approve_images, len(packages))
    if not indexes:
        return {"version": VERSION, "status": "FAILED", "error": "所有项目均被标记为图片不通过，没有可继续处理的项目", "work_dir": str(work)}
    repairs = [] if args.manual_only else code_header_repairs(preview_json)
    atomic_json(work / "code_reviewed.json", {
        "version": VERSION,
        "preview_fingerprint": preview_json.get("fingerprint"),
        "code_header_repairs": repairs,
        "decision": "programmatic",
    })
    command = [*common, "--stage", "all"]
    for index in indexes:
        command.extend(["--include-index", str(index)])
    result = run_json(command)
    if result.get("status") not in {"OK", "HAS_ERRORS"}:
        return result
    review, manifest = build_review(work, input_target, batch_root, modified, python, args)
    return {
        "version": VERSION,
        "status": "NEEDS_UNIFIED_REVIEW",
        "projects": manifest["projects"],
        "review": str(review),
        "review_link": markdown_link(review, "打开统一审核页"),
        "summary": manifest["summary"],
        "question": "请在统一审核页勾选通过项目，然后回复通过序号。",
        "next": manifest["next"],
        "clear_errors": result.get("clear_errors", 0),
        "visual_candidates": result.get("visual_candidates", 0),
        "work_dir": str(work),
    }


def recheck_modified(args, python: Path, input_target: Path, batch_root: Path, extracted: Path, modified: Path, work: Path) -> dict:
    """Rebuild the final review from user-adjusted modified materials only."""
    command = [
        str(python), str(PROCESS), str(batch_root),
        "--extracted", str(extracted), "--modified", str(modified),
        "--work-dir", str(work), "--stage", "verify",
    ]
    if args.manual_only:
        command.append("--manual-only")
    result = run_json(command)
    if result.get("status") not in {"OK", "HAS_ERRORS"}:
        return result
    review, manifest = build_review(work, input_target, batch_root, modified, python, args)
    return {
        "version": VERSION,
        "status": "NEEDS_UNIFIED_REVIEW",
        "projects": manifest["projects"],
        "review": str(review),
        "review_link": markdown_link(review, "打开重新验收页"),
        "summary": manifest["summary"],
        "question": "重新验收已完成；确认结果后可直接导出 PDF 或压缩。",
        "clear_errors": result.get("clear_errors", 0),
        "visual_candidates": result.get("visual_candidates", 0),
        "work_dir": str(work),
    }


def package_approved(args, python: Path, input_target: Path, batch_root: Path, modified: Path, work: Path, selected_archive: Path | None) -> dict:
    manifest = load_json(work / "unified_review.json", {}) or {}
    projects = manifest.get("projects", [])
    if not projects:
        return {"status": "FAILED", "error": "统一审核清单不存在，请先运行完整处理", "work_dir": str(work)}
    expected_input = str(input_target.resolve())
    expected_modified = str(modified.resolve())
    if manifest.get("input_target", manifest.get("batch_root")) != expected_input:
        return {"status": "FAILED", "error": "统一审核清单与当前输入不匹配，请重新运行处理", "work_dir": str(work)}
    if str(manifest.get("modified", "")) != expected_modified:
        return {"status": "FAILED", "error": "统一审核清单与当前修改版目录不匹配，请重新运行处理", "work_dir": str(work)}
    review_path = Path(str(manifest.get("review", "")))
    if not review_path.is_file():
        return {"status": "FAILED", "error": "统一审核页不存在或已失效，请重新运行处理", "work_dir": str(work)}
    indexes = parse_indexes(args.approve, len(projects))
    if args.format == "auto":
        archive_format = (archive_types.delivery_format(selected_archive) or "zip") if selected_archive is not None else detect_format(batch_root)
    else:
        archive_format = args.format
    if archive_format is None:
        return {"status": "NEEDS_ARCHIVE_FORMAT", "question": "输入同时包含 ZIP 和 RAR，请指定 --format zip 或 --format rar。"}
    command = [
        str(python), str(PACKAGE), str(modified), "--name", (selected_archive.stem if selected_archive is not None else batch_root.name),
        "--format", archive_format, "--layout-approved",
    ]
    for index in indexes:
        command.extend(["--include-folder", projects[index - 1]["folder"]])
    result = run_json(command)
    result["version"] = VERSION
    result["approved_indexes"] = indexes
    result["rejected_indexes"] = [i for i in range(1, len(projects) + 1) if i not in indexes]
    return result


def exportable_docx_files(directory: Path) -> list[Path]:
    """目录下可导出的 Word 文档，跳过临时文件和备份副本。"""
    if not directory.is_dir():
        return []
    return [
        path for path in directory.rglob("*.docx")
        if not path.name.startswith("~$")
        and not any(part.casefold().startswith(("backup", ".easysoftware-")) for part in path.parts)
    ]


def resolve_export_directory(modified: Path, input_target: Path) -> Path:
    """导出 PDF 的目标目录。

    不要求目录名以“修改版”“解压版”结尾：优先用户实际选择的目录
    （其“修改版”子目录更精确时用它），推导出的修改版位置只作兜底，
    避免为了找修改版而导出到用户没有选择的兄弟目录。
    """
    if input_target.is_dir():
        nested = input_target / "修改版"
        if exportable_docx_files(nested):
            return nested
        if exportable_docx_files(input_target):
            return input_target
    if exportable_docx_files(modified):
        return modified
    return input_target


def export_modified_pdfs(modified: Path, work: Path, progress: ProgressReporter | None = None) -> dict:
    """Export only modified DOCX files and replace their sibling PDFs."""
    def failed_result(error: str) -> dict:
        if progress:
            terminal_total = max(1, progress.total, progress.completed)
            progress.set_total(terminal_total)
            progress.update(
                "Word导出PDF失败", terminal_total, "导出终止",
                status="FAILED", error=error, active=False,
                failures=max(1, progress.failures + 1),
            )
        return {"version": VERSION, "status": "FAILED", "error": error, "work_dir": str(work)}

    if not modified.is_dir():
        return failed_result(f"未找到修改版目录，也没有可直接导出的 Word 文档：{modified}")
    docx_files = exportable_docx_files(modified)
    if not docx_files:
        return failed_result(f"目录内没有 Word 文档，无法导出 PDF：{modified}")
    if progress:
        progress.set_total(max(1, len(docx_files)))
    shell = shutil.which("powershell.exe") or shutil.which("pwsh") or shutil.which("powershell")
    if not shell:
        return failed_result("未找到 PowerShell")
    work.mkdir(parents=True, exist_ok=True)
    report = work / "modified_pdf_export_report.json"
    command = [shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(EXPORT_PDFS), "-BaseDir", str(modified), "-ReportPath", str(report)]
    if progress:
        command.extend(["-ProgressPath", str(progress.path)])
        progress.update("等待Word导出队列", 0, "Word导出PDF", active=False)
    lock_path = work / ".modified-pdf-export.running"
    lock_acquired = False
    try:
        for attempt in range(2):
            try:
                descriptor = os.open(lock_path, os.O_WRONLY | os.O_CREAT | os.O_EXCL)
                with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
                    json.dump({"pid": os.getpid(), "started_at": time.time()}, handle)
                lock_acquired = True
                break
            except FileExistsError:
                previous = load_json(lock_path, {}) or {}
                if attempt == 0 and not process_is_running(previous.get("pid")):
                    lock_path.unlink(missing_ok=True)
                    continue
                return failed_result("该修改版正在导出中，请等待当前任务结束，不要重复点击导出。")
        with word_automation_lock(progress, "Word导出PDF", timeout_seconds=180):
            completed = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=3600, check=False)
        result = load_json(report, None)
        if result is None:
            return failed_result(decode_process_output(completed.stdout)[-2000:] or "导出未返回报告")
        result["work_dir"] = str(work)
        if progress:
            processed = sum(int(result.get(key, 0) or 0) for key in ("exported", "skipped", "failed"))
            progress.update(
                "Word导出PDF", processed, "导出完成",
                "DONE" if result.get("status") == "EXPORT_OK" and not result.get("failed") else "FAILED",
                failures=int(result.get("failed", 0) or 0),
            )
        return result
    except subprocess.TimeoutExpired:
        return failed_result("PDF 导出超时，请检查 Word 是否弹出了需要人工处理的窗口")
    except Exception as exc:
        return failed_result(f"PDF 导出失败：{exc}")
    finally:
        if lock_acquired:
            lock_path.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="软著材料全自动处理与单页集中审核程序")
    parser.add_argument("batch_root", type=Path)
    parser.add_argument("--approve", help="统一审核通过序号，例如 1,3-5 或 all")
    parser.add_argument("--approve-images", help="兼容旧流程：首轮图片审核通过序号")
    parser.add_argument("--reject-images", help="首轮图片审核不通过序号；none 表示全部继续处理")
    parser.add_argument("--format", choices=("auto", "rar", "zip"), default="auto")
    parser.add_argument("--extracted", type=Path)
    parser.add_argument("--modified", type=Path)
    parser.add_argument("--work-dir", type=Path)
    parser.add_argument("--python")
    parser.add_argument("--workers", type=int, default=max(1, min(8, os.cpu_count() or 4)))
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--manual-only", action="store_true", help="只修复和验证说明书，跳过代码材料")
    parser.add_argument("--export-modified-pdfs", action="store_true", help="只将修改版DOCX导出并覆盖同名PDF")
    parser.add_argument("--recheck-modified", action="store_true", help="只重新验收人工调整后的修改版材料")
    parser.add_argument("--from-extracted", action="store_true", help="直接从已解压材料目录开始首轮审核")
    parser.add_argument("--clean-code-comments", action="store_true", help="独立预览代码DOCX中含汉字的真实注释，不进入其他流程")
    parser.add_argument("--apply-comment-cleanup", action="store_true", help="确认后执行代码注释清理并重新导出代码PDF")
    parser.add_argument("--skip-comment-preview", action="store_true", help="跳过代码注释清理预览（规则稳定后可用）")
    parser.add_argument("--process-documents", action="store_true", help="预览说明书、代码、TXT占位统一处理流程")
    parser.add_argument("--apply-document-processing", action="store_true", help="确认后按说明书、代码、TXT占位顺序执行统一处理")
    parser.add_argument("--document-stages", default="manual,code,txt_placeholder",
                        help="选择文档处理阶段，逗号分隔：manual,code,txt_placeholder；至少包含 manual 或 code")
    args = parser.parse_args()

    input_target = args.batch_root.resolve()
    if not input_target.exists() and archive_types.is_archive_file(input_target):
        nested = input_target.parent / input_target.stem / input_target.name
        if nested.is_file():
            input_target = nested.resolve()
    if args.clean_code_comments or args.apply_comment_cleanup or args.skip_comment_preview:
        work = (args.work_dir or default_work_dir(input_target) / "code-comment-cleanup").resolve()
        progress = ProgressReporter(work)
        try:
            from code_comment_cleaner import run as run_code_comment_cleaner
            apply_cleanup = bool(args.apply_comment_cleanup or args.skip_comment_preview)
            result = run_code_comment_cleaner(
                input_target, work, HERE.parent, apply=apply_cleanup,
                skip_preview=args.skip_comment_preview, progress=progress,
            )
            return emit(result, 0 if result.get("status") == "CODE_COMMENT_CLEANUP_OK" else 2)
        except Exception as exc:
            progress.set_total(max(1, progress.total, progress.completed))
            progress.update(
                "代码注释清理失败", progress.total, "处理终止",
                status="FAILED", error=str(exc), active=False, failures=progress.failures + 1,
            )
            return emit({"version": VERSION, "status": "FAILED", "error": str(exc), "work_dir": str(work)}, 2)
    if args.process_documents or args.apply_document_processing:
        work = (args.work_dir or default_work_dir(input_target) / "document-processing").resolve()
        progress = ProgressReporter(work)
        try:
            from combined_document_workflow import run as run_combined_document_workflow
            result = run_combined_document_workflow(
                input_target, work, HERE.parent,
                apply=bool(args.apply_document_processing), progress=progress,
                document_stages=args.document_stages,
            )
            return emit(result, 0 if result.get("status") == "DOCUMENT_PROCESSING_OK" else 2)
        except Exception as exc:
            progress.set_total(max(1, progress.total, progress.completed))
            progress.update(
                "材料统一处理失败", progress.total, "处理终止",
                status="FAILED", error=str(exc), active=False, failures=progress.failures + 1,
            )
            return emit({"version": VERSION, "status": "FAILED", "error": str(exc), "work_dir": str(work)}, 2)
    selected_archive = input_target if input_target.is_file() and archive_types.is_archive_file(input_target) else None
    if selected_archive is not None:
        selected_archive = prepare_archive_workspace(selected_archive)
        input_target = selected_archive
    if args.from_extracted:
        if not input_target.is_dir() or not any(input_target.iterdir()):
            return emit({"version": VERSION, "status": "FAILED", "error": f"已解压材料目录不存在或为空：{input_target}"}, 2)
        batch_root = input_target.parent
        extracted = input_target
        modified = (args.modified or resolve_modified_directory(batch_root, input_target)).resolve()
    else:
        batch_root = selected_archive.parent if selected_archive is not None else input_target
        if selected_archive is not None:
            extracted = (args.extracted or batch_root / f"{selected_archive.stem}_解压版").resolve()
            modified = (args.modified or resolve_modified_directory(batch_root, batch_root / f"{selected_archive.stem}_解压版")).resolve()
        else:
            extracted = (args.extracted or batch_root / "解压版").resolve()
            modified = (args.modified or resolve_modified_directory(batch_root, input_target if input_target.is_dir() else None)).resolve()
    work = (args.work_dir or default_work_dir(input_target)).resolve()
    python = find_python(args.python)
    if python is None:
        return emit({"version": VERSION, "status": "MISSING_DEPENDENCIES", "required": ["Python 3.10+", "lxml", "Pillow", "pypdf"]}, 2)
    try:
        if args.export_modified_pdfs:
            progress = ProgressReporter(work)
            result = export_modified_pdfs(resolve_export_directory(modified, input_target), work, progress)
        elif args.recheck_modified:
            result = recheck_modified(args, python, input_target, batch_root, extracted, modified, work)
        else:
            result = package_approved(args, python, input_target, batch_root, modified, work, selected_archive) if args.approve else process_all(args, python, input_target, batch_root, extracted, modified, work, selected_archive)
        code = 0 if result.get("status") in {"OK", "EXPORT_OK"} else 2
        return emit(result, code)
    except Exception as exc:
        return emit({"version": VERSION, "status": "FAILED", "error": str(exc), "work_dir": str(work)}, 2)


if __name__ == "__main__":
    raise SystemExit(main())

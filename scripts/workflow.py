from __future__ import annotations

if __name__ == "__main__":
    import runpy
    from pathlib import Path as _Path
    runpy.run_path(str(_Path(__file__).with_name("unified_workflow.py")), run_name="__main__")
    raise SystemExit(0)

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path


VERSION = "2.8.4"
HERE = Path(__file__).resolve().parent
PROCESS = HERE / "process_batch.py"
PACKAGE = HERE / "package_approved_modified.py"


def emit(value: dict, code: int = 0) -> int:
    print(json.dumps(value, ensure_ascii=False, separators=(",", ":")))
    return code


def load_json(path: Path, default=None):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default


def atomic_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


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
    streams = [result.stdout, result.stderr]
    for stream in streams:
        for line in reversed((stream or "").splitlines()):
            try:
                value = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(value, dict):
                return value
    return {"status": "FAILED", "error": (result.stderr or result.stdout or "subprocess failed")[-3000:]}


def run_json(command: list[str]) -> dict:
    result = subprocess.run(command, text=True, encoding="utf-8", errors="replace", stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    value = parse_result(result)
    value.setdefault("returncode", result.returncode)
    return value


def parse_indexes(value: str, maximum: int) -> list[int]:
    value = value.strip().lower()
    if value in {"all", "全部", "全选"}:
        return list(range(1, maximum + 1))
    indexes = set()
    for token in re.split(r"[,，、;；\s]+", value):
        if not token:
            continue
        match = re.fullmatch(r"(\d+)[-~—至](\d+)", token)
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


def common_process_args(args, python: Path, stage: str) -> list[str]:
    command = [str(python), str(PROCESS), str(args.batch_root), "--stage", stage]
    if args.extracted:
        command.extend(["--extracted", str(args.extracted)])
    if args.modified:
        command.extend(["--modified", str(args.modified)])
    if args.work_dir:
        command.extend(["--work-dir", str(args.work_dir)])
    if args.workers:
        command.extend(["--workers", str(args.workers)])
    if args.force:
        command.append("--force")
    if args.no_visuals:
        command.append("--no-visuals")
    return command


def resume_command(args, *extra: str) -> str:
    command = [sys.executable, str(Path(__file__).resolve()), str(args.batch_root)]
    for name, flag in (("extracted", "--extracted"), ("modified", "--modified"), ("work_dir", "--work-dir"), ("python", "--python")):
        value = getattr(args, name, None)
        if value:
            command.extend([flag, str(value)])
    if args.workers:
        command.extend(["--workers", str(args.workers)])
    command.extend(extra)
    return subprocess.list2cmdline(command)


def markdown_file_links(paths) -> list[str]:
    """Return clickable Markdown links for generated local visual artifacts."""
    links = []
    for value in paths or []:
        if not value:
            continue
        path = Path(value).resolve()
        links.append(f"[{path.name}](<{path}>)")
    return links


def compact_preview(raw: dict, work: Path, args) -> dict:
    preview = load_json(work / "preview" / "preview.json", {})
    projects = [{"index": row["index"], "name": row["name"]} for row in preview.get("packages", [])]
    visuals = raw.get("visuals", {})
    display = visuals.get("interface_overviews") or ([visuals.get("interface_overview")] if visuals.get("interface_overview") else [])
    return {
        "version": VERSION,
        "status": "NEEDS_PROJECT_SELECTION",
        "batch": str(args.batch_root),
        "display": display,
        "display_links": markdown_file_links(display),
        "projects": projects,
        "question": "哪些序号可以继续，哪些序号不合格？",
        "next": resume_command(args, "--approve", "序号"),
        "timing": raw.get("timing", {}),
    }


def compact_code_preview(raw: dict, work: Path, args) -> dict:
    preview = load_json(work / "preview" / "preview.json", {})
    findings = []
    for row in preview.get("packages", []):
        for kind in ("clear_errors", "candidates"):
            for item in row.get(kind, []):
                findings.append({"program": row["name"], "severity": kind, **item})
    board = preview.get("visuals", {}).get("code_edges") or raw.get("visuals", {}).get("code_edges")
    header_candidates = sorted({
        item["program"] for item in findings
        if "页眉" in item.get("location", "") or "页码" in item.get("message", "")
    })
    return {
        "version": VERSION,
        "status": "NEEDS_AGENT_CODE_REVIEW",
        "batch": str(args.batch_root),
        "agent_review": [board] if board else [],
        "agent_review_links": markdown_file_links([board] if board else []),
        "findings": findings,
        "code_header_candidates": header_candidates,
        "instruction": "先检查代码PDF第一页和最后一页的格式、符号、括号、缩进、页眉及总页数；先向用户汇报明确错误及位置。确认页眉/页码异常的项目，在继续命令后追加重复的 --code-header-fix 软件名；之后才显示说明程序界面组图。",
        "next": resume_command(args, "--code-reviewed"),
        "timing": raw.get("timing", {}),
    }


def code_review_is_current(work: Path, preview: dict) -> bool:
    marker = load_json(work / "code_reviewed.json", {})
    return marker.get("preview_fingerprint") == preview.get("fingerprint")


def detect_format(batch_root: Path) -> str | None:
    archives = sorted(
        (path for path in batch_root.iterdir() if path.is_file() and path.suffix.lower() in {".zip", ".rar"}),
        key=lambda path: (path.stem.casefold() != batch_root.name.casefold(), -path.stat().st_mtime_ns),
    )
    formats = {path.suffix.lower().lstrip(".") for path in archives}
    if archives and archives[0].stem.casefold() == batch_root.name.casefold():
        return archives[0].suffix.lower().lstrip(".")
    return next(iter(formats)) if len(formats) == 1 else None


def main() -> int:
    parser = argparse.ArgumentParser(description="软著材料跨代理单入口极速状态机")
    parser.add_argument("batch_root", type=Path)
    parser.add_argument("--approve", help="获准项目序号，例如 1,3-5 或 all")
    # v2.8 compatibility no-ops: interface/login review now belongs to the user's project-selection gate.
    parser.add_argument("--login-present", action="append", default=[], help=argparse.SUPPRESS)
    parser.add_argument("--login-absent", action="append", default=[], help=argparse.SUPPRESS)
    parser.add_argument("--login-all-present", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--login-all-absent", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--code-reviewed", action="store_true", help="Continue after reporting code-PDF edge findings")
    parser.add_argument("--code-header-fix", action="append", default=[], help="确认需要规范代码DOCX页眉并重新导出代码PDF的软件名称；可重复")
    parser.add_argument("--agent-reviewed", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--layout-approved", action="store_true")
    parser.add_argument("--format", choices=("auto", "zip", "rar"), default="auto")
    parser.add_argument("--audit-only", action="store_true")
    parser.add_argument("--python", help="Python executable with lxml, Pillow and pypdf")
    parser.add_argument("--extracted", type=Path)
    parser.add_argument("--modified", type=Path)
    parser.add_argument("--work-dir", type=Path)
    parser.add_argument("--workers", type=int)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--no-visuals", action="store_true")
    args = parser.parse_args()

    args.batch_root = args.batch_root.resolve()
    work = (args.work_dir or default_work_dir(args.batch_root)).resolve()
    modified = (args.modified or args.batch_root / "修改版").resolve()
    python = find_python(args.python)
    if python is None:
        return emit({
            "version": VERSION,
            "status": "MISSING_DEPENDENCIES",
            "required": ["Python 3.10+", "lxml", "Pillow", "pypdf", "pdftoppm", "Microsoft Word", "WinRAR for RAR delivery"],
            "hint": "Set SOFTCOPYRIGHT_PYTHON or pass --python to a compatible interpreter.",
        }, 2)

    try:
        if args.layout_approved:
            summary = load_json(work / "summary.json", {})
            overview = summary.get("visuals", {}).get("processed_manual_first_two_overview")
            if not overview or not Path(overview).is_file():
                return emit({"version": VERSION, "status": "NEEDS_LAYOUT_REVIEW_ARTIFACT", "error": "先完成修复和验证并生成说明PDF前两页总组图"}, 2)
            archive_format = detect_format(args.batch_root) if args.format == "auto" else args.format
            if archive_format is None:
                return emit({"version": VERSION, "status": "NEEDS_ARCHIVE_FORMAT", "question": "输入同时包含ZIP和RAR，请指定 --format zip 或 --format rar。"}, 2)
            raw = run_json([str(python), str(PACKAGE), str(modified), "--name", args.batch_root.name, "--format", archive_format, "--layout-approved"])
            raw["version"] = VERSION
            return emit(raw, 0 if raw.get("status") == "OK" else 2)

        if args.code_reviewed:
            raw = run_json(common_process_args(args, python, "preview"))
            preview = load_json(work / "preview" / "preview.json", {})
            known = {row.get("name") for row in preview.get("packages", [])}
            unknown = sorted(set(args.code_header_fix) - known)
            if unknown:
                raise ValueError("代码页眉修复项目不在预览清单：" + "、".join(unknown))
            atomic_json(work / "code_reviewed.json", {
                "version": VERSION,
                "preview_fingerprint": preview.get("fingerprint"),
                "code_header_repairs": sorted(set(args.code_header_fix)),
            })
            return emit(compact_preview(raw, work, args))

        if args.approve or args.login_present or args.login_absent or args.login_all_present or args.login_all_absent or args.agent_reviewed:
            preview = load_json(work / "preview" / "preview.json", {})
            if not preview:
                raw = run_json(common_process_args(args, python, "preview"))
                return emit(compact_code_preview(raw, work, args))
            if not code_review_is_current(work, preview):
                return emit(compact_code_preview({"timing": {}}, work, args))
            command = common_process_args(args, python, "audit")
            if args.approve:
                for index in parse_indexes(args.approve, preview.get("package_count", 0)):
                    command.extend(["--include-index", str(index)])
            raw = run_json(command)
            if args.audit_only:
                raw["version"] = VERSION
                return emit(raw, 0 if raw.get("status") in {"OK", "HAS_ERRORS"} else 2)
            if raw.get("status") not in {"OK", "HAS_ERRORS"}:
                raw["version"] = VERSION
                return emit(raw, 2)
            raw = run_json(common_process_args(args, python, "repair"))
            if raw.get("status") not in {"OK", "HAS_ERRORS"}:
                raw["version"] = VERSION
                return emit(raw, 2)
            verified = run_json(common_process_args(args, python, "verify"))
            if verified.get("status") not in {"OK", "HAS_ERRORS"}:
                verified["version"] = VERSION
                return emit(verified, 2)
            overview = verified.get("visuals", {}).get("processed_manual_first_two_overview")
            return emit({
                "version": VERSION,
                "status": "NEEDS_LAYOUT_APPROVAL",
                "display": [overview] if overview else [],
                "display_links": markdown_file_links([overview] if overview else []),
                "agent_review": [verified.get("visuals", {}).get("modified_code_edges_qa")] if verified.get("visuals", {}).get("modified_code_edges_qa") else [],
                "agent_review_links": markdown_file_links([verified.get("visuals", {}).get("modified_code_edges_qa")] if verified.get("visuals", {}).get("modified_code_edges_qa") else []),
                "question": "处理后说明PDF前两页的排版合适吗？",
                "next": resume_command(args, "--layout-approved"),
                "summary": verified.get("summary"),
                "clear_errors": verified.get("clear_errors", 0),
                "visual_candidates": verified.get("visual_candidates", 0),
                "timing": {"repair": raw.get("timing", {}), "verify": verified.get("timing", {})},
            })

        raw = run_json(common_process_args(args, python, "preview"))
        if raw.get("status") != "NEEDS_PROJECT_SELECTION":
            raw["version"] = VERSION
            return emit(raw, 2)
        return emit(compact_code_preview(raw, work, args))
    except (OSError, ValueError) as exc:
        return emit({"version": VERSION, "status": "FAILED", "error": str(exc)}, 2)


if __name__ == "__main__":
    # Keep the legacy filename callable without allowing its obsolete review
    # state machine to bypass the current two-stage workflow.
    raise SystemExit(subprocess.call([sys.executable, str(HERE / "unified_workflow.py"), *sys.argv[1:]]))

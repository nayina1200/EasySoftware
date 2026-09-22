from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


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


def restore_pair(modified: Path, package: dict, backup_dir: Path, had_pdf: bool) -> None:
    folder = modified / str(package.get("folder", ""))
    docx = folder / str(package.get("docx", ""))
    pdf_name = str(package.get("pdf") or Path(docx.name).with_suffix(".pdf").name)
    pdf = folder / pdf_name
    backup_docx = backup_dir / docx.name
    backup_pdf = backup_dir / pdf.name
    if backup_docx.is_file():
        shutil.copy2(backup_docx, docx)
    if had_pdf and backup_pdf.is_file():
        shutil.copy2(backup_pdf, pdf)
    elif not had_pdf and pdf.exists():
        pdf.unlink()
    lock = docx.with_name("~$" + docx.name[2:])
    if lock.exists():
        lock.unlink()


def list_word_pids() -> set[int]:
    if os.name != "nt":
        return set()
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    import re
    return {int(value) for value in re.findall(r'"WINWORD\.EXE","(\d+)"', result.stdout or "", re.I)}


def terminate_tree(process: subprocess.Popen[str], word_pid_path: Path, fallback_word_pids: set[int]) -> None:
    if os.name == "nt":
        word_pid = 0
        try:
            word_pid = int(word_pid_path.read_text(encoding="utf-8-sig").strip())
        except (OSError, ValueError):
            pass
        targets = {word_pid} if word_pid else fallback_word_pids
        for pid in targets:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    else:
        process.kill()
    try:
        process.communicate(timeout=15)
    except subprocess.TimeoutExpired:
        process.kill()
        process.communicate()


def main() -> int:
    parser = argparse.ArgumentParser(description="Run one Word repair item with timeout and rollback.")
    parser.add_argument("--base", required=True, type=Path)
    parser.add_argument("--script", required=True, type=Path)
    parser.add_argument("--package-json", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--timeout", type=int, default=180)
    args = parser.parse_args()

    modified = args.base.resolve()
    package = load_json(args.package_json, {}) or {}
    name = str(package.get("name", "未命名项目"))
    folder = modified / str(package.get("folder", ""))
    docx = folder / str(package.get("docx", ""))
    pdf_name = str(package.get("pdf") or Path(docx.name).with_suffix(".pdf").name)
    pdf = folder / pdf_name
    backup_dir = args.report.parent / "timeout_backups" / name
    backup_dir.mkdir(parents=True, exist_ok=True)
    had_pdf = pdf.is_file()
    shutil.copy2(docx, backup_dir / docx.name)
    if had_pdf:
        shutil.copy2(pdf, backup_dir / pdf.name)
    elif (backup_dir / pdf.name).exists():
        (backup_dir / pdf.name).unlink()

    shell = shutil.which("powershell.exe") or shutil.which("pwsh") or shutil.which("powershell")
    if not shell:
        raise RuntimeError("未找到PowerShell，无法运行Word批处理")
    with tempfile.TemporaryDirectory(prefix="easysoftware-word-item-") as temporary:
        plan_path = Path(temporary) / "plan.json"
        report_path = Path(temporary) / "report.json"
        atomic_json(plan_path, {"version": "3.0.1", "base": str(modified), "packages": [package]})
        word_pid_path = Path(temporary) / "word.pid"
        command = [
            shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(args.script),
            "-BaseDir", str(modified), "-PlanPath", str(plan_path), "-ReportPath", str(report_path),
            "-PidPath", str(word_pid_path),
        ]
        word_pids_before = list_word_pids()
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        try:
            output, _ = process.communicate(timeout=args.timeout)
            output = decode_process_output(output)
        except subprocess.TimeoutExpired:
            terminate_tree(process, word_pid_path, list_word_pids() - word_pids_before)
            result = {
                "name": name,
                "folder": package.get("folder", ""),
                "status": "PRESERVED_FOR_REVIEW",
                "changed": True,
                "actions_requested": package.get("actions", []),
                "actions_applied": [],
                "failures": [f"Word单文档处理超时（>{args.timeout}秒），已终止并保留当前成果供审核"],
                "word_pages": 0,
                "seconds": args.timeout,
            }
            atomic_json(args.report, {"version": "3.0.1-item", "packages": [result]})
            print(json.dumps(result, ensure_ascii=False))
            return 0
        report = load_json(report_path, {}) or {}
        rows = list(report.get("packages", []))
        if not rows:
            result = {
                "name": name,
                "folder": package.get("folder", ""),
                "status": "PRESERVED_FOR_REVIEW",
                "changed": True,
                "actions_requested": package.get("actions", []),
                "actions_applied": [],
                "failures": ["Word单文档处理未返回报告：" + (output or "")[-1000:]],
                "word_pages": 0,
                "seconds": 0,
            }
            rows = [result]
        # A Word COM server can terminate between completing the final export
        # and PowerShell calling Quit().  The script report is the authoritative
        # result in that case; do not discard a complete report because the
        # wrapper process returns a nonzero code during cleanup.
        atomic_json(args.report, {"version": report.get("version", "3.0.1-item"), "packages": rows})
        print(json.dumps(rows[0], ensure_ascii=False))
        return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def package_dirs(modified: Path) -> list[Path]:
    projects = set()
    for txt in modified.rglob("*.txt"):
        relative = txt.relative_to(modified)
        if any(part.startswith(".") or ".backup-" in part for part in relative.parts):
            continue
        siblings = [path for path in txt.parent.iterdir() if path.is_file()]
        if any(path.suffix.casefold() in {".docx", ".pdf"} for path in siblings):
            projects.add(txt.parent)
    if not projects:
        projects = {
            path for path in modified.iterdir()
            if path.is_dir() and not path.name.startswith(".") and ".backup-" not in path.name
        }
    return sorted(projects, key=lambda path: str(path.relative_to(modified)).casefold())


def remove_delivery_temp_files(modified: Path) -> list[str]:
    removed = []
    for path in sorted((item for item in modified.rglob("*") if item.is_file()), key=lambda item: str(item).casefold()):
        name = path.name.casefold()
        if path.suffix.casefold() in {".tmp", ".temp"} or name.startswith("~$"):
            relative = path.relative_to(modified).as_posix()
            removed.append(relative)
    return removed


def delete_delivery_temp_files(modified: Path, candidates: list[str]) -> None:
    for relative in candidates:
        path = modified / Path(relative)
        if path.is_file():
            path.unlink()


def find_rar() -> Path | None:
    candidates = (
        shutil.which("rar"),
        shutil.which("Rar.exe"),
        r"C:\Program Files\WinRAR\Rar.exe",
        r"C:\Program Files (x86)\WinRAR\Rar.exe",
    )
    return next((Path(item) for item in candidates if item and Path(item).is_file()), None)


def is_delivery_file(path: Path, folder: Path) -> bool:
    relative = path.relative_to(folder)
    name = path.name.casefold()
    if any(part.startswith(".") or ".backup-" in part.casefold() for part in relative.parts):
        return False
    if path.suffix.casefold() in {".tmp", ".temp"} or name.startswith("~$"):
        return False
    return path.is_file()


def delivery_files(folders: list[Path]) -> list[Path]:
    return [
        path
        for folder in folders
        for path in sorted(folder.rglob("*"), key=lambda item: str(item).casefold())
        if is_delivery_file(path, folder)
    ]


def create_zip(output: Path, folders: list[Path], modified: Path) -> None:
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in delivery_files(folders):
            archive.write(path, path.relative_to(modified).as_posix())
    with zipfile.ZipFile(output) as archive:
        broken = archive.testzip()
    if broken is not None:
        raise RuntimeError(f"ZIP integrity test failed: {broken}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Package approved modified software-copyright projects.")
    parser.add_argument("modified_dir", type=Path)
    parser.add_argument("--name", help="Archive basename; defaults to the batch-root directory name")
    parser.add_argument("--format", choices=("rar", "zip"), default="rar", help="Approved delivery archive format")
    parser.add_argument("--layout-approved", action="store_true", help="Confirm the user approved processed manual PDF pages 1-2")
    parser.add_argument("--include-folder", action="append", default=[], help="Package only this project folder relative to modified_dir; repeatable")
    args = parser.parse_args()

    if not args.layout_approved:
        raise SystemExit(json.dumps({
            "status": "NEEDS_LAYOUT_APPROVAL",
            "error": "processed manual PDF pages 1-2 have not been approved",
        }, ensure_ascii=False))

    modified = args.modified_dir.resolve()
    if not modified.is_dir():
        raise SystemExit(json.dumps({"status": "FAILED", "error": f"modified directory missing: {modified}"}, ensure_ascii=False))
    name = args.name or modified.parent.name
    if not name:
        raise SystemExit(json.dumps({"status": "FAILED", "error": "archive name is empty"}, ensure_ascii=False))
    archive_format = args.format
    if archive_format == "rar":
        rar = find_rar()
        if rar is None:
            raise SystemExit(json.dumps({"status": "FAILED", "error": "Rar.exe not found"}, ensure_ascii=False))
    else:
        rar = None
    output = modified / f"{name}.{archive_format}"
    opposite = modified / f"{name}.{'zip' if archive_format == 'rar' else 'rar'}"
    removed_temp_files = remove_delivery_temp_files(modified)
    folders = package_dirs(modified)
    if args.include_folder:
        requested = {Path(value).as_posix().strip("/").casefold() for value in args.include_folder}
        known = {folder.relative_to(modified).as_posix().casefold(): folder for folder in folders}
        missing = sorted(requested - set(known))
        if missing:
            raise SystemExit(json.dumps({"status": "FAILED", "error": "approved project folder missing: " + ", ".join(missing)}, ensure_ascii=False))
        folders = [folder for folder in folders if folder.relative_to(modified).as_posix().casefold() in requested]
    if not folders:
        raise SystemExit(json.dumps({"status": "FAILED", "error": "no processed project directories"}, ensure_ascii=False))

    temp = modified / f".{name}.building.{archive_format}"
    files = delivery_files(folders)
    count = len(files)
    staging_root: Path | None = None
    try:
        if temp.exists():
            temp.unlink()
        if archive_format == "rar":
            staging_root = Path(tempfile.mkdtemp(prefix="easysoftware-package-"))
            staged_folders = []
            for folder in folders:
                relative_folder = folder.relative_to(modified)
                staged_folder = staging_root / relative_folder
                staged_folder.mkdir(parents=True, exist_ok=True)
                for path in (item for item in files if folder == item or folder in item.parents):
                    destination = staging_root / path.relative_to(modified)
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(path, destination)
                staged_folders.append(str(relative_folder))
            create = subprocess.run(
                [str(rar), "a", "-r", "-idq", "-m5", str(temp), *staged_folders],
                cwd=staging_root,
                capture_output=True,
                text=True,
            )
            if create.returncode != 0:
                raise RuntimeError(f"RAR creation failed: {(create.stderr or create.stdout)[-500:]}")
            test = subprocess.run([str(rar), "t", "-idq", str(temp)], capture_output=True, text=True)
            if test.returncode != 0:
                raise RuntimeError(f"RAR integrity test failed: {(test.stderr or test.stdout)[-500:]}")
        else:
            create_zip(temp, folders, modified)
        try:
            os.replace(temp, output)
        except OSError:
            # A prior delivery archive already exists; replace it instead of
            # failing. A genuine lock by another process still surfaces here.
            output.unlink(missing_ok=True)
            os.replace(temp, output)
        delete_delivery_temp_files(modified, removed_temp_files)
    finally:
        if staging_root is not None:
            shutil.rmtree(staging_root, ignore_errors=True)
        if temp.exists():
            temp.unlink()
    removed_opposite_archive = False
    if opposite.exists():
        opposite.unlink()
        removed_opposite_archive = True

    print(json.dumps({
        "status": "OK",
        "archive": str(output),
        "projects": [folder.name for folder in folders],
        "files": count,
        "bytes": output.stat().st_size,
        "sha256": sha256(output),
        "format": archive_format,
        "removed_temp_files": removed_temp_files,
        "removed_opposite_archive": removed_opposite_archive,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

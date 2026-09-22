from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path


VERSION = "1.0.0"
DEFAULT_ROOT = Path(r"C:\Users\82707\Desktop\新建文件夹")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def atomic_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as target:
            json.dump(value, target, ensure_ascii=False, indent=2)
        os.replace(temporary, path)
    except Exception:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        raise


def load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def find_7z() -> str | None:
    candidates = [shutil.which("7z"), shutil.which("7z.exe")]
    candidates.extend((
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        r"C:\Program Files\AMD\CIM\Bin64\7z.exe",
        r"C:\Program Files\AMD\CNext\CNext\7z.exe",
        r"C:\Program Files\NVIDIA Corporation\NVIDIA GeForce Experience\7z.exe",
    ))
    return next((str(candidate) for candidate in candidates if candidate and Path(candidate).exists()), None)


def safe_zip_extract(archive: Path, target: Path) -> None:
    target_resolved = target.resolve()
    with zipfile.ZipFile(archive) as source:
        for member in source.infolist():
            output = (target / member.filename).resolve()
            if output != target_resolved and target_resolved not in output.parents:
                raise ValueError(f"压缩包包含越界路径：{archive.name} -> {member.filename}")
        source.extractall(target)


def extract_archive(archive: Path, target: Path, seven_zip: str | None) -> None:
    target.mkdir(parents=True, exist_ok=False)
    if archive.suffix.casefold() == ".zip":
        safe_zip_extract(archive, target)
        return
    if not seven_zip:
        raise RuntimeError(f"解压RAR需要7-Zip：{archive.name}")
    result = subprocess.run(
        [seven_zip, "x", "-y", f"-o{target}", str(archive)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"解压失败：{archive.name}：{result.stderr[-500:]}")


def prepare_one(archive: Path, root: Path, seven_zip: str | None, force: bool) -> dict:
    started = time.perf_counter()
    digest = sha256(archive)
    destination = root / archive.stem
    marker = root / f".{archive.name}.softcopy-extract.json"
    old = load_json(marker)
    if not force and destination.is_dir() and old.get("archive_sha256") == digest:
        return {
            "archive": archive.name,
            "destination": str(destination),
            "status": "REUSED",
            "backup": None,
            "seconds": round(time.perf_counter() - started, 3),
        }
    if destination.exists() and not destination.is_dir():
        raise RuntimeError(f"同名目标被文件占用：{destination}")

    temporary = root / f".{archive.stem}.extracting-{uuid.uuid4().hex}"
    backup = None
    try:
        extract_archive(archive, temporary, seven_zip)
        if destination.exists():
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S%f")
            backup = root / f"{archive.stem}.backup-{stamp}"
            destination.replace(backup)
        try:
            temporary.replace(destination)
        except Exception:
            if backup and backup.exists() and not destination.exists():
                backup.replace(destination)
            raise
        atomic_json(marker, {
            "version": VERSION,
            "archive": archive.name,
            "archive_sha256": digest,
            "destination": str(destination),
            "backup": str(backup) if backup else None,
            "extracted_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        })
        return {
            "archive": archive.name,
            "destination": str(destination),
            "status": "EXTRACTED",
            "backup": str(backup) if backup else None,
            "seconds": round(time.perf_counter() - started, 3),
        }
    except Exception:
        if temporary.exists() and temporary.parent.resolve() == root.resolve():
            shutil.rmtree(temporary, ignore_errors=True)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description="将桌面收件目录中的RAR解压到旁边的同名文件夹")
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--force", action="store_true", help="强制重新解压；旧同名文件夹会保留为时间戳备份")
    args = parser.parse_args()
    root = args.root.resolve()
    if not root.is_dir():
        raise SystemExit(json.dumps({"status": "FAILED", "error": f"目录不存在：{root}"}, ensure_ascii=False))
    archives = sorted((path for path in root.iterdir() if path.is_file() and path.suffix.lower() == ".rar"), key=lambda p: p.name.lower())
    duplicate_stems = {path.stem.lower() for path in archives if sum(other.stem.lower() == path.stem.lower() for other in archives) > 1}
    if duplicate_stems:
        raise SystemExit(json.dumps({"status": "FAILED", "error": f"存在同名RAR，目标文件夹会冲突：{sorted(duplicate_stems)}"}, ensure_ascii=False))
    if not archives:
        print(json.dumps({"status": "NO_ARCHIVES", "root": str(root), "archives": 0, "results": []}, ensure_ascii=False))
        return 0
    seven_zip = find_7z()
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, min(args.workers, 4))) as pool:
        futures = [pool.submit(prepare_one, archive, root, seven_zip, args.force) for archive in archives]
        for future in concurrent.futures.as_completed(futures):
            results.append(future.result())
    results.sort(key=lambda row: row["archive"].lower())
    print(json.dumps({
        "status": "OK",
        "root": str(root),
        "archives": len(results),
        "extracted": sum(row["status"] == "EXTRACTED" for row in results),
        "reused": sum(row["status"] == "REUSED" for row in results),
        "results": results,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

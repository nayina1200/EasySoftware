from __future__ import annotations

import argparse
import shutil
import subprocess
import zipfile
from pathlib import Path

import archive_types


def safe_zip_extract(archive: Path, target: Path) -> None:
    target_resolved = target.resolve()
    with zipfile.ZipFile(archive) as zf:
        for member in zf.infolist():
            output = (target / member.filename).resolve()
            if target_resolved != output and target_resolved not in output.parents:
                raise ValueError(f"unsafe archive member: {member.filename}")
        zf.extractall(target)


def extract_one(archive: Path, target: Path) -> None:
    target.mkdir(parents=True, exist_ok=False)
    if archive.suffix.casefold() == ".zip":
        safe_zip_extract(archive, target)
        return
    candidates = [
        shutil.which("7z"),
        shutil.which("7z.exe"),
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        r"C:\Program Files\AMD\CIM\Bin64\7z.exe",
        r"C:\Program Files\AMD\CNext\CNext\7z.exe",
        r"C:\Program Files\NVIDIA Corporation\NVIDIA GeForce Experience\7z.exe",
    ]
    seven_zip = next((str(candidate) for candidate in candidates if candidate and Path(candidate).exists()), None)
    if not seven_zip:
        raise RuntimeError(f"7z is required for non-ZIP archive: {archive.name}")
    subprocess.run([seven_zip, "x", "-y", f"-o{target}", str(archive)], check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract each software-copyright archive into an isolated folder.")
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    args = parser.parse_args()
    args.destination.mkdir(parents=True, exist_ok=True)
    if any(args.destination.iterdir()):
        raise SystemExit("destination must be empty")
    archives = [p for p in sorted(args.source.iterdir()) if p.is_file() and archive_types.is_archive_file(p)]
    if not archives:
        raise SystemExit("批次目录没有可解压的压缩包")
    for archive in archives:
        target = args.destination / archive.stem
        extract_one(archive, target)
        print(f"{archive.name}\t{target}")


if __name__ == "__main__":
    main()

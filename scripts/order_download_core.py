from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


def safe_component(value: str) -> str:
    """Return a Windows-safe single path component for a customer name."""
    cleaned = INVALID_WINDOWS_CHARS.sub("_", str(value)).strip().rstrip(".")
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned:
        raise ValueError("客户名称为空或无法用于文件名")
    if cleaned.split(".", 1)[0].upper() in WINDOWS_RESERVED_NAMES:
        cleaned = "_" + cleaned
    return cleaned[:120].rstrip(" .") or "客户"


def validate_zip(path: Path) -> int:
    if not path.is_file() or path.suffix.casefold() != ".zip":
        raise ValueError("下载结果不是 ZIP 文件")
    if not zipfile.is_zipfile(path):
        raise ValueError("下载文件不是有效 ZIP")
    try:
        with zipfile.ZipFile(path) as archive:
            broken_member = archive.testzip()
            file_count = sum(not member.is_dir() for member in archive.infolist())
    except (OSError, zipfile.BadZipFile) as exc:
        raise ValueError(f"ZIP 校验失败：{exc}") from exc
    if broken_member is not None:
        raise ValueError(f"ZIP 中存在损坏文件：{broken_member}")
    if file_count == 0:
        raise ValueError("ZIP 中没有可下载文件")
    return file_count


def _copy_validated_zip(source: Path, destination: Path, customer: str, file_count: int) -> Path:
    basename = f"{safe_component(customer)}_{file_count}件"
    output_dir = destination.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    batch_dir = output_dir / basename
    batch_dir.mkdir(exist_ok=True)
    output = batch_dir / f"{basename}{source.suffix}"

    # Publish only a fully written file. Hard-link creation is atomic and fails
    # when the final name already exists, unlike replace-style rename.
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output.name}.", suffix=".part", dir=batch_dir
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as target, source.open("rb") as downloaded:
            shutil.copyfileobj(downloaded, target, length=1024 * 1024)
            target.flush()
            os.fsync(target.fileno())
        try:
            os.link(temporary, output)
        except FileExistsError as exc:
            raise FileExistsError(f"目标文件已存在，未覆盖：{output}") from exc
        except OSError as exc:
            raise OSError(f"无法以不覆盖方式原子提交下载文件：{exc}") from exc
    finally:
        temporary.unlink(missing_ok=True)
    return output


def finalize_download(source: Path, destination: Path, customer: str) -> Path:
    """Validate a staged ZIP and copy it without replacing an existing file."""
    source = source.resolve(strict=True)
    file_count = validate_zip(source)
    return _copy_validated_zip(source, destination, customer, file_count)


def main() -> int:
    parser = argparse.ArgumentParser(description="Finalize a validated order-material ZIP download.")
    parser.add_argument("finalize", choices=("finalize",), help=argparse.SUPPRESS)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--destination", type=Path, required=True)
    parser.add_argument("--customer", required=True)
    args = parser.parse_args()

    try:
        source = args.source.resolve(strict=True)
        file_count = validate_zip(source)
        output = _copy_validated_zip(source, args.destination, args.customer, file_count)
    except Exception as exc:
        print(json.dumps({"status": "FAILED", "error": str(exc)}, ensure_ascii=False))
        return 1
    print(json.dumps({
        "status": "OK",
        "path": str(output),
        "directory": str(output.parent),
        "filename": output.name,
        "count": file_count,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""统一声明程序识别的压缩包类型。

输入侧（选择压缩包、批次文件夹扫描、嵌套解压）接受所有常见压缩包类型，
不再只认 RAR/ZIP；交付侧（打包输出）仍然只有 ZIP 和 RAR，因为交付格式
需要保留客户原有的压缩习惯，7z 等格式只作为输入读取。

所有需要判断"这是不是压缩包"的代码都必须从这里取常量，避免各处维护
不同的后缀集合（此前正是因此出现"文件夹入口只找 RAR"的问题）。
"""
from __future__ import annotations

from pathlib import Path

#: 只作为交付（打包输出）格式；必须同时属于 ARCHIVE_SUFFIXES。
DELIVERY_SUFFIXES = {".zip", ".rar"}

#: 可作为输入读取的全部压缩包类型。
ARCHIVE_SUFFIXES = {
    ".7z", ".rar", ".rarx", ".zip", ".jar", ".iso", ".cab",
    ".tar", ".gz", ".tgz", ".bz2", ".tbz", ".tbz2", ".xz", ".txz",
    ".lz4", ".lzma", ".lz", ".zst", ".zstd", ".z", ".pax", ".zpaq",
    ".ace", ".arj", ".lzh", ".lha", ".wim", ".msi", ".deb", ".rpm",
    ".dmg", ".hfs", ".pet", ".vhd", ".vhdx", ".gem",
}

#: 双扩展名压缩包。Path.suffix 只返回最后一节，需要整体识别。
COMPOUND_SUFFIXES = (".tar.gz", ".tar.bz2", ".tar.xz", ".tar.lz4")
ARCHIVE_SUFFIXES |= set(COMPOUND_SUFFIXES)

assert DELIVERY_SUFFIXES <= ARCHIVE_SUFFIXES, "交付格式必须是可读取的压缩包类型"


def archive_suffix(path: str | Path) -> str:
    """返回压缩包扩展名，双扩展名（例如 .tar.gz）整体返回。"""
    name = Path(path).name.casefold()
    for suffix in COMPOUND_SUFFIXES:
        if name.endswith(suffix):
            return suffix
    return Path(path).suffix.casefold()


def is_archive_file(path: str | Path) -> bool:
    """判断路径名是否为受支持的压缩包（只看文件名，避免误读内容）。"""
    return archive_suffix(path) in ARCHIVE_SUFFIXES


def is_delivery_archive(path: str | Path) -> bool:
    return archive_suffix(path) in DELIVERY_SUFFIXES


def delivery_format(path: str | Path) -> str:
    """交付格式名（zip/rar）；不是交付格式时返回空字符串。"""
    suffix = archive_suffix(path)
    return suffix.lstrip(".") if suffix in DELIVERY_SUFFIXES else ""


def extracted_sibling(archive: str | Path) -> Path | None:
    """返回压缩包的同名兄弟目录（本程序解压时的默认目标名），不存在则返回 None。"""
    archive = Path(archive)
    sibling = archive.parent / archive.stem
    return sibling if sibling.is_dir() else None


def is_leftover_archive(archive: str | Path) -> bool:
    """判断压缩包是否只是"已就地解压后的残留"：同名兄弟目录已存在。

    文件夹入口据此区分"已经解压好的材料文件夹"和"待解压的批次压缩包"，
    避免对已经解压的内容重复解压。
    """
    return extracted_sibling(archive) is not None

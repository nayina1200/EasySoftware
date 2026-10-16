"""Durable stage checkpoints for the unified document-processing workflow."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Any


CHECKPOINT_NAME = "document_processing_checkpoint.json"
CHECKPOINT_VERSION = 1


def build_identity(source: Path, material_root: Path, stages: list[str], skip_topic_mismatch: bool) -> dict[str, Any]:
    """Identify the exact input set and choices that a checkpoint belongs to."""
    digest = hashlib.sha256()
    for path in sorted((p for p in material_root.rglob("*") if p.is_file()), key=lambda p: p.as_posix().casefold()):
        try:
            stat = path.stat()
            relative = path.relative_to(material_root).as_posix()
            digest.update(f"{relative}\0{stat.st_size}\0{stat.st_mtime_ns}\n".encode("utf-8", "surrogatepass"))
        except OSError as exc:
            raise RuntimeError(f"无法确认续跑输入是否变化：{path}") from exc
    identity = {
        "source": str(source.resolve()),
        "material_root": str(material_root.resolve()),
        "input_fingerprint": digest.hexdigest(),
        "selected_stages": list(stages),
        "skip_topic_mismatch": bool(skip_topic_mismatch),
    }
    if "manual_login_module" in stages:
        from manual_login_module import RULE_VERSION

        identity["login_module_rule_version"] = RULE_VERSION
    return identity


def load(path: Path, identity: dict[str, Any]) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"断点文件无法读取，请保留后人工检查：{path}") from exc
    if not isinstance(state, dict) or state.get("version") != CHECKPOINT_VERSION:
        raise RuntimeError(f"断点文件版本无法识别，请保留后人工检查：{path}")
    if state.get("identity") != identity:
        raise RuntimeError("发现另一批次或输入已变化的断点。请先选回原输入及原流程，避免覆盖已有处理结果。")
    return state


def save(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    os.replace(temporary, path)


def checkpoint_path(work: Path) -> Path:
    return work / CHECKPOINT_NAME


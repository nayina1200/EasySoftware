"""Versioned, atomic cache for title/body qualification results."""
from __future__ import annotations

import hashlib
import json
import os
import threading
import time
from pathlib import Path
from typing import Any, Iterable


CACHE_SCHEMA_VERSION = 1
_CACHE_LOCK = threading.RLock()


def semantic_input_hash(
    title: str,
    paragraphs: Iterable[dict[str, Any]],
    *,
    strategy_version: str,
    model_id: str,
    prompt_version: str,
    policy_version: str = "1",
    extra_dependencies: dict[str, Any] | None = None,
) -> str:
    """Hash complete semantic inputs and classifier dependencies, excluding positional IDs."""
    body = []
    for paragraph in paragraphs:
        body.append({
            "text": str(paragraph.get("text", "")),
            "style_id": paragraph.get("style_id"),
            "outline_level": paragraph.get("outline_level"),
            "is_heading": bool(paragraph.get("is_heading")),
            "is_toc": bool(paragraph.get("is_toc")),
            "in_table": bool(paragraph.get("in_table")),
            "in_textbox": bool(paragraph.get("in_textbox")),
            "is_title_repeat": bool(paragraph.get("is_title_repeat")),
            "structure_uncertain": bool(paragraph.get("structure_uncertain")),
            "section_title": paragraph.get("section_title"),
            "section_role": paragraph.get("section_role"),
            "is_overview": bool(paragraph.get("is_overview")),
            "is_functional": bool(paragraph.get("is_functional")),
            "is_main_function": bool(paragraph.get("is_main_function")),
        })
    payload = {
        "cache_schema_version": CACHE_SCHEMA_VERSION,
        "title": title,
        "paragraphs": body,
        "strategy_version": strategy_version,
        "model_id": model_id,
        "prompt_version": prompt_version,
        "policy_version": policy_version,
        "extra_dependencies": extra_dependencies or {},
    }
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


class TitleBodyCache:
    """Small JSON result cache. Writes replace atomically; unreadable files are quarantined."""

    def __init__(self, path: str | Path):
        self.path = Path(path)

    def get(self, input_hash: str, paragraphs: list[dict[str, Any]] | None = None) -> dict[str, Any] | None:
        with _CACHE_LOCK:
            data = self._read()
            entry = data.get("entries", {}).get(input_hash)
            if not isinstance(entry, dict) or not isinstance(entry.get("result"), dict):
                return None
            result = json.loads(json.dumps(entry["result"], ensure_ascii=False))
            if str(result.get("status", "")).upper() not in {"MATCH", "MISMATCH"}:
                return None
            if paragraphs is not None:
                # Import at call time to keep the cache usable independently of the parser module.
                from title_body_qualification import validate_evidence_result

                validated = validate_evidence_result(result, paragraphs, allow_rebind=True)
                if validated.get("status") not in {"MATCH", "MISMATCH"}:
                    return None
                if validated.get("evidence_valid") is False:
                    return None
                result = validated
            return result

    def put(self, input_hash: str, result: dict[str, Any]) -> None:
        if not isinstance(result, dict):
            raise TypeError("cache result must be a dictionary")
        if str(result.get("status", "")).upper() not in {"MATCH", "MISMATCH"}:
            return
        with _CACHE_LOCK:
            data = self._read()
            entries = data.setdefault("entries", {})
            entries[input_hash] = {"result": result}
            self._write(data)

    def _read(self) -> dict[str, Any]:
        if not self.path.exists():
            return {"schema_version": CACHE_SCHEMA_VERSION, "entries": {}}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            if (
                not isinstance(data, dict)
                or data.get("schema_version") != CACHE_SCHEMA_VERSION
                or not isinstance(data.get("entries"), dict)
            ):
                raise ValueError("unsupported or malformed title/body cache")
            return data
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError):
            self._quarantine()
            return {"schema_version": CACHE_SCHEMA_VERSION, "entries": {}}

    def _quarantine(self) -> None:
        if not self.path.exists():
            return
        quarantine = self.path.with_name(f"{self.path.name}.corrupt.{time.time_ns()}")
        try:
            os.replace(self.path, quarantine)
        except OSError:
            # A bad cache is a miss even if the filesystem prevents moving it aside.
            pass

    def _write(self, data: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_name(f"{self.path.name}.{os.getpid()}.{threading.get_ident()}.tmp")
        try:
            with temporary.open("w", encoding="utf-8", newline="\n") as stream:
                json.dump(data, stream, ensure_ascii=False, sort_keys=True, indent=2)
                stream.write("\n")
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.path)
        finally:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass


__all__ = ["CACHE_SCHEMA_VERSION", "TitleBodyCache", "semantic_input_hash"]

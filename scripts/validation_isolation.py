"""Path guards used by the isolated optimization-development copy."""
from __future__ import annotations

import os
import ctypes
from pathlib import Path

DEVELOPMENT_MARKER = Path(__file__).resolve().parents[1] / ".optimization-dev"
FILE_ATTRIBUTE_REPARSE_POINT = 0x0400
INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF


class ValidationIsolationError(RuntimeError):
    """Raised when a development workflow points outside its validation area."""


def activate_optimization_isolation(environ: dict[str, str] | None = None) -> bool:
    """Enable the fixed sibling validation root for direct development CLI runs."""
    if not DEVELOPMENT_MARKER.is_file():
        return False
    environment = os.environ if environ is None else environ
    validation_root_path = DEVELOPMENT_MARKER.parent.parent / "EasySoftware_优化验证材料"
    _reject_reparse_components(validation_root_path)
    fixed_root = _resolved(validation_root_path)
    configured_root = environment.get("EASYSOFTWARE_VALIDATION_ROOT", "").strip()
    if configured_root and _resolved(configured_root) != fixed_root:
        raise ValidationIsolationError(f"优化开发版隔离根不可通过环境变量改写：必须使用 {fixed_root}")
    isolated_directories = {
        "TEMP": fixed_root / "temp",
        "PYTHONPYCACHEPREFIX": fixed_root / "cache" / "python",
        "YOLO_CONFIG_DIR": fixed_root / "cache" / "yolo",
        "EASYSOFTWARE_USER_DATA_ROOT": fixed_root / "state" / "user-data",
    }
    for isolated_path in isolated_directories.values():
        _reject_reparse_components(isolated_path)
        isolated_path.mkdir(parents=True, exist_ok=True)
    environment["TMP"] = str(isolated_directories["TEMP"])
    for name, isolated_path in isolated_directories.items():
        environment[name] = str(isolated_path)
    environment["EASYSOFTWARE_DATE_TOOL_ROOT"] = str(fixed_root / "dependencies" / "date_ocr")
    environment["EASYSOFTWARE_LLAMA_HOME"] = str(fixed_root / "dependencies" / "llama")
    environment["EASYSOFTWARE_VISION_CONFIG"] = str(fixed_root / "config" / "vision_config.json")
    environment["EASYSOFTWARE_VALIDATION_ROOT"] = str(fixed_root)
    environment["EASYSOFTWARE_VALIDATION_RUNS_ROOT"] = str(fixed_root / "runs")
    environment["EASYSOFTWARE_DEVELOPMENT_READ_ONLY"] = "1"
    return True


def _resolved(path: str | os.PathLike[str]) -> Path:
    return Path(path).expanduser().resolve()


def _inside(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def _reparse_attribute(path: Path) -> bool:
    if not path.exists():
        return False
    if os.name != "nt":
        return path.is_symlink()
    attributes = ctypes.windll.kernel32.GetFileAttributesW(str(path))
    return attributes != INVALID_FILE_ATTRIBUTES and bool(attributes & FILE_ATTRIBUTE_REPARSE_POINT)


def _reject_reparse_components(path: Path) -> None:
    candidate = Path(os.path.abspath(os.fspath(path)))
    components = [candidate, *candidate.parents]
    for component in reversed(components):
        if _reparse_attribute(component):
            raise ValidationIsolationError(f"优化开发版隔离不接受链接或联接路径：{component}")


def require_validation_path(
    path: str | os.PathLike[str],
    *,
    kind: str = "workspace",
    environ: dict[str, str] | None = None,
    same_run_as: str | os.PathLike[str] | None = None,
) -> Path:
    """Resolve and validate a path when the development-copy guard is enabled.

    ``kind='material'`` may point to any local material path. Workspace files
    remain restricted to dedicated temp/cache/report/state roots. Optimization
    CLI entry points call ``activate_optimization_isolation`` before dispatch;
    imported workflow functions remain independently testable against fixtures.
    """
    environment = os.environ if environ is None else environ
    configured_root = environment.get("EASYSOFTWARE_VALIDATION_ROOT", "").strip()
    development_copy = DEVELOPMENT_MARKER.is_file() and bool(configured_root)
    if development_copy:
        fixed_root = _resolved(DEVELOPMENT_MARKER.parent.parent / "EasySoftware_优化验证材料")
        if _resolved(configured_root) != fixed_root:
            raise ValidationIsolationError(f"优化开发版隔离根不可通过环境变量改写：必须使用 {fixed_root}")
    resolved = _resolved(path)
    if not configured_root:
        return resolved

    _reject_reparse_components(Path(path))
    root = _resolved(configured_root)
    if kind == "material":
        return resolved
    elif kind == "workspace":
        work_roots = [_resolved(root / name) for name in ("temp", "cache", "reports", "state")]
        allowed = any(_inside(resolved, work_root) for work_root in work_roots) and all(_inside(work_root, root) for work_root in work_roots)
        location = "专用 temp/cache/reports/state 工作目录"
    else:
        raise ValueError(f"未知隔离路径类型：{kind}")

    if not allowed:
        raise ValidationIsolationError(
            f"优化开发版工作区隔离：路径必须位于{location}内：{resolved}"
        )
    return resolved

from __future__ import annotations

import inspect
import base64
import io
import json
import math
import os
import queue
import re
import shutil
import subprocess
import sys
import time
import tempfile
import threading
import zipfile
import urllib.request
import urllib.error
from urllib.parse import urlsplit, urlunsplit
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from lxml import etree


TARGET_YEAR = 2025
CONFIDENCE_THRESHOLD = 0.85
VISION_MODEL = "sensenova-6.8-flash-lite"
VISION_BASE_URL = "http://127.0.0.1:3458/v1"
VISION_CONFIG_PATH = Path(__file__).resolve().parent.parent / "vision_config.json"
VISION_MAX_CALLS_PER_PROJECT = 2
DATE_REWRITER_TIMEOUT_SECONDS = 900
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
YEAR_RE = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
GENERIC_TERMS = {"系统", "平台", "软件", "管理", "功能", "模块", "用户", "操作", "页面", "数据", "信息"}
# These are domain terms, not project names. They provide explainable evidence for
# manual review when a document repeatedly describes a different physical process.
BODY_TOPIC_TERMS = {
    "照明": ("路灯", "灯具", "灯杆", "调光"),
    "储能": ("储能", "充放电", "蓄电池", "电池组"),
    "切片": ("金刚线", "切片", "硅棒", "晶圆"),
    "射频": ("毫米波", "射频", "芯片", "天线"),
    "散热供暖": ("散热", "制热", "热泵", "供暖", "水泵"),
    "铝箔编织": ("铝箔", "编织", "压线", "夹编织"),
    "移动协议": ("TD-SCDMA", "协议栈", "终端软件", "RTOS"),
    "视觉识别": ("色块", "摄像头", "图像直方图", "RGB565"),
}
BODY_GENERIC_OPERATIONS = ("新增", "删除", "导出", "查询")


def equidistant_sample_indexes(total: int, ratio: float = 0.25, minimum: int = 4) -> list[int]:
    """Return stable zero-based indexes, including both ends when possible."""
    if total <= 0:
        return []
    count = min(total, max(minimum, math.ceil(total * ratio)))
    if count == 1:
        return [0]
    return sorted({round(index * (total - 1) / (count - 1)) for index in range(count)})


def classify_year_context(text: str, year: str | None = None, target_year: int = TARGET_YEAR) -> dict[str, Any]:
    """Conservative deterministic baseline for a single year occurrence."""
    value = year or (YEAR_RE.search(text).group(0) if YEAR_RE.search(text) else "")
    if not value or int(value) == target_year or int(value) > 2026:
        return {"decision": "KEEP", "confidence": 1.0, "reason": "目标年份、远期年份或未发现年份"}
    keep_patterns = (
        r"(?:历史|成立于|发布于|颁布于|修订于|自).{0,8}" + re.escape(value),
        re.escape(value) + r".{0,8}(?:年生|年成立|年发布|年颁布|年修订|年度数据|年统计)",
        r"(?:GB|GB/T|ISO|IEC|RFC|标准|法规|条例|法律|版权|©|版本|合同|案例).{0,16}" + re.escape(value),
    )
    update_patterns = (
        r"(?:当前|目前|现行|最新|截至|截止|今年|本年度|本年|近期|现阶段).{0,16}" + re.escape(value),
        re.escape(value) + r".{0,12}(?:当前|最新版|最新版本|最新|现行|本年度|年度版)",
    )
    if any(re.search(pattern, text, re.IGNORECASE) for pattern in keep_patterns):
        return {"decision": "KEEP", "confidence": 0.95, "reason": "历史、标准、版本或固定记录语境"}
    if any(re.search(pattern, text, re.IGNORECASE) for pattern in update_patterns):
        return {"decision": "UPDATE_TO_2025", "confidence": 0.95, "reason": "当前或最新时效语境"}
    return {"decision": "REVIEW", "confidence": 0.5, "reason": "缺少足够的时效语境"}


def replace_update_years(
    text: str,
    classifier: Callable[..., Any] | None = None,
    target_year: int = TARGET_YEAR,
) -> tuple[str, list[dict[str, Any]]]:
    decisions: list[dict[str, Any]] = []
    pieces: list[str] = []
    cursor = 0
    for match in YEAR_RE.finditer(text):
        pieces.append(text[cursor:match.start()])
        context = text[max(0, match.start() - 32):min(len(text), match.end() + 32)]
        if int(match.group(0)) == target_year or int(match.group(0)) > 2026:
            raw = {"decision": "KEEP", "confidence": 1.0, "reason": "目标年份或未来年份"}
        else:
            raw = _call_adapter(classifier, context, match.group(0)) if classifier else classify_year_context(context, match.group(0), target_year)
        decision = _normalise_year_decision(raw)
        replacement = str(target_year) if decision["decision"] == "UPDATE_TO_TARGET_YEAR" else match.group(0)
        pieces.append(replacement)
        reported_decision = decision["decision"]
        if reported_decision == "UPDATE_TO_TARGET_YEAR" and target_year == 2025:
            reported_decision = "UPDATE_TO_2025"
        decisions.append({
            "source": match.group(0),
            "replacement": replacement,
            "context": context,
            **decision,
            "decision": reported_decision,
        })
        cursor = match.end()
    pieces.append(text[cursor:])
    return "".join(pieces), decisions


def _normalise_year_decision(value: Any) -> dict[str, Any]:
    if isinstance(value, str):
        value = {"decision": value}
    if not isinstance(value, dict):
        return {"decision": "REVIEW", "confidence": 0.0, "reason": "分类器返回格式无效"}
    decision = str(value.get("decision", "REVIEW")).upper()
    # Accept the prior public classifier label for callers that inject a legacy adapter.
    if decision == "UPDATE_TO_2025":
        decision = "UPDATE_TO_TARGET_YEAR"
    if decision not in {"UPDATE_TO_TARGET_YEAR", "KEEP", "REVIEW"}:
        decision = "REVIEW"
    confidence = float(value.get("confidence", 0.0))
    if decision == "UPDATE_TO_TARGET_YEAR" and confidence < CONFIDENCE_THRESHOLD:
        decision = "REVIEW"
    return {
        "decision": decision,
        "confidence": confidence,
        "reason": str(value.get("reason", "")),
    }


def _call_adapter(adapter: Callable[..., Any] | None, *args: Any) -> Any:
    if adapter is None:
        return None
    try:
        count = len(inspect.signature(adapter).parameters)
    except (TypeError, ValueError):
        count = len(args)
    return adapter(*args[:count])


class DefaultOcrAdapter:
    def __init__(self, script_dir: Path):
        self.script_dir = script_dir
        self.backend = "unavailable"
        self.detail = "RapidOCR 和 Windows OCR 均不可用"
        self._rapid = None
        self._performance = {"batch_calls": 0, "image_count": 0, "batch_seconds": 0.0}
        configured_tool_root = os.getenv("EASYSOFTWARE_DATE_TOOL_ROOT")
        self.tool_root = (
            Path(configured_tool_root).expanduser().resolve()
            if configured_tool_root
            else script_dir.parent.parent / "图片日期智能处理工具_用户版_20260903"
        )
        internal = self.tool_root / "_internal"
        self.external_python = internal / "python.exe"
        self.external_bridge = script_dir / "rapid_ocr_bridge.py"
        external_available = self.external_python.is_file() and self.external_bridge.is_file()
        if external_available:
            self.backend = "rapidocr_subprocess"
            self.detail = str(self.tool_root)
        elif internal.is_dir():
            try:
                sys.path.insert(0, str(internal))
                from rapidocr import RapidOCR

                self._rapid = RapidOCR(params={"Global.text_score": 0.2})
                self.backend = "rapidocr"
                self.detail = str(internal)
            except Exception as exc:  # optional packaged runtime
                self.detail = f"RapidOCR 初始化失败：{exc}"
            finally:
                if str(internal) in sys.path:
                    sys.path.remove(str(internal))
        windows_script = script_dir / "windows_ocr.ps1"
        if self._rapid is None and os.name == "nt" and windows_script.is_file():
            if self.backend == "unavailable":
                self.backend = "windows_ocr"
                self.detail = str(windows_script)

    def __call__(self, image_bytes: bytes, suffix: str = ".png") -> dict[str, Any]:
        if self._rapid is not None:
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as handle:
                handle.write(image_bytes)
                temporary = Path(handle.name)
            try:
                result = self._rapid(str(temporary))
                texts = getattr(result, "txts", None) or []
                scores = getattr(result, "scores", None) or []
                if not texts and isinstance(result, (list, tuple)) and result:
                    rows = result[0] or []
                    texts = [row[1] for row in rows if isinstance(row, (list, tuple)) and len(row) >= 2]
                    scores = [row[2] for row in rows if isinstance(row, (list, tuple)) and len(row) >= 3]
                return {"text": "\n".join(map(str, texts)), "confidence": min(map(float, scores)) if scores else 0.0, "backend": self.backend}
            except Exception as exc:
                return {"text": "", "confidence": 0.0, "backend": self.backend, "error": str(exc)}
            finally:
                temporary.unlink(missing_ok=True)
        if self.backend == "rapidocr_subprocess":
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as handle:
                handle.write(image_bytes)
                temporary = Path(handle.name)
            try:
                completed = subprocess.run(
                    [str(self.external_python), str(self.external_bridge), str(self.tool_root), str(temporary)],
                    capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
                )
                lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
                payload = json.loads(lines[-1]) if lines else {}
                if completed.returncode != 0:
                    payload.setdefault("error", completed.stderr.strip() or "RapidOCR 子进程失败")
                return payload
            except Exception as exc:
                return {"text": "", "confidence": 0.0, "backend": self.backend, "error": str(exc)}
            finally:
                temporary.unlink(missing_ok=True)
        if self.backend == "windows_ocr":
            suffix = suffix if suffix.lower() in IMAGE_SUFFIXES else ".png"
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as handle:
                handle.write(image_bytes)
                temporary = Path(handle.name)
            try:
                result = subprocess.run(
                    ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", self.detail, "-ImagePath", str(temporary)],
                    capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=90,
                )
                return {"text": result.stdout.strip() if result.returncode == 0 else "", "confidence": 0.0, "backend": self.backend, "error": result.stderr.strip() or None}
            except Exception as exc:
                return {"text": "", "confidence": 0.0, "backend": self.backend, "error": str(exc)}
            finally:
                temporary.unlink(missing_ok=True)
        return {"text": "", "confidence": 0.0, "backend": self.backend, "error": self.detail}

    def batch(self, images: list[tuple[bytes, str]]) -> list[dict[str, Any]]:
        if not images:
            return []
        perf_trace = os.getenv("EASYSOFTWARE_PERF_TRACE") == "1"
        started = time.perf_counter() if perf_trace else None
        try:
            return self._batch_images(images)
        finally:
            if perf_trace:
                self._performance["batch_calls"] += 1
                self._performance["image_count"] += len(images)
                self._performance["batch_seconds"] += time.perf_counter() - started

    def _batch_images(self, images: list[tuple[bytes, str]]) -> list[dict[str, Any]]:
        if self.backend != "rapidocr_subprocess":
            return [self(data, suffix) for data, suffix in images]
        with tempfile.TemporaryDirectory(prefix="easysoftware-ocr-batch-") as directory:
            paths = []
            for index, (data, suffix) in enumerate(images):
                path = Path(directory) / f"image-{index:04d}{suffix if suffix.lower() in IMAGE_SUFFIXES else '.png'}"
                path.write_bytes(data)
                paths.append(path)
            try:
                completed = subprocess.run(
                    [str(self.external_python), str(self.external_bridge), str(self.tool_root), *map(str, paths)],
                    capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=max(180, 90 * len(paths)),
                )
                lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
                payload = json.loads(lines[-1]) if lines else {}
                rows = payload.get("results") if isinstance(payload, dict) else None
                if len(images) == 1 and isinstance(payload, dict) and "text" in payload:
                    rows = [payload]
                if not isinstance(rows, list) or len(rows) != len(images):
                    raise RuntimeError(completed.stderr.strip() or "RapidOCR 批量结果数量不匹配")
                return [row if isinstance(row, dict) else {"text": "", "confidence": 0.0, "backend": self.backend} for row in rows]
            except Exception as exc:
                return [{"text": "", "confidence": 0.0, "backend": self.backend, "error": str(exc)} for _ in images]


def load_vision_config(path: str | Path | None = None) -> dict[str, Any]:
    """Load the editable OpenAI-compatible vision endpoint configuration."""
    config_path = Path(path) if path else Path(os.getenv("EASYSOFTWARE_VISION_CONFIG", VISION_CONFIG_PATH))
    if not config_path.is_file():
        return {}
    try:
        value = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"视觉模型配置读取失败（{config_path}）：{exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"视觉模型配置必须是 JSON 对象：{config_path}")
    return value


def _safe_endpoint(value: str) -> str:
    parts = urlsplit(value)
    host = parts.hostname or ""
    if parts.port:
        host += f":{parts.port}"
    return urlunsplit((parts.scheme, host, parts.path, "", ""))


class OpenAICompatibleVisionAdapter:
    """OpenAI-compatible image reviewer with bounded payloads and API-key rotation."""

    backend = "openai_compatible"

    def __init__(
        self,
        base_url: str | None = None,
        api_key: str | None = None,
        timeout: float | None = None,
        *,
        model: str | None = None,
        config: dict[str, Any] | None = None,
    ):
        settings = dict(load_vision_config() if config is None else config)
        self.base_url = str(base_url or os.getenv("EASYSOFTWARE_VISION_BASE_URL") or settings.get("base_url") or VISION_BASE_URL).rstrip("/")
        self.model = str(model or settings.get("model") or VISION_MODEL)
        env_key = os.getenv("EASYSOFTWARE_VISION_API_KEY")
        configured_keys = settings.get("api_keys", settings.get("api_key", []))
        if isinstance(configured_keys, str):
            configured_keys = [configured_keys]
        if not isinstance(configured_keys, list):
            configured_keys = []
        key_values = [api_key] if api_key else ([env_key] if env_key else configured_keys)
        self.api_keys = [str(item).strip() for item in key_values if str(item).strip()]
        if not self.api_keys:
            self.api_keys = ["proxy"]
        self.timeout = float(timeout if timeout is not None else settings.get("timeout_seconds", 30))
        self.max_image_bytes = max(16 * 1024, int(settings.get("max_image_bytes", 4_000_000)))
        self.max_image_side = max(64, int(settings.get("max_image_side", 1600)))
        self.jpeg_quality = max(40, min(95, int(settings.get("jpeg_quality", 85))))
        self.retry_count = max(0, min(5, int(settings.get("retry_count", 2))))
        self._next_key = 0

    def _prepare_image(self, image_bytes: bytes) -> tuple[bytes, str]:
        from PIL import Image, ImageOps

        with Image.open(io.BytesIO(image_bytes)) as original:
            image = ImageOps.exif_transpose(original).convert("RGB")
            image.thumbnail((self.max_image_side, self.max_image_side), Image.Resampling.LANCZOS)
            quality = self.jpeg_quality
            while True:
                output = io.BytesIO()
                image.save(output, format="JPEG", quality=quality, optimize=True)
                prepared = output.getvalue()
                if len(prepared) <= self.max_image_bytes:
                    return prepared, "image/jpeg"
                if quality > 50:
                    quality -= 10
                    continue
                width, height = image.size
                if max(width, height) <= 512:
                    raise ValueError(f"图片压缩后仍超过上限 {self.max_image_bytes} 字节")
                image.thumbnail((max(512, int(width * .75)), max(512, int(height * .75))), Image.Resampling.LANCZOS)
                quality = min(self.jpeg_quality, 75)

    def __call__(self, image_bytes: bytes, suffix: str, title: str, context: str, ocr_text: str) -> dict[str, Any]:
        original_size = len(image_bytes)
        try:
            image_bytes, mime = self._prepare_image(image_bytes)
        except Exception as exc:
            return {"status": "unknown", "confidence": 0.0, "reason": "图片预处理失败", "backend": self.backend, "image_original_bytes": original_size, "image_max_bytes": self.max_image_bytes, "error": str(exc), "error_type": type(exc).__name__}
        image_metrics = {
            "image_original_bytes": original_size,
            "image_sent_bytes": len(image_bytes),
            "image_max_bytes": self.max_image_bytes,
            "image_sent_mime": mime,
        }
        prompt = (
            "核对这张软件说明书截图与项目标题的主题是否一致。只依据图片中可见的界面、文字和图表，"
            "不要把标题或 OCR 文本当作图片证据。仅当图片明确属于另一领域时判 irrelevant；"
            "看不清、通用界面、证据不足时判 unknown。请输出 JSON 对象，字段 status "
            "(relevant/irrelevant/unknown)、confidence (0 到 1)、evidence (具体可见证据)。\n"
            f"项目标题：{title[:160]}\n正文开头：{context[:300]}\nOCR 提示：{ocr_text[:300]}"
        )
        payload = {
            "model": self.model,
            "messages": [{"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{base64.b64encode(image_bytes).decode('ascii')}"}},
            ]}],
            "max_tokens": 1200,
            "reasoning_effort": "none",
        }
        endpoint = self.base_url if self.base_url.endswith("/chat/completions") else f"{self.base_url}/chat/completions"
        reported_endpoint = _safe_endpoint(endpoint)
        errors: list[str] = []
        attempts = min(len(self.api_keys), self.retry_count + 1)
        reply: dict[str, Any] | None = None
        for attempt in range(attempts):
            key_index = self._next_key % len(self.api_keys)
            self._next_key += 1
            request = urllib.request.Request(
                endpoint,
                data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                headers={"Authorization": f"Bearer {self.api_keys[key_index]}", "Content-Type": "application/json"},
                method="POST",
            )
            try:
                with urllib.request.urlopen(request, timeout=self.timeout) as response:
                    reply = json.load(response)
                break
            except urllib.error.HTTPError as exc:
                detail = exc.read(1200).decode("utf-8", errors="replace")
                for secret in self.api_keys:
                    detail = detail.replace(secret, "<REDACTED>")
                errors.append(f"HTTP {exc.code}: {detail or exc.reason}")
                if attempt + 1 >= attempts or exc.code not in {401, 403, 408, 413, 429, 500, 502, 503, 504}:
                    return {"status": "unknown", "confidence": 0.0, "reason": "视觉接口 HTTP 请求失败", "backend": self.backend, "model": self.model, "endpoint": reported_endpoint, **image_metrics, "error": errors[-1], "error_type": "HTTPError"}
            except Exception as exc:
                detail = str(exc)
                for secret in self.api_keys:
                    detail = detail.replace(secret, "<REDACTED>")
                errors.append(f"{type(exc).__name__}: {detail}")
                if attempt + 1 >= attempts:
                    return {"status": "unknown", "confidence": 0.0, "reason": "视觉接口连接或请求失败", "backend": self.backend, "model": self.model, "endpoint": reported_endpoint, **image_metrics, "error": "；".join(errors), "error_type": type(exc).__name__}
        if reply is None:
            return {"status": "unknown", "confidence": 0.0, "reason": "视觉接口重试次数耗尽", "backend": self.backend, "model": self.model, "endpoint": reported_endpoint, **image_metrics, "error": "；".join(errors), "error_type": "RetryError"}
        try:
            message = reply["choices"][0]["message"]
            content = message.get("content") or ""
            if isinstance(content, list):
                content = "\n".join(part.get("text", "") for part in content if isinstance(part, dict))
            match = re.search(r"\{[\s\S]*\}", str(content))
            if not match:
                raise ValueError("视觉模型未返回 JSON content")
            verdict = json.loads(match.group(0))
            if not isinstance(verdict, dict):
                raise ValueError("视觉模型 JSON 不是对象")
            evidence = str(verdict.get("evidence", "")).strip()
            result = _normalise_relevance({
                "status": verdict.get("status"), "confidence": verdict.get("confidence", 0),
                "reason": evidence,
            }, CONFIDENCE_THRESHOLD)
            if result["status"] == "irrelevant" and len(evidence) < 8:
                result["status"] = "unknown"
                result["reason"] = "视觉模型未提供充分的图片证据"
            return {**result, "backend": self.backend, "model": self.model, "endpoint": reported_endpoint, **image_metrics, "usage": reply.get("usage", {})}
        except Exception as exc:
            detail = str(exc)
            for secret in self.api_keys:
                detail = detail.replace(secret, "<REDACTED>")
            return {"status": "unknown", "confidence": 0.0, "reason": "视觉接口未给出可用结论", "backend": self.backend, "model": self.model, "endpoint": reported_endpoint, **image_metrics, "error": detail, "error_type": type(exc).__name__}


LocalVisionAdapter = OpenAICompatibleVisionAdapter


class DefaultDateRewriter:
    """Invoke the existing packaged date engine through its matching Python runtime."""

    def __init__(self, script_dir: Path):
        self.bridge = script_dir / "date_year_rewriter_bridge.py"
        configured_tool_root = os.getenv("EASYSOFTWARE_DATE_TOOL_ROOT")
        self.tool_root = (
            Path(configured_tool_root).expanduser().resolve()
            if configured_tool_root
            else script_dir.parent.parent / "图片日期智能处理工具_用户版_20260903"
        )
        self.python = self.tool_root / "_internal" / "python.exe"
        self.available = self.bridge.is_file() and self.python.is_file() and (self.tool_root / "app.py").is_file()
        self.reason = "未找到图片日期工具的运行时或年份重绘桥接脚本"
        self._performance: dict[str, Any] = {}

    def __call__(self, path: Path, target_year: int) -> dict[str, Any]:
        return self.rewrite_with_metadata(path, target_year)

    def rewrite_with_metadata(self, path: Path, target_year: int) -> dict[str, Any]:
        if not self.available:
            return {"status": "UNAVAILABLE", "changed": False, "reason": self.reason, "effective_target_year": target_year, "trigger_date": None, "rollback_reason": None}
        perf_trace = os.getenv("EASYSOFTWARE_PERF_TRACE") == "1"
        started = time.perf_counter() if perf_trace else None
        self._performance["subprocess_calls"] = self._performance.get("subprocess_calls", 0) + 1
        call_started = time.perf_counter()
        try:
            completed = subprocess.run(
                [str(self.python), str(self.bridge), str(self.tool_root), str(path), str(target_year)],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=DATE_REWRITER_TIMEOUT_SECONDS,
            )
        except subprocess.TimeoutExpired:
            self._performance["wall_seconds"] = self._performance.get("wall_seconds", 0.0) + time.perf_counter() - call_started
            return {"status": "FAILED_TIMEOUT", "changed": False, "reason": "该 DOCX 日期 OCR 超过 900 秒，已终止子进程"}
        lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
        try:
            result = json.loads(lines[-1]) if lines else {}
        except json.JSONDecodeError:
            result = {}
        if completed.returncode != 0 or not isinstance(result, dict):
            failure = {
                "status": "FAILED", "changed": False,
                "reason": result.get("error") if isinstance(result, dict) else (completed.stderr.strip() or completed.stdout.strip()),
            }
            if perf_trace:
                failure["performance"] = {"subprocess_seconds": round(time.perf_counter() - started, 4)}
            self._performance["wall_seconds"] = self._performance.get("wall_seconds", 0.0) + time.perf_counter() - call_started
            return failure
        if perf_trace:
            result.setdefault("performance", {})["subprocess_seconds"] = round(time.perf_counter() - started, 4)
            self._performance["engine_load_seconds"] = self._performance.get("engine_load_seconds", 0.0) + float(
                result.get("performance", {}).get("engine_load_seconds", 0.0) or 0.0
            )
        self._performance["wall_seconds"] = self._performance.get("wall_seconds", 0.0) + time.perf_counter() - call_started
        return result

    def iter_rewrite_many_with_metadata(self, paths: list[Path], target_year: int):
        """Stream one DOCX at a time through one child with an independent 900s deadline."""
        if not self.available:
            for _path in paths:
                yield {"status": "UNAVAILABLE", "changed": False, "reason": self.reason,
                       "effective_target_year": target_year, "trigger_date": None, "rollback_reason": None}
            return
        perf_trace = os.getenv("EASYSOFTWARE_PERF_TRACE") == "1"
        launch_started = time.perf_counter()
        process = None
        reader_thread = None
        batch_recorded = False
        responses: queue.Queue[str | None] = queue.Queue()
        self._performance["subprocess_calls"] = self._performance.get("subprocess_calls", 0) + 1
        self._performance["batch_count"] = self._performance.get("batch_count", 0) + 1

        def record_batch_wall():
            nonlocal batch_recorded
            if not batch_recorded:
                elapsed = time.perf_counter() - launch_started
                self._performance["batch_wall_seconds"] = self._performance.get("batch_wall_seconds", 0.0) + elapsed
                self._performance["wall_seconds"] = self._performance.get("wall_seconds", 0.0) + elapsed
                batch_recorded = True

        try:
            process = subprocess.Popen(
                [str(self.python), str(self.bridge), "--stream", str(self.tool_root), str(target_year)],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                text=True, encoding="utf-8", errors="replace", bufsize=1,
            )

            def read_responses():
                try:
                    for line in process.stdout:
                        responses.put(line)
                finally:
                    responses.put(None)

            reader_thread = threading.Thread(target=read_responses, daemon=True)
            reader_thread.start()
            ready_line = responses.get(timeout=DATE_REWRITER_TIMEOUT_SECONDS)
            ready = json.loads(ready_line) if ready_line else {}
            if not isinstance(ready, dict) or not ready.get("ready"):
                raise RuntimeError(str(ready.get("error", "日期 OCR 批处理子进程启动失败")))
            if perf_trace:
                engine_seconds = float(ready.get("performance", {}).get("engine_load_seconds", 0.0) or 0.0)
                self._performance["engine_load_seconds"] = self._performance.get("engine_load_seconds", 0.0) + engine_seconds

            for index, path in enumerate(paths):
                try:
                    process.stdin.write(json.dumps({"op": "rewrite", "path": str(path)}) + "\n")
                    process.stdin.flush()
                    process_deadline = time.perf_counter() + DATE_REWRITER_TIMEOUT_SECONDS
                    response = responses.get(timeout=max(0.01, process_deadline - time.perf_counter()))
                    if response is None:
                        raise RuntimeError("日期 OCR 批处理子进程提前退出")
                    result = json.loads(response)
                    if not isinstance(result, dict):
                        raise RuntimeError("日期 OCR 批处理返回无效结果")
                except queue.Empty:
                    self._stop_date_process(process)
                    self._close_date_process_pipes(process, reader_thread)
                    process = None
                    record_batch_wall()
                    retry = self.rewrite_with_metadata(path, target_year)
                    if str(retry.get("status", "")).upper() not in {
                        "CHANGED", "UNCHANGED", "CHANGED_WITH_AUTO_SKIPS", "AUTO_SKIPPED_UNCERTAIN_IMAGES",
                    }:
                        retry["reason"] = f"批处理单件超时；单件回退失败：{retry.get('reason', retry.get('status'))}"
                    retry.setdefault("performance", {})["batch_timeout_fallback"] = True
                    yield retry
                    for remaining in paths[index + 1:]:
                        yield self.rewrite_with_metadata(remaining, target_year)
                    return
                except Exception as exc:
                    self._stop_date_process(process)
                    self._close_date_process_pipes(process, reader_thread)
                    process = None
                    record_batch_wall()
                    retry = self.rewrite_with_metadata(path, target_year)
                    if str(retry.get("status", "")).upper() not in {
                        "CHANGED", "UNCHANGED", "CHANGED_WITH_AUTO_SKIPS", "AUTO_SKIPPED_UNCERTAIN_IMAGES",
                    }:
                        retry["reason"] = f"批处理单件中断；单件回退失败：{retry.get('reason', retry.get('status'))}"
                    retry.setdefault("performance", {})["batch_interruption_fallback"] = True
                    yield retry
                    for remaining in paths[index + 1:]:
                        yield self.rewrite_with_metadata(remaining, target_year)
                    return
                yield result

            try:
                process.stdin.write(json.dumps({"op": "close"}) + "\n")
                process.stdin.flush()
                process.stdin.close()
                process.wait(timeout=10)
                self._close_date_process_pipes(process, reader_thread)
            except Exception:
                self._stop_date_process(process)
                self._close_date_process_pipes(process, reader_thread)
            process = None
            record_batch_wall()
        except Exception:
            if process is not None:
                self._stop_date_process(process)
                self._close_date_process_pipes(process, reader_thread)
                process = None
            record_batch_wall()
            # Startup failure occurs before any request; use the compatible isolated path per DOCX.
            for path in paths:
                yield self.rewrite_with_metadata(path, target_year)
        finally:
            if process is not None:
                self._stop_date_process(process)
                self._close_date_process_pipes(process, reader_thread)
            record_batch_wall()

    @staticmethod
    def _stop_date_process(process) -> None:
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=2)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)

    @staticmethod
    def _close_date_process_pipes(process, reader_thread=None) -> None:
        if reader_thread is not None:
            reader_thread.join(timeout=2)
        for stream in (process.stdin, process.stdout, process.stderr):
            if stream is not None and not stream.closed:
                stream.close()


def read_docx_paragraphs(path: str | Path) -> list[tuple[str, str]]:
    texts: list[tuple[str, str]] = []
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
        for index, paragraph in enumerate(root.iter(f"{{{WORD_NS}}}p")):
            value = "".join(node.text or "" for node in paragraph.iter(f"{{{WORD_NS}}}t"))
            if value.strip():
                texts.append((f"paragraph:{index + 1}", value))
    return texts


def _is_ai_hollow_rectangle(width: float, height: float) -> bool:
    if width <= 0 or height <= 0:
        return False
    return width >= 5_000_000 and height <= 220_000 and width / height >= 24


def find_ai_hollow_lines(path: str | Path) -> list[dict[str, Any]]:
    """Find long, thin outlined rectangles that resemble generated AI divider artifacts."""
    findings: list[dict[str, Any]] = []
    namespaces = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "w": WORD_NS,
    }
    with zipfile.ZipFile(path) as archive:
        for part_name in archive.namelist():
            if not re.fullmatch(r"word/(document|header[0-9]*|footer[0-9]*|footnotes|endnotes)\.xml", part_name, re.IGNORECASE):
                continue
            try:
                document = etree.fromstring(archive.read(part_name))
            except etree.XMLSyntaxError:
                continue
            for shape in document.xpath(".//wps:sp", namespaces=namespaces):
                geometry = shape.find(".//{%s}prstGeom" % namespaces["a"])
                if geometry is None or geometry.get("prst") not in {"rect", "roundRect"}:
                    continue
                extent = shape.find(".//{%s}xfrm/{%s}ext" % (namespaces["a"], namespaces["a"]))
                if extent is None:
                    continue
                try:
                    width, height = float(extent.get("cx", "0")), float(extent.get("cy", "0"))
                except (TypeError, ValueError):
                    continue
                if not _is_ai_hollow_rectangle(width, height):
                    continue
                shape_properties = shape.find("{%s}spPr" % namespaces["wps"])
                has_outline = shape_properties is not None and shape_properties.find(".//{%s}ln" % namespaces["a"]) is not None
                no_fill = shape_properties is not None and shape_properties.find(".//{%s}noFill" % namespaces["a"]) is not None
                fill = shape_properties.find(".//{%s}solidFill" % namespaces["a"]) if shape_properties is not None else None
                def is_white_fill(color: Any) -> bool:
                    if color.tag.endswith("srgbClr"):
                        return color.get("val", "").upper() == "FFFFFF"
                    if color.tag.endswith("sysClr"):
                        return color.get("lastClr", "").upper() == "FFFFFF"
                    return color.get("val", "").casefold() in {"bg1", "lt1"}

                white_fill = fill is not None and any(is_white_fill(color) for color in fill)
                if not has_outline or not (no_fill or white_fill):
                    continue
                paragraph = shape
                while paragraph is not None and paragraph.tag != "{%s}p" % namespaces["w"]:
                    paragraph = paragraph.getparent()
                context = "".join(paragraph.xpath(".//w:t/text()", namespaces=namespaces)).strip() if paragraph is not None else ""
                findings.append({"part": part_name, "width_emu": int(width), "height_emu": int(height), "context": context})
    return findings


def _docx_parts(path: Path) -> tuple[list[tuple[str, bytes]], list[tuple[str, str]]]:
    images: list[tuple[str, bytes]] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.lower().startswith("word/media/") and Path(name).suffix.lower() in IMAGE_SUFFIXES:
                images.append((name, archive.read(name)))
    texts = read_docx_paragraphs(path)
    return images, texts


def _docx_image_names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        return [
            name for name in archive.namelist()
            if name.lower().startswith("word/media/") and Path(name).suffix.lower() in IMAGE_SUFFIXES
        ]


def _docx_images(path: Path) -> list[tuple[str, bytes]]:
    with zipfile.ZipFile(path) as archive:
        return [
            (name, archive.read(name))
            for name in archive.namelist()
            if name.lower().startswith("word/media/") and Path(name).suffix.lower() in IMAGE_SUFFIXES
        ]


def _qualification_paragraphs(value: Any) -> list[tuple[str, str]]:
    if not isinstance(value, list):
        return []
    converted: list[tuple[str, str]] = []
    for index, item in enumerate(value, 1):
        if isinstance(item, dict):
            paragraph_id = str(item.get("paragraph_id", f"paragraph:{index}"))
            text = str(item.get("text", ""))
        elif isinstance(item, (tuple, list)) and len(item) >= 2:
            paragraph_id, text = str(item[0]), str(item[1])
        else:
            continue
        if text.strip():
            converted.append((paragraph_id, text))
    return converted


def _semantic_status(value: Any) -> str:
    normalized = str(value or "").strip().casefold()
    if normalized in {"match", "matched", "符合"}:
        return "MATCH"
    if normalized in {"mismatch", "not_match", "不匹配", "错配", "review_topic_mismatch"}:
        return "MISMATCH"
    return "NOT_EVALUATED"


def _semantic_relevance(source_item: dict[str, Any] | None) -> tuple[dict[str, Any], str, str, str]:
    if source_item is None:
        result: dict[str, Any] = {}
        metadata: dict[str, Any] = {}
        title = ""
        diagnostic = "TITLE_BODY_RESULT_MISSING"
    else:
        raw_result = source_item.get("result", source_item.get("semantic_result", {}))
        result = raw_result if isinstance(raw_result, dict) else {}
        raw_metadata = source_item.get("metadata", {})
        metadata = raw_metadata if isinstance(raw_metadata, dict) else {}
        title = str(source_item.get("title", metadata.get("raw_title", "")) or "").strip()
        diagnostic = str(
            result.get("diagnostic_code") or metadata.get("reason")
            or result.get("reason") or ""
        )
        if not result and source_item.get("semantic_status"):
            result = {"status": source_item.get("semantic_status"), "reason": source_item.get("reason", "")}
    status = _semantic_status(result.get("status", source_item.get("semantic_status") if source_item else None))
    reason = str(result.get("reason") or diagnostic or ("标题正文未完成评估" if status == "NOT_EVALUATED" else ""))
    relevance = {
        "status": "review_topic_mismatch" if status == "MISMATCH" else ("match" if status == "MATCH" else "not_evaluated"),
        "similarity_score": None,
        "reason": reason or ("标题正文未评估" if status == "NOT_EVALUATED" else ""),
        "semantic_status": status,
        "evidence": result.get("evidence", []),
        "diagnostic_code": result.get("diagnostic_code") or metadata.get("diagnostic_code"),
    }
    return relevance, status, title, diagnostic


def _qualification_action(mode: str, semantic_status: str, skip_requested: bool, not_evaluated_action: str) -> str:
    if mode == "shadow":
        return "CONTINUE" if semantic_status == "MATCH" else "CONTINUE_WITH_REVIEW"
    if semantic_status == "MISMATCH":
        return "SKIP_MISMATCH" if skip_requested else "CONTINUE_WITH_REVIEW"
    if semantic_status == "NOT_EVALUATED":
        return "HOLD_NOT_EVALUATED" if not_evaluated_action == "hold_changes" else "CONTINUE_WITH_REVIEW"
    return "CONTINUE"


def _default_relevance(text: str, title: str, context: str) -> dict[str, Any]:
    if not text.strip():
        return {"status": "unknown", "confidence": 0.0, "reason": "OCR 无有效文字"}
    tokens = set(re.findall(r"[A-Za-z0-9]{2,}|[\u4e00-\u9fff]{2,}", title + " " + context))
    matched = [token for token in tokens if token.casefold() in text.casefold()]
    if matched:
        return {"status": "relevant", "confidence": 0.9, "reason": f"匹配主题词：{matched[0]}"}
    return {"status": "unknown", "confidence": 0.4, "reason": "确定性基线证据不足"}


def _semantic_tokens(value: str) -> set[str]:
    tokens = {item.casefold() for item in re.findall(r"[A-Za-z][A-Za-z0-9_-]{1,}", value)}
    for run in re.findall(r"[\u4e00-\u9fff]{2,}", value):
        tokens.update(run[index:index + 2] for index in range(len(run) - 1))
    return tokens - GENERIC_TERMS


def _semantic_score(text: str, reference: str) -> float:
    observed = _semantic_tokens(text)
    expected = _semantic_tokens(reference)
    if not observed or not expected:
        return 0.0
    return len(observed & expected) / math.sqrt(len(observed) * len(expected))


def _body_similarity_score(title: str, body: str) -> float:
    """Return an explainable 0-100 title/body topic overlap score."""
    lowered_title = title.casefold()
    lowered_body = body.casefold()
    title_topics = {
        topic for topic, terms in BODY_TOPIC_TERMS.items()
        if any(term.casefold() in lowered_title for term in terms)
    }
    counts = {
        topic: sum(lowered_body.count(term.casefold()) for term in terms)
        for topic, terms in BODY_TOPIC_TERMS.items()
    }
    total = sum(counts.values())
    if title_topics and total:
        return round(100 * sum(counts[topic] for topic in title_topics) / total, 1)
    return round(100 * _semantic_score(title, body), 1)


def _batch_relevance(text: str, current: int, references: list[str]) -> dict[str, Any]:
    if not text.strip():
        return {"status": "unknown", "confidence": 0.0, "reason": "OCR 无有效文字"}
    own = _semantic_score(text, references[current])
    competitors = [(_semantic_score(text, reference), index) for index, reference in enumerate(references) if index != current]
    best_other, best_index = max(competitors, default=(0.0, -1))
    if own >= 0.08:
        return {"status": "relevant", "confidence": min(0.99, 0.85 + own), "reason": f"与本项目语义特征匹配（{own:.3f}）"}
    if best_other >= 0.08 and best_other - own >= 0.05:
        return {
            "status": "irrelevant", "confidence": min(0.99, 0.86 + best_other),
            "reason": f"与其他项目 #{best_index + 1} 的语义特征更接近（本项目 {own:.3f}，其他 {best_other:.3f}）",
        }
    return {"status": "unknown", "confidence": 0.4, "reason": "语义证据不足，不作不相关判定"}


def assess_body_relevance(title: str, paragraphs: list[tuple[str, str]]) -> dict[str, Any]:
    """Flag a strong topic conflict for review; never use this to block writes.

    Full title repetitions (including running headers) are removed before counting.
    The result is evidence for a reviewer, not a semantic model's verdict.
    """
    body = "\n".join(value for _, value in paragraphs)
    if title:
        body = re.sub(re.escape(title), "", body, flags=re.IGNORECASE)
    similarity_score = _body_similarity_score(title, body)
    if len(body.strip()) < 80:
        return {"status": "unknown", "reason": "正文可分析文字不足", "similarity_score": similarity_score, "title_terms": {}, "other_topics": {}}
    lowered_title = title.casefold()
    lowered_body = body.casefold()
    later_body = "\n".join(value for _, value in paragraphs[10:])
    if title:
        later_body = re.sub(re.escape(title), "", later_body, flags=re.IGNORECASE)
    later_body = later_body.casefold()
    title_terms = {
        term: lowered_body.count(term.casefold())
        for terms in BODY_TOPIC_TERMS.values() for term in terms
        if term.casefold() in lowered_title
    }
    if not title_terms:
        return {"status": "unknown", "reason": "标题缺少可核对的领域词", "similarity_score": similarity_score, "title_terms": {}, "other_topics": {}}
    title_topics = {
        topic for topic, terms in BODY_TOPIC_TERMS.items()
        if any(term.casefold() in lowered_title for term in terms)
    }
    other_topics = {
        topic: {term: lowered_body.count(term.casefold()) for term in terms if lowered_body.count(term.casefold())}
        for topic, terms in BODY_TOPIC_TERMS.items() if topic not in title_topics
    }
    other_topics = {topic: counts for topic, counts in other_topics.items() if sum(counts.values()) >= 5}
    absent = [term for term, count in title_terms.items() if count == 0]
    strongest_other = max((sum(counts.values()) for counts in other_topics.values()), default=0)
    late_title_hits = sum(later_body.count(term.casefold()) for term in title_terms)
    generic_operations = {term: lowered_body.count(term) for term in BODY_GENERIC_OPERATIONS}
    generic_operation_total = sum(generic_operations.values())
    if other_topics and (
        len(absent) >= 2
        or (sum(title_terms.values()) <= 3 and strongest_other >= 8)
        or (strongest_other >= 8 and strongest_other > sum(title_terms.values()) and late_title_hits <= 1)
    ):
        reason = (
            f"标题领域词缺席：{'、'.join(absent)}；正文反复出现其他主题词"
            if absent else "标题领域词集中在开头，后文反复出现其他主题词"
        )
        return {
            "status": "review_topic_mismatch",
            "reason": reason,
            "similarity_score": similarity_score,
            "title_terms": title_terms,
            "other_topics": other_topics,
        }
    if (len(absent) == len(title_terms) and generic_operation_total >= 20
            and all(generic_operations.values())):
        return {
            "status": "review_topic_mismatch",
            "reason": "标题专有领域词在正文缺席，正文主要重复通用增删查导操作",
            "similarity_score": similarity_score,
            "title_terms": title_terms, "other_topics": {"通用操作": generic_operations},
        }
    return {
        "status": "unknown", "reason": "未发现足够强的标题与正文主题冲突证据",
        "similarity_score": similarity_score,
        "title_terms": title_terms, "other_topics": other_topics,
    }


def _normalise_relevance(value: Any, threshold: float) -> dict[str, Any]:
    if isinstance(value, str):
        value = {"status": value}
    if not isinstance(value, dict):
        value = {}
    status = str(value.get("status", value.get("decision", "unknown"))).lower()
    confidence = float(value.get("confidence", 0.0))
    if status not in {"relevant", "irrelevant", "unknown"}:
        status = "unknown"
    if status == "irrelevant" and confidence < threshold:
        status = "unknown"
    return {"status": status, "confidence": confidence, "reason": str(value.get("reason", ""))}


def _rewrite_document_years(path: Path, classifier: Callable[..., Any] | None, target_year: int) -> list[dict[str, Any]]:
    changes: list[dict[str, Any]] = []
    with zipfile.ZipFile(path, "r") as source:
        document_xml = source.read("word/document.xml")
    root = etree.fromstring(document_xml)
    changed = False
    for node in root.iter(f"{{{WORD_NS}}}t"):
        if not node.text or not YEAR_RE.search(node.text):
            continue
        replaced, decisions = replace_update_years(node.text, classifier, target_year)
        for decision in decisions:
            item_changed = decision["replacement"] != decision["source"]
            changes.append({**decision, "changed": item_changed})
            changed = changed or item_changed
        node.text = replaced
    if not changed:
        return changes

    temporary = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename.lower() == "word/document.xml":
                data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
            target.writestr(item, data)
    os.replace(temporary, path)
    return changes


def _progress(progress: Any, phase: str, current: int, name: str) -> None:
    if progress is None:
        return
    if callable(progress):
        progress(phase, current, name)
    elif hasattr(progress, "update"):
        progress.update(phase, current, name)


def run_manual_content_stage(
    input_root: str | Path,
    output_root: str | Path,
    report_path: str | Path,
    progress: Any = None,
    *,
    ocr: Callable[..., Any] | None = None,
    relevance_classifier: Callable[..., Any] | None = None,
    vision_classifier: Callable[..., Any] | None = None,
    vision_max_calls_per_project: int = VISION_MAX_CALLS_PER_PROJECT,
    year_classifier: Callable[..., Any] | None = None,
    date_rewriter: Callable[..., Any] | None = None,
    confidence_threshold: float = CONFIDENCE_THRESHOLD,
    target_year: int = TARGET_YEAR,
    apply: bool = True,
    skip_topic_mismatch: bool = True,
    check_body_relevance: bool = True,
    title_body_inputs_by_relative_path: dict[str, dict[str, Any]] | None = None,
    title_body_mode: str = "shadow",
    skip_mismatch: bool | None = None,
    not_evaluated_action: str = "hold_changes",
    unknown_action: str | None = None,
    check_image_relevance: bool = True,
    check_image_vision: bool = True,
    rewrite_body_years: bool = True,
    rewrite_image_dates: bool = True,
    check_ai_hollow_lines: bool = False,
) -> dict[str, Any]:
    """Preflight every manual, then copy and mutate only after batch gating completes."""
    perf_trace = os.getenv("EASYSOFTWARE_PERF_TRACE") == "1"
    run_started = time.perf_counter() if perf_trace else None
    preparation_started = time.perf_counter() if perf_trace else None
    source_root = Path(input_root).resolve()
    destination_root = Path(output_root).resolve()
    report_file = Path(report_path).resolve()
    manuals = sorted(source_root.rglob("*说明.docx"), key=lambda path: str(path).casefold())
    mode = str(title_body_mode or "shadow").strip().casefold()
    if mode not in {"off", "shadow", "enforced"}:
        raise ValueError("title_body_mode 必须是 off、shadow 或 enforced")
    if unknown_action is not None:
        not_evaluated_action = unknown_action
    if not_evaluated_action not in {"hold_changes", "continue_with_review"}:
        raise ValueError("not_evaluated_action 必须是 hold_changes 或 continue_with_review")
    use_title_body_inputs = title_body_inputs_by_relative_path is not None and mode != "off"
    report_mode = mode if title_body_inputs_by_relative_path is not None else "legacy"
    skip_requested = skip_topic_mismatch if skip_mismatch is None else bool(skip_mismatch)
    skip_effective = (
        (mode == "enforced" and skip_requested) if use_title_body_inputs
        else (skip_topic_mismatch if not use_title_body_inputs else False)
    )
    qualification_items = {
        str(key).replace("\\", "/"): item
        for key, item in (title_body_inputs_by_relative_path or {}).items()
        if isinstance(item, dict)
    }
    ocr_adapter = ocr or DefaultOcrAdapter(Path(__file__).resolve().parent)
    relevance = relevance_classifier or _default_relevance
    vision_adapter = vision_classifier or OpenAICompatibleVisionAdapter()
    vision_backend = getattr(vision_adapter, "backend", "injected")
    vision_disabled_reason = ""
    date_adapter = date_rewriter or DefaultDateRewriter(Path(__file__).resolve().parent)
    date_adapter_performance = getattr(date_adapter, "_performance", None)
    if isinstance(date_adapter_performance, dict):
        date_adapter_performance.clear()
    projects: list[dict[str, Any]] = []
    prepared: list[dict[str, Any]] = []
    for manual in manuals:
        relative_path = manual.relative_to(source_root).as_posix()
        title_body_item = qualification_items.get(relative_path) if use_title_body_inputs else None
        if use_title_body_inputs:
            body_relevance, semantic_status, title, diagnostic_code = _semantic_relevance(title_body_item)
            paragraphs = _qualification_paragraphs(title_body_item.get("paragraphs")) if title_body_item else []
            if title_body_item is None:
                body_relevance["reason"] = "材料清单中缺少该说明书的资格结果"
                body_relevance["diagnostic_code"] = "TITLE_BODY_RESULT_MISSING"
                diagnostic_code = "TITLE_BODY_RESULT_MISSING"
            if not paragraphs:
                body_relevance["semantic_status"] = semantic_status = "NOT_EVALUATED"
                body_relevance["status"] = "not_evaluated"
                body_relevance["reason"] = body_relevance.get("reason") or "资格结果未携带已抽取的说明书正文"
                diagnostic_code = diagnostic_code or "TITLE_BODY_PARAGRAPHS_MISSING"
            action = _qualification_action(mode, semantic_status, skip_requested, not_evaluated_action)
            if title_body_item and not title:
                title = ""
        else:
            paragraphs = read_docx_paragraphs(manual)
            title = manual.stem.removesuffix("说明")
            body_relevance = None
            semantic_status = ""
            diagnostic_code = ""
            action = "LEGACY"
        context = "\n".join(value for _, value in paragraphs[:3])
        prepared.append({
            "manual": manual,
            "relative_path": relative_path,
            "title_body_item": title_body_item,
            "paragraphs": paragraphs,
            "title": title,
            "context": context,
            "body_relevance": body_relevance,
            "semantic_status": semantic_status,
            "diagnostic_code": diagnostic_code,
            "qualification_action": action,
        })
    references = [f"{item['title']}\n{item['context']}" for item in prepared]
    if not use_title_body_inputs:
        for item in prepared:
            body_relevance = (
                assess_body_relevance(item["title"], item["paragraphs"])
                if check_body_relevance else
                {"status": "not_requested", "similarity_score": None, "reason": "本次未选择标题与正文主题检测"}
            )
            topic_mismatch = body_relevance.get("status") == "review_topic_mismatch"
            item["body_relevance"] = body_relevance
            item["semantic_status"] = "MISMATCH" if topic_mismatch else "NOT_EVALUATED"
            item["qualification_action"] = "SKIP_MISMATCH" if (skip_topic_mismatch and topic_mismatch) else "LEGACY"

    preparation_seconds = time.perf_counter() - preparation_started if perf_trace else None
    preflight_started = time.perf_counter() if perf_trace else None
    for document_index, prepared_item in enumerate(prepared, 1):
        manual = prepared_item["manual"]
        relative_path = prepared_item["relative_path"]
        paragraphs = prepared_item["paragraphs"]
        title = prepared_item["title"]
        context = prepared_item["context"]
        if use_title_body_inputs:
            body_relevance = prepared_item["body_relevance"]
            semantic_status = prepared_item["semantic_status"]
            qualification_action = prepared_item["qualification_action"]
            topic_mismatch = semantic_status == "MISMATCH"
            qualification_status = "TOPIC_MISMATCH" if topic_mismatch else ("QUALIFIED" if semantic_status == "MATCH" else "NOT_EVALUATED")
            pending_review = qualification_action == "CONTINUE_WITH_REVIEW" or semantic_status == "NOT_EVALUATED"
            changes_held = qualification_action in {"SKIP_MISMATCH", "HOLD_NOT_EVALUATED"}
        else:
            body_relevance = prepared_item["body_relevance"]
            semantic_status = prepared_item["semantic_status"]
            topic_mismatch = body_relevance.get("status") == "review_topic_mismatch"
            qualification_status = "TOPIC_MISMATCH" if topic_mismatch else "QUALIFIED"
            qualification_action = prepared_item["qualification_action"]
            pending_review = False
            changes_held = bool(skip_topic_mismatch and topic_mismatch)
        image_names = _docx_image_names(manual)
        hollow_line_findings = find_ai_hollow_lines(manual) if check_ai_hollow_lines else []
        manual_title = title or Path(relative_path).stem.removesuffix("说明")
        base_project = {
            "source": str(manual),
            "relative_path": relative_path,
            "title": title,
            "txt_path": prepared_item["title_body_item"].get("txt_path") if prepared_item["title_body_item"] else None,
            "project_id": prepared_item["title_body_item"].get("project_id") if prepared_item["title_body_item"] else None,
            "semantic_input_hash": prepared_item["title_body_item"].get("semantic_input_hash") if prepared_item["title_body_item"] else None,
            "title_body_result": prepared_item["title_body_item"].get("result") if prepared_item["title_body_item"] else None,
            "body_relevance": body_relevance,
            "semantic_status": semantic_status if use_title_body_inputs else None,
            "diagnostic_code": prepared_item["diagnostic_code"] or None,
            "qualification_mode": report_mode,
            "qualification_status": qualification_status,
            "qualification_action": qualification_action,
            "qualification_pending_review": pending_review,
            "changes_held": changes_held,
            "skip_requested": skip_requested,
            "skip_effective": qualification_action == "SKIP_MISMATCH" if use_title_body_inputs else (skip_topic_mismatch and topic_mismatch),
            "image_count": len(image_names),
            "ai_hollow_line_findings": hollow_line_findings,
            "ai_hollow_line_alert": f"需要将该说明书（{manual_title}）发给其他人。" if hollow_line_findings else None,
        }
        if changes_held:
            hold_status = "SKIPPED_TOPIC_MISMATCH" if qualification_action == "SKIP_MISMATCH" else "HELD_QUALIFICATION_NOT_EVALUATED"
            projects.append({
                **base_project,
                "sample_indexes": [],
                "samples": [],
                "irrelevant_count": 0,
                "decisive_count": 0,
                "visual_backend": vision_backend,
                "visual_call_count": 0,
                "visual_decision_count": 0,
                "visual_failure_count": 0,
                "visual_total_tokens": 0,
                "visual_candidate_count": 0,
                "visual_selected_indexes": [],
                "blocked": False,
                "status": hold_status,
            })
            _progress(progress, "跳过标题正文资格未通过项目", document_index, manual.name)
            continue
        images = _docx_images(manual) if (check_image_relevance or check_image_vision) else []
        samples = []
        visual_calls = 0
        visual_decisions = 0
        visual_failures = 0
        visual_tokens = 0
        sample_indexes = equidistant_sample_indexes(len(images)) if (check_image_relevance or check_image_vision) else []
        sample_inputs = [(images[index][1], Path(images[index][0]).suffix) for index in sample_indexes]
        if hasattr(ocr_adapter, "batch"):
            raw_results = ocr_adapter.batch(sample_inputs)
        else:
            raw_results = [_call_adapter(ocr_adapter, data, suffix) for data, suffix in sample_inputs]
        for image_index, raw_ocr in zip(sample_indexes, raw_results):
            image_name, _image_bytes = images[image_index]
            if isinstance(raw_ocr, str):
                raw_ocr = {"text": raw_ocr}
            raw_ocr = raw_ocr if isinstance(raw_ocr, dict) else {}
            text = str(raw_ocr.get("text", "")).strip()
            if not text:
                judgement = {"status": "unknown", "confidence": 0.0, "reason": "OCR 无有效文字"}
            else:
                raw_judgement = (
                    _call_adapter(relevance, text, title, context)
                    if relevance_classifier is not None
                    else _batch_relevance(text, document_index - 1, references)
                )
                judgement = _normalise_relevance(raw_judgement, confidence_threshold)
            samples.append({
                "index": image_index,
                "image": image_name,
                "ocr_text": text,
                "ocr_backend": raw_ocr.get("backend", getattr(ocr_adapter, "backend", "injected")),
                **judgement,
                "visual_review": {"status": "not_requested", "backend": vision_backend},
            })
        visual_candidates = [position for position, sample in enumerate(samples) if sample["status"] in {"unknown", "irrelevant"}] if check_image_vision else []
        vision_limit = max(0, min(int(vision_max_calls_per_project), len(visual_candidates)))
        selected_positions = [
            visual_candidates[index] for index in equidistant_sample_indexes(len(visual_candidates), ratio=0, minimum=vision_limit)
        ] if vision_limit else []
        for position in selected_positions:
            sample = samples[position]
            image_name, image_bytes = images[sample["index"]]
            suffix = Path(image_name).suffix
            if vision_disabled_reason:
                visual_review = {"status": "unknown", "backend": vision_backend, "error": vision_disabled_reason, "reason": "本批视觉接口此前失败，跳过后续调用"}
            elif suffix.lower() in IMAGE_SUFFIXES:
                visual_calls += 1
                raw_visual = _call_adapter(vision_adapter, image_bytes, suffix, title, context, sample["ocr_text"])
                visual_review = raw_visual if isinstance(raw_visual, dict) else {"status": "unknown", "error": "视觉分类器返回格式无效"}
                visual_review.setdefault("backend", vision_backend)
                if visual_review.get("error"):
                    visual_failures += 1
                    if visual_review.get("error_type") in {"URLError", "TimeoutError", "ConnectionRefusedError"}:
                        vision_disabled_reason = str(visual_review["error"])
                elif visual_review.get("status") in {"relevant", "irrelevant"}:
                    visual_decisions += 1
                usage = visual_review.get("usage", {})
                if isinstance(usage, dict):
                    visual_tokens += int(usage.get("total_tokens", 0) or 0)
            else:
                visual_review = {"status": "unknown", "backend": vision_backend, "reason": "图片格式或文件签名不支持视觉接口"}
            sample["visual_review"] = visual_review
        irrelevant = sum(item["status"] == "irrelevant" for item in samples)
        decisive = sum(item["status"] in {"relevant", "irrelevant"} for item in samples)
        blocked = bool(check_image_relevance and decisive and irrelevant / decisive > 0.5)
        if use_title_body_inputs and prepared_item["qualification_action"] == "CONTINUE_WITH_REVIEW":
            status = "BLOCKED_IRRELEVANT_WITH_REVIEW" if blocked else "PREFLIGHT_PASSED_WITH_REVIEW"
        else:
            status = "BLOCKED_IRRELEVANT" if blocked else "PREFLIGHT_PASSED"
        projects.append({
            **base_project,
            "sample_indexes": [item["index"] for item in samples],
            "samples": samples,
            "irrelevant_count": irrelevant,
            "decisive_count": decisive,
            "visual_backend": vision_backend,
            "visual_call_count": visual_calls,
            "visual_decision_count": visual_decisions,
            "visual_failure_count": visual_failures,
            "visual_total_tokens": visual_tokens,
            "visual_candidate_count": len(visual_candidates),
            "visual_selected_indexes": [samples[position]["index"] for position in selected_positions],
            "blocked": blocked,
            "status": status,
        })
        del images
        _progress(progress, "说明相关性预检", document_index, manual.name)

    eligible_for_image_gate = [
        project for project in projects
        if not project.get("changes_held")
    ]
    blocked_count = sum(project["blocked"] for project in eligible_for_image_gate)
    batch_fused = bool(eligible_for_image_gate and blocked_count / len(eligible_for_image_gate) > 0.2)
    preflight_seconds = time.perf_counter() - preflight_started if perf_trace else None

    # The first filesystem mutation intentionally happens only after every document is preflighted.
    mutation_started = time.perf_counter() if perf_trace else None
    date_documents_attempted = 0
    batch_stage_by_source: dict[str, Path] = {}
    batch_iterator = None
    if apply:
        try:
            batch_iterator_factory = getattr(date_adapter, "iter_rewrite_many_with_metadata", None)
            if callable(batch_iterator_factory) and rewrite_image_dates and not batch_fused:
                batch_projects = [
                    project for project in projects
                    if not project.get("changes_held") and not project["blocked"]
                ]
                batch_stages = []
                for project in batch_projects:
                    source = Path(project["source"])
                    destination = destination_root / project["relative_path"]
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    if source != destination:
                        shutil.copy2(source, destination)
                    project["output"] = str(destination)
                    stage_fd, stage_name = tempfile.mkstemp(
                        prefix=f".{destination.stem}.date-year-stage-", suffix=".docx", dir=destination.parent
                    )
                    os.close(stage_fd)
                    stage_path = Path(stage_name)
                    batch_stage_by_source[str(source)] = stage_path
                    shutil.copy2(destination, stage_path)
                    batch_stages.append(stage_path)
                if batch_stages:
                    batch_iterator = iter(batch_iterator_factory(batch_stages, target_year))

            for document_index, project in enumerate(projects, 1):
                source = Path(project["source"])
                destination = destination_root / project["relative_path"]
                destination.parent.mkdir(parents=True, exist_ok=True)
                if project.get("changes_held"):
                    project["output"] = str(destination)
                    hold_status = "SKIPPED_TOPIC_MISMATCH" if project.get("qualification_action") == "SKIP_MISMATCH" else "HELD_QUALIFICATION_NOT_EVALUATED"
                    project["status"] = hold_status
                    project["date_rewrite"] = {"status": hold_status, "changed": False}
                    project["year_changes"] = []
                    project["year_update_skipped_reason"] = project.get("body_relevance", {}).get("reason") or "项目资格状态要求暂停修改"
                    _progress(progress, "跳过标题正文资格未通过项目", document_index, source.name)
                    continue
                if source != destination and project.get("output") != str(destination):
                    shutil.copy2(source, destination)
                project["output"] = str(destination)
                year_update_skipped_reason = None
                if batch_fused and rewrite_image_dates:
                    project["date_rewrite"] = {"status": "SKIPPED_BATCH_FUSE", "changed": False}
                    project["status"] = "BATCH_FUSED"
                    year_update_skipped_reason = "批次图像日期阶段被熔断，无法确定与截图一致的目标年份"
                elif project["blocked"] and rewrite_image_dates:
                    project["date_rewrite"] = {"status": "SKIPPED_PROJECT_BLOCKED", "changed": False}
                    year_update_skipped_reason = "项目因截图相关性门控而跳过日期处理，无法确定与截图一致的目标年份"
                elif not rewrite_image_dates:
                    project["date_rewrite"] = {"status": "SKIPPED_NOT_SELECTED", "changed": False}
                else:
                    metadata_rewriter = getattr(date_adapter, "rewrite_with_metadata", None)
                    stage_path = batch_stage_by_source.get(str(source))
                    if stage_path is None:
                        stage_fd, stage_name = tempfile.mkstemp(
                            prefix=f".{destination.stem}.date-year-stage-", suffix=".docx", dir=destination.parent
                        )
                        os.close(stage_fd)
                        stage_path = Path(stage_name)
                        batch_stage_by_source[str(source)] = stage_path
                        shutil.copy2(destination, stage_path)
                    date_documents_attempted += 1
                    try:
                        if batch_iterator is not None and str(source) in batch_stage_by_source:
                            result = next(batch_iterator)
                        else:
                            result = metadata_rewriter(stage_path, target_year) if callable(metadata_rewriter) else _call_adapter(date_adapter, stage_path, target_year)
                    except StopIteration:
                        result = {"status": "FAILED_BATCH_INTERRUPTED", "changed": False,
                                  "reason": "批处理未返回该 DOCX 的结果"}
                    except Exception as exc:
                        result = {"status": "FAILED", "changed": False, "reason": str(exc)}
                    if not isinstance(result, dict) or not result:
                        result = {"status": "FAILED_INVALID_RESULT", "changed": False, "reason": "date_rewriter 未返回有效结果字典"}
                    result.setdefault("status", "CHANGED" if result.get("changed") else "UNCHANGED")
                    result.setdefault("changed", False)
                    successful_date_statuses = {
                        "CHANGED", "UNCHANGED", "CHANGED_WITH_AUTO_SKIPS", "AUTO_SKIPPED_UNCERTAIN_IMAGES",
                    }
                    date_status = str(result["status"]).upper()
                    if date_status in successful_date_statuses:
                        stage_path.with_suffix(stage_path.suffix + ".date-year.tmp").unlink(missing_ok=True)
                        os.replace(stage_path, destination)
                        batch_stage_by_source.pop(str(source), None)
                    else:
                        stage_path.unlink(missing_ok=True)
                        stage_path.with_suffix(stage_path.suffix + ".date-year.tmp").unlink(missing_ok=True)
                        batch_stage_by_source.pop(str(source), None)
                    project["date_rewrite"] = result
                    project["effective_target_year"] = int(result.get("effective_target_year", target_year))
                    project["trigger_date"] = result.get("trigger_date")
                    project["rollback_reason"] = result.get("rollback_reason")
                    if date_status == "UNAVAILABLE":
                        project["status"] = "DATE_REWRITER_UNAVAILABLE"
                    if date_status not in successful_date_statuses:
                        detail = result.get("reason") or result.get("error") or date_status
                        year_update_skipped_reason = f"截图日期处理未成功（{detail}），跳过正文年份改写以避免与截图不一致"
                effective_target_year = int(project.get("effective_target_year", target_year))
                project.setdefault("effective_target_year", effective_target_year)
                project.setdefault("trigger_date", None)
                project.setdefault("rollback_reason", None)
                if year_update_skipped_reason:
                    project["year_changes"] = []
                    project["year_update_skipped_reason"] = year_update_skipped_reason
                elif rewrite_body_years:
                    body_target_year = target_year if not rewrite_image_dates else effective_target_year
                    project["year_changes"] = _rewrite_document_years(destination, year_classifier, body_target_year)
                else:
                    project["year_changes"] = []
                _progress(progress, "说明内容处理", document_index, source.name)

        finally:
            if batch_iterator is not None and hasattr(batch_iterator, "close"):
                batch_iterator.close()
            for remaining_stage in batch_stage_by_source.values():
                remaining_stage.unlink(missing_ok=True)
                remaining_stage.with_suffix(remaining_stage.suffix + ".date-year.tmp").unlink(missing_ok=True)
            batch_stage_by_source.clear()

    for project in projects:
        project.setdefault("effective_target_year", target_year)
        project.setdefault("trigger_date", None)
        project.setdefault("rollback_reason", None)
        if not apply and rewrite_image_dates:
            project.setdefault("date_rewrite", {
                "status": "PREVIEW_NOT_SCANNED",
                "changed": False,
                "reason": "预览模式不扫描全部截图日期；有效年份与触发日期需在实际应用时确定",
            })
            project.setdefault("year_changes", [])
            project.setdefault("year_update_skipped_reason", "预览模式未扫描截图日期，无法确定与截图一致的正文目标年份")

    mutation_seconds = time.perf_counter() - mutation_started if perf_trace else None

    disqualified = [project for project in projects if project["qualification_status"] == "TOPIC_MISMATCH"]
    qualification_reviews = [project for project in projects if project.get("qualification_pending_review")]
    report = {
        "schema_version": 1,
        "stage": "manual_content",
        "created_at": datetime.now().astimezone().isoformat(),
        "input_root": str(source_root),
        "output_root": str(destination_root),
        "target_year": target_year,
        "applied": apply,
        "confidence_threshold": confidence_threshold,
        "manual_count": len(projects),
        "blocked_project_count": blocked_count,
        "batch_fused": batch_fused,
        "visual_backend": vision_backend,
        "visual_endpoint": _safe_endpoint(getattr(vision_adapter, "base_url", "")) if getattr(vision_adapter, "base_url", None) else None,
        "visual_model": getattr(vision_adapter, "model", None),
        "visual_max_image_bytes": getattr(vision_adapter, "max_image_bytes", None),
        "visual_max_image_side": getattr(vision_adapter, "max_image_side", None),
        "visual_api_key_count": len(getattr(vision_adapter, "api_keys", [])),
        "visual_retry_count": getattr(vision_adapter, "retry_count", None),
        "skip_topic_mismatch": skip_effective if use_title_body_inputs else skip_topic_mismatch,
        "skip_requested": skip_requested,
        "skip_effective": skip_effective,
        "title_body_mode": report_mode,
        "not_evaluated_action": not_evaluated_action,
        "options": {
            "check_body_relevance": check_body_relevance,
            "title_body_mode": report_mode,
            "check_image_relevance": check_image_relevance,
            "check_image_vision": check_image_vision,
            "rewrite_body_years": rewrite_body_years,
            "rewrite_image_dates": rewrite_image_dates,
            "check_ai_hollow_lines": check_ai_hollow_lines,
        },
        "disqualified_project_count": len(disqualified),
        "skipped_project_count": sum(project.get("qualification_action") == "SKIP_MISMATCH" for project in projects),
        "held_not_evaluated_count": sum(project.get("qualification_action") == "HOLD_NOT_EVALUATED" for project in projects),
        "qualification_review_count": len(qualification_reviews),
        "disqualified_projects": [
            {
                "name": project.get("title") or Path(project["relative_path"]).stem.removesuffix("说明"),
                "relative_path": project["relative_path"],
                "semantic_status": project.get("semantic_status"),
                "similarity_score": project["body_relevance"].get("similarity_score"),
                "evidence": project["body_relevance"].get("evidence", []),
                "action": project.get("qualification_action"),
                "reason": project["body_relevance"].get("reason", ""),
            }
            for project in disqualified
        ],
        "qualification_review_projects": [
            {
                "name": project.get("title") or Path(project["relative_path"]).stem.removesuffix("说明"),
                "relative_path": project["relative_path"],
                "semantic_status": project.get("semantic_status"),
                "action": project.get("qualification_action"),
                "diagnostic_code": project.get("diagnostic_code"),
                "reason": project["body_relevance"].get("reason", ""),
            }
            for project in qualification_reviews
        ],
        "visual_max_calls_per_project": vision_max_calls_per_project,
        "visual_call_count": sum(project["visual_call_count"] for project in projects),
        "visual_decision_count": sum(project["visual_decision_count"] for project in projects),
        "visual_failure_count": sum(project["visual_failure_count"] for project in projects),
        "visual_total_tokens": sum(project["visual_total_tokens"] for project in projects),
        "visual_timeout_seconds": vision_adapter.timeout if isinstance(vision_adapter, LocalVisionAdapter) else None,
        "visual_error": vision_disabled_reason or None,
        "projects": projects,
    }
    if perf_trace:
        report["performance"] = {
            "preparation_seconds": round(preparation_seconds, 4),
            "preflight_seconds": round(preflight_seconds, 4),
            "mutation_seconds": round(mutation_seconds, 4),
            "date_rewriter": {
                "documents_attempted": date_documents_attempted,
                "subprocess_calls": date_adapter_performance.get("subprocess_calls") if isinstance(date_adapter_performance, dict) else None,
                "batch_count": date_adapter_performance.get("batch_count", 0) if isinstance(date_adapter_performance, dict) else 0,
                "wall_seconds": round(float(date_adapter_performance.get("wall_seconds", 0.0)), 4) if isinstance(date_adapter_performance, dict) else None,
                "batch_wall_seconds": round(float(date_adapter_performance.get("batch_wall_seconds", 0.0)), 4) if isinstance(date_adapter_performance, dict) else None,
                "engine_load_seconds": round(float(date_adapter_performance.get("engine_load_seconds", 0.0)), 4) if isinstance(date_adapter_performance, dict) else None,
            },
            "ocr_adapter": dict(getattr(ocr_adapter, "_performance", {})),
            "total_seconds": round(time.perf_counter() - run_started, 4),
        }
    report_file.parent.mkdir(parents=True, exist_ok=True)
    temporary_report = report_file.with_suffix(report_file.suffix + ".tmp")
    temporary_report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary_report, report_file)
    return report


__all__ = [
    "LocalVisionAdapter",
    "OpenAICompatibleVisionAdapter",
    "assess_body_relevance",
    "classify_year_context",
    "equidistant_sample_indexes",
    "read_docx_paragraphs",
    "replace_update_years",
    "run_manual_content_stage",
]

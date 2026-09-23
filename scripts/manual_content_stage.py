from __future__ import annotations

import inspect
import base64
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from lxml import etree


TARGET_YEAR = 2025
CONFIDENCE_THRESHOLD = 0.85
VISION_MODEL = "sensenova-6.8-flash-lite"
VISION_BASE_URL = "http://127.0.0.1:3458/v1"
VISION_MAX_CALLS_PER_PROJECT = 2
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


def classify_year_context(text: str, year: str | None = None) -> dict[str, Any]:
    """Conservative deterministic baseline for a single year occurrence."""
    value = year or (YEAR_RE.search(text).group(0) if YEAR_RE.search(text) else "")
    if not value or int(value) >= TARGET_YEAR:
        return {"decision": "KEEP", "confidence": 1.0, "reason": "目标年份、未来年份或未发现年份"}
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
        if int(match.group(0)) >= target_year:
            raw = {"decision": "KEEP", "confidence": 1.0, "reason": "目标年份或未来年份"}
        else:
            raw = _call_adapter(classifier, context, match.group(0)) if classifier else classify_year_context(context, match.group(0))
        decision = _normalise_year_decision(raw)
        replacement = str(target_year) if decision["decision"] == "UPDATE_TO_2025" else match.group(0)
        pieces.append(replacement)
        decisions.append({
            "source": match.group(0),
            "replacement": replacement,
            "context": context,
            **decision,
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
    if decision not in {"UPDATE_TO_2025", "KEEP", "REVIEW"}:
        decision = "REVIEW"
    confidence = float(value.get("confidence", 0.0))
    if decision == "UPDATE_TO_2025" and confidence < CONFIDENCE_THRESHOLD:
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
        self.tool_root = script_dir.parent.parent / "图片日期智能处理工具_用户版_20260903"
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


class LocalVisionAdapter:
    """Bounded OpenAI-compatible image reviewer; transport errors remain unknown."""

    backend = "local_vision"

    def __init__(self, base_url: str | None = None, api_key: str | None = None, timeout: float = 20.0):
        self.base_url = (base_url or os.getenv("EASYSOFTWARE_VISION_BASE_URL") or VISION_BASE_URL).rstrip("/")
        self.api_key = api_key or os.getenv("EASYSOFTWARE_VISION_API_KEY") or "proxy"
        self.timeout = timeout

    def __call__(self, image_bytes: bytes, suffix: str, title: str, context: str, ocr_text: str) -> dict[str, Any]:
        suffix = suffix.lower()
        mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}.get(suffix)
        if not mime:
            return {"status": "unknown", "confidence": 0.0, "reason": "图片格式不支持视觉接口", "backend": self.backend, "error": f"unsupported image format: {suffix}"}
        prompt = (
            "核对这张软件说明书截图与项目标题的主题是否一致。只依据图片中可见的界面、文字和图表，"
            "不要把标题或 OCR 文本当作图片证据。仅当图片明确属于另一领域时判 irrelevant；"
            "看不清、通用界面、证据不足时判 unknown。请输出 JSON 对象，字段 status "
            "(relevant/irrelevant/unknown)、confidence (0 到 1)、evidence (具体可见证据)。\n"
            f"项目标题：{title[:160]}\n正文开头：{context[:300]}\nOCR 提示：{ocr_text[:300]}"
        )
        payload = {
            "model": VISION_MODEL,
            "messages": [{"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{base64.b64encode(image_bytes).decode('ascii')}"}},
            ]}],
            "temperature": 0,
            "max_tokens": 1200,
        }
        request = urllib.request.Request(
            f"{self.base_url}/chat/completions",
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                reply = json.load(response)
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
            return {**result, "backend": self.backend, "model": VISION_MODEL, "usage": reply.get("usage", {})}
        except Exception as exc:
            return {"status": "unknown", "confidence": 0.0, "reason": "视觉接口未给出可用结论", "backend": self.backend, "error": str(exc), "error_type": type(exc).__name__}


class DefaultDateRewriter:
    """Invoke the existing packaged date engine through its matching Python runtime."""

    def __init__(self, script_dir: Path):
        self.bridge = script_dir / "date_year_rewriter_bridge.py"
        self.tool_root = script_dir.parent.parent / "图片日期智能处理工具_用户版_20260903"
        self.python = self.tool_root / "_internal" / "python.exe"
        self.available = self.bridge.is_file() and self.python.is_file() and (self.tool_root / "app.py").is_file()
        self.reason = "未找到图片日期工具的运行时或年份重绘桥接脚本"

    def __call__(self, path: Path, target_year: int) -> dict[str, Any]:
        if not self.available:
            return {"status": "UNAVAILABLE", "changed": False, "reason": self.reason}
        completed = subprocess.run(
            [str(self.python), str(self.bridge), str(self.tool_root), str(path), str(target_year)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900,
        )
        lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
        try:
            result = json.loads(lines[-1]) if lines else {}
        except json.JSONDecodeError:
            result = {}
        if completed.returncode != 0 or not isinstance(result, dict):
            return {
                "status": "FAILED", "changed": False,
                "reason": result.get("error") if isinstance(result, dict) else (completed.stderr.strip() or completed.stdout.strip()),
            }
        return result


def _docx_parts(path: Path) -> tuple[list[tuple[str, bytes]], list[tuple[str, str]]]:
    images: list[tuple[str, bytes]] = []
    texts: list[tuple[str, str]] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            lowered = name.lower()
            if lowered.startswith("word/media/") and Path(name).suffix.lower() in IMAGE_SUFFIXES:
                images.append((name, archive.read(name)))
            elif lowered == "word/document.xml":
                root = etree.fromstring(archive.read(name))
                for index, paragraph in enumerate(root.iter(f"{{{WORD_NS}}}p")):
                    value = "".join(node.text or "" for node in paragraph.iter(f"{{{WORD_NS}}}t"))
                    if value.strip():
                        texts.append((f"paragraph:{index + 1}", value))
    return images, texts


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
    if len(body.strip()) < 80:
        return {"status": "unknown", "reason": "正文可分析文字不足", "title_terms": {}, "other_topics": {}}
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
        return {"status": "unknown", "reason": "标题缺少可核对的领域词", "title_terms": {}, "other_topics": {}}
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
            "title_terms": title_terms,
            "other_topics": other_topics,
        }
    if (len(absent) == len(title_terms) and generic_operation_total >= 20
            and all(generic_operations.values())):
        return {
            "status": "review_topic_mismatch",
            "reason": "标题专有领域词在正文缺席，正文主要重复通用增删查导操作",
            "title_terms": title_terms, "other_topics": {"通用操作": generic_operations},
        }
    return {
        "status": "unknown", "reason": "未发现足够强的标题与正文主题冲突证据",
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


def _supported_vision_image(data: bytes, suffix: str) -> bool:
    suffix = suffix.lower()
    return (
        (suffix == ".png" and data.startswith(b"\x89PNG\r\n\x1a\n"))
        or (suffix in {".jpg", ".jpeg"} and data.startswith(b"\xff\xd8"))
        or (suffix == ".webp" and data.startswith(b"RIFF") and data[8:12] == b"WEBP")
    )


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
) -> dict[str, Any]:
    """Preflight every manual, then copy and mutate only after batch gating completes."""
    source_root = Path(input_root).resolve()
    destination_root = Path(output_root).resolve()
    report_file = Path(report_path).resolve()
    manuals = sorted(source_root.rglob("*说明.docx"), key=lambda path: str(path).casefold())
    ocr_adapter = ocr or DefaultOcrAdapter(Path(__file__).resolve().parent)
    relevance = relevance_classifier or _default_relevance
    vision_adapter = vision_classifier or LocalVisionAdapter()
    vision_backend = getattr(vision_adapter, "backend", "injected")
    vision_disabled_reason = ""
    date_adapter = date_rewriter or DefaultDateRewriter(Path(__file__).resolve().parent)
    projects: list[dict[str, Any]] = []
    prepared: list[tuple[Path, list[tuple[str, bytes]], list[tuple[str, str]], str, str]] = []
    for manual in manuals:
        images, paragraphs = _docx_parts(manual)
        title = manual.stem.removesuffix("说明")
        context = "\n".join(value for _, value in paragraphs[:3])
        prepared.append((manual, images, paragraphs, title, context))
    references = [f"{title}\n{context}" for _manual, _images, _paragraphs, title, context in prepared]

    for document_index, (manual, images, paragraphs, title, context) in enumerate(prepared, 1):
        samples = []
        visual_calls = 0
        visual_decisions = 0
        visual_failures = 0
        visual_tokens = 0
        sample_indexes = equidistant_sample_indexes(len(images))
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
        visual_candidates = [position for position, sample in enumerate(samples) if sample["status"] in {"unknown", "irrelevant"}]
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
            elif _supported_vision_image(image_bytes, suffix):
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
        blocked = bool(decisive and irrelevant / decisive > 0.5)
        projects.append({
            "source": str(manual),
            "relative_path": str(manual.relative_to(source_root)),
            "body_relevance": assess_body_relevance(title, paragraphs),
            "image_count": len(images),
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
            "status": "BLOCKED_IRRELEVANT" if blocked else "PREFLIGHT_PASSED",
        })
        _progress(progress, "说明相关性预检", document_index, manual.name)

    blocked_count = sum(project["blocked"] for project in projects)
    batch_fused = bool(projects and blocked_count / len(projects) > 0.2)

    # The first filesystem mutation intentionally happens only after every document is preflighted.
    if apply:
        for document_index, project in enumerate(projects, 1):
            source = Path(project["source"])
            destination = destination_root / project["relative_path"]
            destination.parent.mkdir(parents=True, exist_ok=True)
            if source != destination:
                shutil.copy2(source, destination)
            project["output"] = str(destination)
            project["year_changes"] = _rewrite_document_years(destination, year_classifier, target_year)
            if batch_fused:
                project["date_rewrite"] = {"status": "SKIPPED_BATCH_FUSE", "changed": False}
                project["status"] = "BATCH_FUSED"
            elif project["blocked"]:
                project["date_rewrite"] = {"status": "SKIPPED_PROJECT_BLOCKED", "changed": False}
            else:
                result = _call_adapter(date_adapter, destination, target_year)
                if not isinstance(result, dict):
                    result = {"status": "FAILED_INVALID_RESULT", "changed": False, "reason": "date_rewriter 必须返回字典"}
                result.setdefault("status", "CHANGED" if result.get("changed") else "UNCHANGED")
                result.setdefault("changed", False)
                project["date_rewrite"] = result
                if result["status"] == "UNAVAILABLE":
                    project["status"] = "DATE_REWRITER_UNAVAILABLE"
            _progress(progress, "说明内容处理", document_index, source.name)

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
        "visual_max_calls_per_project": vision_max_calls_per_project,
        "visual_model": VISION_MODEL if isinstance(vision_adapter, LocalVisionAdapter) else None,
        "visual_call_count": sum(project["visual_call_count"] for project in projects),
        "visual_decision_count": sum(project["visual_decision_count"] for project in projects),
        "visual_failure_count": sum(project["visual_failure_count"] for project in projects),
        "visual_total_tokens": sum(project["visual_total_tokens"] for project in projects),
        "visual_timeout_seconds": vision_adapter.timeout if isinstance(vision_adapter, LocalVisionAdapter) else None,
        "visual_error": vision_disabled_reason or None,
        "projects": projects,
    }
    report_file.parent.mkdir(parents=True, exist_ok=True)
    temporary_report = report_file.with_suffix(report_file.suffix + ".tmp")
    temporary_report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary_report, report_file)
    return report


__all__ = [
    "LocalVisionAdapter",
    "assess_body_relevance",
    "classify_year_context",
    "equidistant_sample_indexes",
    "replace_update_years",
    "run_manual_content_stage",
]

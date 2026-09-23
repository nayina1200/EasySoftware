from __future__ import annotations

import argparse
import concurrent.futures
import contextlib
import ctypes
import hashlib
import io
import json
import math
import os
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path, PurePosixPath

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")

from lxml import etree
from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageOps

# The batch contains user-supplied, high-resolution interface screenshots.  The
# review-board generator splits only between projects, so allow a single project
# sheet to retain its original pixels instead of rejecting it before that split.
Image.MAX_IMAGE_PIXELS = None
from pypdf import PdfReader

import archive_types


VERSION = "3.4.26"
CODE_WORD_TIMEOUT_SECONDS = 20
CODE_FALLBACK_TIMEOUT_SECONDS = 30
WORD_AUTOMATION_MUTEX = "Local\\EasySoftware.WordAutomation.v1"
INTERFACE_GROUP_SIZE = 5
MAX_REVIEW_DIMENSION = 30000
MAX_REVIEW_PIXELS = 220_000_000
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W, "a": A, "r": R, "v": V, "rel": REL, "mc": MC}

MANUAL_LABELS = (
    "使用说明书", "使用手册", "用户手册", "用户使用手册",
    "用户操作手册", "用户操作说明", "操作说明书", "操作手册",
    "软件设计说明书", "软件设计说明",
)
LABEL_RE = re.compile(r"^[【\[]?(?:" + "|".join(map(re.escape, MANUAL_LABELS)) + r")[】\]]?$")
LEGACY_LABEL_RE = re.compile(r"^[【\[]?(?:" + "|".join(map(re.escape, MANUAL_LABELS + ("说明书",))) + r")[】\]]?$")
SIMSUN_NAMES = {"宋体", "SimSun", "NSimSun", "宋体-简"}
ARCHIVE_SUFFIXES = archive_types.ARCHIVE_SUFFIXES
RELEVANT_SUFFIXES = {".txt", ".pdf", ".docx"}
MAIN_FUNCTION_RE = re.compile(r"^\s*(?:主要功能|软件主要功能)\s*[:：]\s*(.*)$")
MAIN_FUNCTION_VERB_RE = re.compile(r"(?:管理|查询|录入|维护|统计|分析|生成|导出|导入|审核|审批|发布|展示|监控|处理|配置|登录|注册|提醒|计算|检索|编辑|保存|删除|上传|下载|同步|控制|服务)")
LOGIN_REQUIRED_TERMS = (
    "管理", "平台", "系统", "后台", "中台", "门户", "商城", "医院",
    "业务", "运营", "办公", "审批", "审方", "调度", "监控", "协同",
    "政务", "教务", "财务", "人事", "会员", "订单", "仓储", "库存",
)
LOGIN_REQUIRED_LATIN_RE = re.compile(r"(?i)(?<![A-Z0-9])(?:ERP|CRM|OA|WMS|MES|SAAS)(?![A-Z0-9])")
LOGIN_ACTION_RE = re.compile(r"登录|登陆|进入系统|sign\s*in|log\s*in", re.I)
LOGIN_USER_RE = re.compile(r"账号|帐号|用户名|用户账号|登录名|账户|user\s*name|account", re.I)
LOGIN_PASSWORD_RE = re.compile(r"密码|口令|pwd|pass\s*word", re.I)
LOGIN_EXTRA_RE = re.compile(r"验证码|记住密码|忘记密码|captcha|登录框|login", re.I)
BODY_ORDINAL_PREFIX_RE = re.compile(
    r"^\s*(?:\d+[、.．)）]|[（(]\s*(?:\d+|[一二三四五六七八九十]+)\s*[）)]|"
    r"第[一二三四五六七八九十]+[章节部分、，.]?|[一二三四五六七八九十]+[、.．)）]|[•●○◆▪])"
)
IMAGE_CAPTION_RE = re.compile(r"^\s*(?:图|表|截图)\s*(?:[0-9０-９]+|[一二三四五六七八九十]+)")
ORDINAL_MARKER_RE = re.compile(
    r"(?:^|(?<=[，、；。\s]))(?:\d+[、.．]\s*|[（(]\s*(?:\d+|[一二三四五六七八九十]+)\s*[）)]\s*|"
    r"第[一二三四五六七八九十]+(?:、|，)?\s*|[一二三四五六七八九十]+、\s*|(?:首先|其次|然后|最后)(?:、|，)?\s*)"
)
REPEATED_PUNCTUATION_RE = re.compile(
    r"(?P<first>[。，“、；：！？,;:!?])(?:\s*[。，“、；：！？,;:!?])+"
)

TOPICS = {
    "地震": ("地震", "震情", "震级", "震源", "烈度", "地质", "灾害"),
    "医疗": ("医院", "医疗", "患者", "处方", "药品", "诊疗", "医生", "审方"),
    "环保": ("环保", "环境", "污染", "排放", "生态", "废水", "废气"),
    "教育": ("教育", "教学", "课程", "学生", "学习", "培训", "题库"),
    "婴幼儿": ("宝宝", "婴儿", "幼儿", "儿歌", "早教", "母婴"),
    "金融": ("金融", "银行", "证券", "基金", "贷款", "征信", "保险"),
    "农业": ("农业", "农作物", "种植", "养殖", "农田", "农产品"),
    "交通": ("交通", "车辆", "道路", "公交", "物流", "运输", "驾驶"),
    "电商": ("商城", "电商", "商品", "订单", "购物", "库存", "支付"),
}


def now_iso() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S%z")


def clean_text(value: str | None) -> str:
    return re.sub(r"\s+", " ", (value or "").replace("\r", " ").replace("\n", " ")).strip()


def compact_text(value: str | None) -> str:
    return re.sub(r"\s+", "", value or "")


def is_manual_label_candidate(value: str | None) -> bool:
    """Recognize cover labels beyond the fixed legacy vocabulary."""
    normalized = compact_text(value)
    normalized = re.sub(r"^[0-9一二三四五六七八九十]+[、.)）]", "", normalized)
    normalized = re.sub(r"[《》「」『』【】\[\]（）()\"“”]", "", normalized)
    if LEGACY_LABEL_RE.fullmatch(normalized):
        return True
    signals = {term for term in ("书", "说明", "软件", "设计", "使用") if term in normalized}
    return len(signals) >= 2 and bool(signals & {"书", "说明", "使用"})


COVER_VERSION_SUFFIX_RE = re.compile(r"(?i)(?:\s*(?:V|版本)\s*\d+(?:\.\d+)*|\s*\d+\.\d+(?:\.\d+)*)\s*$")


def strip_cover_version(value: str) -> str:
    return COVER_VERSION_SUFFIX_RE.sub("", value)


def is_cover_title(value: str | None, title: str | None) -> bool:
    normalized = compact_text(value)
    normalized = re.sub(
        r"^(?:\d+[、.．)）]|[（(](?:\d+|[一二三四五六七八九十]+)[）)]|"
        r"第[一二三四五六七八九十]+[章节部分、，.]?|[一二三四五六七八九十]+[、.．)）])",
        "",
        normalized,
    )
    # A supplied cover often puts the version on the title line (“软件名 V1.0”).
    # The cover itself must carry only the pure software name; the version lives
    # in the running header. Strip a trailing version marker before matching so
    # such a cover is detected and then repaired to the pure name.
    return strip_cover_version(normalized) == compact_text(title)


def is_image_caption(text: str, props: dict | None = None) -> bool:
    style = str((props or {}).get("style") or "")
    return bool(IMAGE_CAPTION_RE.match(text)) or bool(re.search(r"caption|题注|图注", style, re.I))


def safe_name(value: str) -> str:
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value).strip(" .")
    return value[:100] or "unnamed"


def atomic_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


class ProgressReporter:
    """Small stable status file consumed by the desktop GUI."""

    def __init__(self, work: Path, total: int = 0):
        self.path = work / "progress.json"
        self.total = total
        self.started = time.perf_counter()
        self.completed = 0

    def update(self, phase: str, current: int = 0, name: str = "", status: str = "RUNNING", error: str | None = None) -> None:
        if current > self.completed:
            self.completed = current
        completed = min(self.total, self.completed)
        elapsed = time.perf_counter() - self.started
        rate = (elapsed / completed) if completed else 0
        row = {
            "version": VERSION,
            "status": status,
            "phase": phase,
            "current": current,
            "total": self.total,
            "project_name": name,
            "elapsed_seconds": round(elapsed, 1),
            "completed_projects": completed,
            "remaining_projects": max(0, self.total - completed),
            "estimated_remaining_seconds": round(rate * max(0, self.total - completed), 1) if rate else None,
            "current_module": phase,
            "current_project_active": False,
            "updated_at": now_iso(),
        }
        if error:
            row["error"] = error
        atomic_json(self.path, row)


def load_json(path: Path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (FileNotFoundError, json.JSONDecodeError, UnicodeDecodeError):
        return default


def decode_process_output(value: str | bytes | None) -> str:
    """Decode Windows child-process output without turning Chinese errors into mojibake."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return value.decode(encoding)
        except UnicodeDecodeError:
            continue
    return value.decode("gb18030", errors="replace")


@contextlib.contextmanager
def word_automation_lock(
    progress: ProgressReporter | None,
    module: str,
    timeout_seconds: int = 180,
):
    """Allow only one EasySoftware batch to automate Microsoft Word at once."""
    if os.name != "nt":
        yield
        return
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.WaitForSingleObject.argtypes = (ctypes.c_void_p, ctypes.c_uint32)
    kernel32.ReleaseMutex.argtypes = (ctypes.c_void_p,)
    kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
    handle = kernel32.CreateMutexW(None, False, WORD_AUTOMATION_MUTEX)
    if not handle:
        raise RuntimeError("无法创建 Word 处理队列")
    acquired = False
    started = time.monotonic()
    try:
        while True:
            result = kernel32.WaitForSingleObject(handle, 5_000)
            if result in (0, 0x80):
                acquired = True
                break
            if result == 0x102:
                if progress:
                    progress.update("等待其他任务完成 Word 导出", 0, module or "Word处理")
                if time.monotonic() - started >= timeout_seconds:
                    raise TimeoutError(
                        f"等待 Word 处理队列超时（>{timeout_seconds}秒）。"
                        "请关闭其他 EasySoftware 任务或 Word 的阻塞窗口后重试。"
                    )
                continue
            raise RuntimeError("等待 Word 处理队列时发生异常")
        yield
    finally:
        if acquired:
            kernel32.ReleaseMutex(handle)
        kernel32.CloseHandle(handle)


class HashCache:
    def __init__(self, path: Path):
        self.path = path
        self.rows = load_json(path, {})
        self.changed = False

    def sha256(self, path: Path) -> str:
        stat = path.stat()
        key = str(path.resolve())
        signature = f"{stat.st_size}:{stat.st_mtime_ns}:{stat.st_ctime_ns}"
        row = self.rows.get(key)
        sentinel = self._sentinel(path, stat.st_size)
        if row and row.get("signature") == signature and row.get("sentinel") == sentinel and row.get("sha256"):
            return str(row["sha256"])
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        value = digest.hexdigest()
        entry = {"signature": signature, "sentinel": sentinel, "sha256": value}
        if row != entry:
            self.rows[key] = entry
            self.changed = True
        return value

    @staticmethod
    def _sentinel(path: Path, size: int) -> str:
        """Read only the edges before trusting a cached full-file SHA-256."""
        chunk_size = min(64 * 1024, size)
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            digest.update(handle.read(chunk_size))
            if size > chunk_size:
                handle.seek(max(0, size - chunk_size))
                digest.update(handle.read(chunk_size))
        return digest.hexdigest()

    def save(self) -> None:
        if self.changed:
            atomic_json(self.path, self.rows)


def font(size: int, bold: bool = False):
    paths = (
        r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    )
    for path in paths:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def find_7z() -> str | None:
    candidates = [shutil.which("7z"), shutil.which("7z.exe")]
    candidates += [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        r"C:\Program Files\AMD\CIM\Bin64\7z.exe",
        r"C:\Program Files\AMD\CNext\CNext\7z.exe",
        r"C:\Program Files\NVIDIA Corporation\NVIDIA GeForce Experience\7z.exe",
    ]
    return next((str(p) for p in candidates if p and Path(p).exists()), None)


def safe_zip_extract(archive: Path, target: Path) -> None:
    target_resolved = target.resolve()
    with zipfile.ZipFile(archive) as source:
        for member in source.infolist():
            output = (target / member.filename).resolve()
            if output != target_resolved and target_resolved not in output.parents:
                raise ValueError(f"压缩包包含越界路径：{archive.name} -> {member.filename}")
        source.extractall(target)


def extract_one(archive: Path, target: Path, seven_zip: str | None) -> dict:
    started = time.perf_counter()
    target.mkdir(parents=True, exist_ok=False)
    try:
        if archive.suffix.casefold() == ".zip":
            safe_zip_extract(archive, target)
        else:
            if not seven_zip:
                raise RuntimeError(f"RAR/7z 解压需要 7-Zip：{archive.name}")
            result = subprocess.run(
                [seven_zip, "x", "-y", f"-o{target}", str(archive)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                text=True,
            )
            if result.returncode != 0:
                raise RuntimeError(f"解压失败：{archive.name}：{result.stderr[-500:]}")
        return {"archive": archive.name, "target": str(target), "seconds": round(time.perf_counter() - started, 3)}
    except Exception:
        shutil.rmtree(target, ignore_errors=True)
        raise


def expand_nested_archives(
    extracted: Path,
    seven_zip: str | None,
    workers: int,
    hashes: HashCache,
    prior_results: list[dict] | None = None,
    max_depth: int = 8,
) -> list[dict]:
    """Expand nested archives and reuse targets only when their SHA-256 is unchanged."""
    expanded = set()
    results = []
    extracted_resolved = extracted.resolve()
    prior_by_archive = {
        str(row.get("archive", "")).casefold(): row
        for row in (prior_results or [])
        if row.get("archive") and row.get("sha256")
    }
    for depth in range(1, max_depth + 1):
        jobs = []
        for archive in sorted(
            (path for path in extracted.rglob("*") if path.is_file() and archive_types.is_archive_file(path)),
            key=lambda path: str(path.relative_to(extracted)).casefold(),
        ):
            key = str(archive.resolve()).casefold()
            if key in expanded:
                continue
            expanded.add(key)
            relative = str(archive.relative_to(extracted))
            digest = hashes.sha256(archive)
            target = archive.parent / archive.stem
            target_resolved = target.resolve()
            if target_resolved == extracted_resolved or extracted_resolved not in target_resolved.parents:
                raise RuntimeError(f"嵌套压缩包目标越界：{archive}")
            if target.is_dir() and any(target.iterdir()):
                prior = prior_by_archive.get(relative.casefold())
                if not prior or prior.get("sha256") != digest:
                    raise RuntimeError(f"嵌套压缩包解压目录缺少匹配的SHA-256记录，请使用新的解压目录：{target}")
                results.append({
                    "archive": relative,
                    "target": str(target),
                    "sha256": digest,
                    "depth": depth,
                    "status": "REUSED",
                })
                continue
            if target.exists():
                raise RuntimeError(f"嵌套压缩包解压目标被文件占用：{target}")
            jobs.append((archive, target, seven_zip, relative, digest, depth))
        if not jobs:
            break
        if len(expanded) > 500:
            raise RuntimeError("嵌套压缩包数量超过500，已停止以避免异常递归")
        with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, min(workers, 4))) as pool:
            future_map = {
                pool.submit(extract_one, archive, target, seven_zip): (archive, relative, digest, depth)
                for archive, target, seven_zip, relative, digest, depth in jobs
            }
            for future in concurrent.futures.as_completed(future_map):
                _, relative, digest, job_depth = future_map[future]
                row = future.result()
                row.update({
                    "archive": relative,
                    "sha256": digest,
                    "depth": job_depth,
                    "status": "EXTRACTED",
                })
                results.append(row)
    else:
        remaining = [
            path for path in extracted.rglob("*")
            if path.is_file() and archive_types.is_archive_file(path) and str(path.resolve()).casefold() not in expanded
        ]
        if remaining:
            raise RuntimeError(f"嵌套压缩层级超过{max_depth}层")
    results.sort(key=lambda row: (row.get("depth", 0), row["archive"].casefold()))
    return results


def ensure_extracted(batch_root: Path, extracted: Path, work: Path, hashes: HashCache, workers: int, input_archives: list[Path] | None = None) -> list[dict]:
    # Use escapes so this source remains correct even when edited from a legacy
    # Windows console code page.
    excluded_top_level = {"\u89e3\u538b\u7248", "\u4fee\u6539\u7248"}
    if extracted.parent.resolve() == batch_root.resolve():
        excluded_top_level.add(extracted.name)
    try:
        work_relative = work.resolve().relative_to(batch_root.resolve())
        if work_relative.parts:
            excluded_top_level.add(work_relative.parts[0])
    except ValueError:
        pass

    def is_input_file(path: Path) -> bool:
        relative = path.relative_to(batch_root)
        return not relative.parts or relative.parts[0] not in excluded_top_level

    if input_archives:
        archives = sorted((path.resolve() for path in input_archives), key=lambda path: str(path).casefold())
        for archive in archives:
            if not archive.is_file() or not archive_types.is_archive_file(archive):
                raise RuntimeError(f"指定的输入压缩包无效：{archive}")
            try:
                archive.relative_to(batch_root)
            except ValueError as exc:
                raise RuntimeError(f"输入压缩包不在批次目录内：{archive}") from exc
    else:
        all_archives = sorted(
            (p for p in batch_root.rglob("*") if p.is_file() and archive_types.is_archive_file(p) and is_input_file(p)),
            key=lambda p: str(p.relative_to(batch_root)).lower(),
        )
        # 文件夹入口以已解压材料为准：同名目录已存在的压缩包按"就地解压后的残留"
        # 处理，不再重复解压（压缩包有独立的选择入口）。
        leftover = {str(p) for p in all_archives if archive_types.is_leftover_archive(p)}
        archives = [p for p in all_archives if str(p) not in leftover]
        ignored_archives = [p for p in all_archives if str(p) in leftover]
    archive_state = [
        {"name": str(p.relative_to(batch_root)), "size": p.stat().st_size, "sha256": hashes.sha256(p)}
        for p in archives
    ]
    source_files = sorted(
        (p for p in batch_root.rglob("*") if p.is_file() and p.suffix.lower() in RELEVANT_SUFFIXES and is_input_file(p)),
        key=lambda p: str(p.relative_to(batch_root)).lower(),
    )
    source_state = [
        {"name": str(p.relative_to(batch_root)), "size": p.stat().st_size, "sha256": hashes.sha256(p)}
        for p in source_files
    ]
    state_path = work / "extraction.json"
    old_state = load_json(state_path, {})

    if extracted.exists() and any(extracted.iterdir()):
        if not old_state:
            raise RuntimeError("解压版已存在但缺少 extraction.json；请指定新的解压目录，避免复用未经验证的旧数据。")
        if (
            old_state.get("archives", []) != archive_state
            or old_state.get("source_files", []) != source_state
        ):
            raise RuntimeError("输入材料已变化，但解压版已存在。请指定新的解压目录，避免覆盖只读审计副本。")
        seven_zip = find_7z()
        nested_results = expand_nested_archives(
            extracted,
            seven_zip,
            workers,
            hashes,
            old_state.get("nested_archives", []),
        )
        results = old_state.get("results", [])
        if not old_state or old_state.get("version") != VERSION or nested_results:
            atomic_json(state_path, {
                "version": VERSION,
                "archives": archive_state,
                "source_files": source_state,
                "results": results,
                "nested_archives": nested_results,
                "reused_existing": True,
                "at": now_iso(),
            })
        return results

    if not archives:
        if not source_files:
            raise RuntimeError(f"批次目录没有压缩包，也没有已解压的软著材料（txt/pdf/docx）：{batch_root}")
        if extracted.exists():
            extracted.rmdir()

        def ignore_outputs(directory: str, names: list[str]) -> set[str]:
            # 残留压缩包留在原目录，不复制进只读审计副本。
            ignored = {name for name in names if archive_types.is_archive_file(name)}
            if Path(directory).resolve() == batch_root.resolve():
                ignored |= {name for name in names if name in excluded_top_level}
            return ignored

        shutil.copytree(batch_root, extracted, ignore=ignore_outputs)
        result = {
            "source": str(batch_root), "target": str(extracted), "mode": "PREEXTRACTED_COPY",
            "files": len(source_files),
        }
        atomic_json(state_path, {
            "version": VERSION,
            "archives": [],
            "source_files": source_state,
            "results": [result],
            "ignored_archives": [str(p.relative_to(batch_root)) for p in ignored_archives],
            "at": now_iso(),
        })
        return [result]
    extracted.mkdir(parents=True, exist_ok=True)
    seven_zip = find_7z()
    jobs = [(archive, extracted / archive.relative_to(batch_root).with_suffix(""), seven_zip) for archive in archives]
    results = []
    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, min(workers, 4))) as pool:
            futures = [pool.submit(extract_one, *job) for job in jobs]
            for future in concurrent.futures.as_completed(futures):
                results.append(future.result())
    except Exception:
        # Preserve successfully extracted siblings for diagnosis, but do not mark the stage complete.
        raise
    results.sort(key=lambda row: row["archive"])
    nested_results = expand_nested_archives(extracted, seven_zip, workers, hashes)
    results.extend(row for row in nested_results if row.get("status") == "EXTRACTED")
    atomic_json(state_path, {
        "version": VERSION,
        "archives": archive_state,
        "source_files": [],
        "results": results,
        "nested_archives": nested_results,
        "ignored_archives": [str(p.relative_to(batch_root)) for p in ignored_archives],
        "at": now_iso(),
    })
    return results


def decode_txt(path: Path) -> tuple[str, str, list[str]]:
    raw = path.read_bytes()
    candidates = []
    if raw.startswith(b"\xef\xbb\xbf"):
        candidates.append(("utf-8-sig", "utf-8-sig"))
    elif raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        candidates.append(("utf-16", "utf-16"))
    candidates.extend((("utf-8", "utf-8"), ("gb18030", "gb18030")))
    for label, codec in candidates:
        try:
            text = raw.decode(codec)
            problems = []
            if "\ufffd" in text:
                problems.append("包含Unicode替换字符")
            if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", text):
                problems.append("包含异常控制字符")
            mojibake_hits = sum(text.count(x) for x in ("锟斤拷", "烫烫烫", "屯屯屯", "銆", "鈥"))
            if mojibake_hits >= 3:
                problems.append("疑似明显乱码")
            return label, text, problems
        except UnicodeDecodeError:
            continue
    return "unknown", raw.decode("utf-8", errors="replace"), ["编码无法可靠识别"]


def topic_scores(text: str) -> dict[str, int]:
    return {name: sum(text.count(word) for word in words) for name, words in TOPICS.items()}


def dominant_topic(text: str) -> tuple[str | None, int]:
    scores = topic_scores(text)
    if not scores:
        return None, 0
    name = max(scores, key=scores.get)
    return (name, scores[name]) if scores[name] else (None, 0)


def classify_login_requirement(title: str) -> tuple[bool, list[str]]:
    normalized = clean_text(title)
    matches = [term for term in LOGIN_REQUIRED_TERMS if term in normalized]
    latin = LOGIN_REQUIRED_LATIN_RE.search(normalized)
    if latin:
        matches.append(latin.group(0).strip())
    return bool(matches), list(dict.fromkeys(matches))


def inspect_txt(path: Path, title: str) -> dict:
    encoding, text, problems = decode_txt(path)
    lines = text.splitlines()
    fields = []
    for number, line in enumerate(lines, 1):
        match = re.match(r"^\s*([^：:]{2,40})[：:]\s*(.*)$", line)
        if match:
            fields.append({"line": number, "name": clean_text(match.group(1)), "value": clean_text(match.group(2))})
    empty_fields = [f for f in fields if not f["value"]]
    duplicate_fields = sorted({f["name"] for f in fields if sum(x["name"] == f["name"] for x in fields) > 1})
    clear = [{"location": "编码", "message": item} for item in problems]
    main_function_lines = []
    for number, line in enumerate(lines, 1):
        match = MAIN_FUNCTION_RE.match(line)
        if match:
            main_function_lines.append((number, clean_text(match.group(1))))
    for number, value in main_function_lines:
        location = f"第{number}行主要功能"
        if ORDINAL_MARKER_RE.search(value):
            clear.append({"location": location, "message": "主要功能包含序号或序数词"})
        if re.search(r"[()（）]", value):
            clear.append({"location": location, "message": "主要功能包含括号"})
        if not value:
            clear.append({"location": location, "message": "主要功能为空，表述不通畅"})
        elif not re.search(r"[\u4e00-\u9fff]", value):
            clear.append({"location": location, "message": "主要功能缺少中文功能描述"})
        elif re.search(r"^[，、；。！？!?]|[，、；]$", value) or REPEATED_PUNCTUATION_RE.search(value):
            clear.append({"location": location, "message": "主要功能标点位置或连续标点异常，表述不通畅"})
        elif not MAIN_FUNCTION_VERB_RE.search(value):
            clear.append({"location": location, "message": "主要功能未包含可识别的功能动作，表述不符合要求"})
    if title not in text:
        clear.append({"location": "全文", "message": f"未找到软件名称“{title}”"})
    for field in empty_fields:
        clear.append({"location": f"第{field['line']}行/{field['name']}", "message": "字段值为空"})
    candidates = []
    if encoding == "gb18030":
        candidates.append({"location": "编码", "message": "编码为GB18030，批次要求UTF-8时需转换"})
    if duplicate_fields:
        candidates.append({"location": "字段", "message": "存在重复字段：" + "、".join(duplicate_fields)})
    title_topic, title_score = dominant_topic(title)
    body_topic, body_score = dominant_topic(text)
    if title_topic and body_topic and title_topic != body_topic and body_score >= 3 and title_score >= 1:
        candidates.append({"location": "内容字段", "message": f"名称偏向“{title_topic}”，TXT内容偏向“{body_topic}”"})
    return {
        "encoding": encoding,
        "line_count": len(lines),
        "field_count": len(fields),
        "name_present": title in text,
        "clear_errors": clear,
        "candidates": candidates,
        "text": text,
    }


def repair_main_function_text(value: str) -> tuple[str, dict[str, int]]:
    """Apply the requested conservative punctuation cleanup to a TXT value."""
    counts = {"parentheses_to_commas": 0, "quotes_removed": 0, "duplicate_punctuation_removed": 0, "ordinal_markers_removed": 0}

    def remove_ordinal(match: re.Match[str]) -> str:
        counts["ordinal_markers_removed"] += 1
        return ""

    while True:
        updated = ORDINAL_MARKER_RE.sub(remove_ordinal, value)
        if updated == value:
            break
        value = updated

    def replace_parenthesis(match: re.Match[str]) -> str:
        counts["parentheses_to_commas"] += 1
        return "\uFF0C"

    value = re.sub(r"[()\uFF08\uFF09]", replace_parenthesis, value)

    def remove_quote(match: re.Match[str]) -> str:
        counts["quotes_removed"] += 1
        return ""

    value = re.sub(r"[\"\u201C\u201D]", remove_quote, value)

    def collapse_punctuation(match: re.Match[str]) -> str:
        symbols = re.findall(r"[。，“、；：！？,;:!?]", match.group(0))
        counts["duplicate_punctuation_removed"] += len(symbols) - 1
        return match.group("first")

    value = REPEATED_PUNCTUATION_RE.sub(collapse_punctuation, value)
    return value, {key: count for key, count in counts.items() if count}


def write_txt_with_original_encoding(path: Path, original: bytes, text: str, encoding: str) -> None:
    if original.startswith(b"\xef\xbb\xbf"):
        path.write_bytes(b"\xef\xbb\xbf" + text.encode("utf-8"))
    elif original.startswith(b"\xff\xfe"):
        path.write_bytes(b"\xff\xfe" + text.encode("utf-16-le"))
    elif original.startswith(b"\xfe\xff"):
        path.write_bytes(b"\xfe\xff" + text.encode("utf-16-be"))
    else:
        path.write_bytes(text.encode(encoding if encoding != "unknown" else "utf-8"))


def repair_selected_txt(modified: Path, packages: list[dict], work: Path, progress: ProgressReporter | None = None) -> dict:
    report = {"version": VERSION, "created_at": now_iso(), "packages": []}
    field_re = re.compile(r"^(\s*(?:\u4e3b\u8981\u529f\u80fd|\u8f6f\u4ef6\u4e3b\u8981\u529f\u80fd)\s*[:\uFF1A]\s*)(.*)$")
    for index, package in enumerate(packages, 1):
        relative = package.get("files", {}).get("txt", {}).get("name")
        folder = modified / str(package.get("folder", ""))
        path = folder / str(relative or "")
        row = {
            "name": package.get("name", ""),
            "folder": package.get("folder", ""),
            "original_valid": not bool(package.get("txt", {}).get("clear_errors", [])),
            "repairs": {},
            "main_function_changes": [],
        }
        if not path.is_file():
            row["original_valid"] = False
            row["error"] = "TXT file is missing from modified project"
            report["packages"].append(row)
            continue
        if progress:
            progress.update("TXT repair", index, str(package.get("name", "")))
        original = path.read_bytes()
        encoding, text, _ = decode_txt(path)
        updated_lines = []
        changed = False
        for line in text.splitlines(keepends=True):
            match = field_re.match(line.rstrip("\r\n"))
            if not match:
                updated_lines.append(line)
                continue
            clean_value, counts = repair_main_function_text(match.group(2))
            row["repairs"] = {key: row["repairs"].get(key, 0) + value for key, value in counts.items()}
            if clean_value != match.group(2):
                row["main_function_changes"].append({"before": match.group(2), "after": clean_value})
                suffix = line[len(line.rstrip("\r\n")):]
                updated_lines.append(match.group(1) + clean_value + suffix)
                changed = True
            else:
                updated_lines.append(line)
        row["original_valid"] = row["original_valid"] and not bool(row["repairs"])
        row["changed"] = changed
        if changed:
            write_txt_with_original_encoding(path, original, "".join(updated_lines), encoding)
        report["packages"].append(row)
    atomic_json(work / "txt_repair_report.json", report)
    return report


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def paragraph_text(paragraph) -> str:
    return clean_text("".join(paragraph.xpath(".//w:t/text()", namespaces=NS)))


SENTENCE_BREAK_RE = re.compile(r"[。！？；!?;]")


def explicit_run_font_signature(run, paragraph_default: tuple) -> tuple | None:
    rpr = run.find(qn(W, "rPr"))
    if rpr is None:
        return None
    rfonts = rpr.find(qn(W, "rFonts"))
    font_names = tuple(sorted({
        value for attr in ("ascii", "hAnsi", "eastAsia", "cs")
        if rfonts is not None
        for value in [rfonts.get(qn(W, attr))]
        if value
    }))
    size = rpr.find(qn(W, "sz"))
    size_value = size.get(qn(W, "val")) if size is not None else None
    bold = rpr.find(qn(W, "b"))
    italic = rpr.find(qn(W, "i"))
    signature = (
        font_names or paragraph_default[0],
        size_value or paragraph_default[1],
        None if bold is None else bold.get(qn(W, "val"), "1") not in {"0", "false", "off"},
        None if italic is None else italic.get(qn(W, "val"), "1") not in {"0", "false", "off"},
    )
    return signature if any(value not in (None, ()) for value in signature) else None


def paragraph_sentence_font_mismatches(paragraph) -> list[dict]:
    """Find direct-format conflicts inside one Chinese sentence only."""
    ppr = paragraph.find(qn(W, "pPr"))
    default_rpr = ppr.find(qn(W, "rPr")) if ppr is not None else None
    default_fonts = default_rpr.find(qn(W, "rFonts")) if default_rpr is not None else None
    default_font_names = tuple(sorted({
        value for attr in ("ascii", "hAnsi", "eastAsia", "cs")
        if default_fonts is not None
        for value in [default_fonts.get(qn(W, attr))]
        if value
    }))
    default_size = None
    if default_rpr is not None:
        default_size_node = default_rpr.find(qn(W, "sz"))
        default_size = default_size_node.get(qn(W, "val")) if default_size_node is not None else None
    paragraph_default = (default_font_names, default_size)
    current_text, signatures, findings = "", set(), []

    def finish_sentence() -> None:
        nonlocal current_text, signatures
        compact = clean_text(current_text)
        if re.search(r"[\u3400-\u9fff]", compact) and len(signatures) > 1:
            findings.append({"text": compact[:80], "style_count": len(signatures)})
        current_text, signatures = "", set()

    for run in paragraph.xpath(".//w:r[w:t]", namespaces=NS):
        text = "".join(run.xpath(".//w:t/text()", namespaces=NS))
        signature = explicit_run_font_signature(run, paragraph_default)
        for fragment in re.split(r"([。！？；!?;])", text):
            if not fragment:
                continue
            current_text += fragment
            if signature is not None and re.search(r"[\u3400-\u9fff]", fragment):
                signatures.add(signature)
            if SENTENCE_BREAK_RE.fullmatch(fragment):
                finish_sentence()
    finish_sentence()
    return findings


def paragraph_props(paragraph) -> dict:
    fonts, sizes, bold_values, italic_values = set(), set(), [], []
    runs = paragraph.xpath(".//w:r[w:t]", namespaces=NS)
    for run in runs:
        rpr = run.find(qn(W, "rPr"))
        if rpr is None:
            continue
        rfonts = rpr.find(qn(W, "rFonts"))
        if rfonts is not None:
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                value = rfonts.get(qn(W, attr))
                if value:
                    fonts.add(value)
        size = rpr.find(qn(W, "sz"))
        if size is not None and size.get(qn(W, "val")):
            sizes.add(size.get(qn(W, "val")))
        bold = rpr.find(qn(W, "b"))
        if bold is not None:
            bold_values.append(bold.get(qn(W, "val"), "1") not in {"0", "false", "off"})
        italic = rpr.find(qn(W, "i"))
        if italic is not None:
            italic_values.append(italic.get(qn(W, "val"), "1") not in {"0", "false", "off"})
    ppr = paragraph.find(qn(W, "pPr"))
    spacing_before_points = 0.0
    alignment = None
    if ppr is not None:
        spacing = ppr.find(qn(W, "spacing"))
        if spacing is not None and spacing.get(qn(W, "before")):
            try:
                spacing_before_points = int(spacing.get(qn(W, "before"))) / 20
            except ValueError:
                pass
        justification = ppr.find(qn(W, "jc"))
        if justification is not None:
            alignment = justification.get(qn(W, "val"))
    style_node = ppr.find(qn(W, "pStyle")) if ppr is not None else None
    outline_node = ppr.find(qn(W, "outlineLvl")) if ppr is not None else None
    try:
        outline_level = int(outline_node.get(qn(W, "val"))) if outline_node is not None else None
    except (TypeError, ValueError):
        outline_level = None
    return {
        "fonts": sorted(fonts),
        "sizes_half_points": sorted(sizes),
        "bold": False if any(v is False for v in bold_values) else (True if bold_values else None),
        "italic": False if any(v is False for v in italic_values) else (True if italic_values else None),
        "spacing_before_points": spacing_before_points,
        "alignment": alignment,
        "style": style_node.get(qn(W, "val")) if style_node is not None else None,
        "outline_level": outline_level,
    }


def dhash_image(image: Image.Image) -> int:
    gray = image.convert("L").resize((9, 8), Image.Resampling.LANCZOS)
    value = 0
    pixels = list(gray.get_flattened_data() if hasattr(gray, "get_flattened_data") else gray.getdata())
    for y in range(8):
        offset = y * 9
        for x in range(8):
            value = (value << 1) | int(pixels[offset + x] > pixels[offset + x + 1])
    return value


def normalized_rms(first: Image.Image, second: Image.Image) -> float:
    a = first.convert("L").resize((160, 120), Image.Resampling.LANCZOS)
    b = second.convert("L").resize((160, 120), Image.Resampling.LANCZOS)
    histogram = ImageChops.difference(a, b).histogram()
    squares = sum(value * (index ** 2) for index, value in enumerate(histogram))
    return math.sqrt(squares / (160 * 120))


def interface_style_signature(image: Image.Image) -> tuple[int, int, int]:
    """Return a cheap palette signature while discounting white/gray backgrounds."""
    resized = image.convert("RGB").resize((64, 64), Image.Resampling.LANCZOS)
    pixels = list(resized.get_flattened_data() if hasattr(resized, "get_flattened_data") else resized.getdata())
    colored = [pixel for pixel in pixels if max(pixel) - min(pixel) >= 28 and max(pixel) <= 248]
    sample = colored if len(colored) >= 40 else pixels
    return tuple(round(sum(pixel[index] for pixel in sample) / len(sample)) for index in range(3))


def relationship_map(zip_file: zipfile.ZipFile) -> dict[str, str]:
    name = "word/_rels/document.xml.rels"
    if name not in zip_file.namelist():
        return {}
    root = etree.fromstring(zip_file.read(name))
    result = {}
    for rel in root.xpath("./rel:Relationship", namespaces=NS):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            path = PurePosixPath("word") / PurePosixPath(target)
            parts = []
            for part in path.parts:
                if part == "..":
                    if parts:
                        parts.pop()
                elif part != ".":
                    parts.append(part)
            result[rid] = "/".join(parts)
    return result


def page_field_count_in_story_xml(story: etree._Element) -> int:
    """Count visible PAGE fields while ignoring AlternateContent fallbacks."""
    fields = story.xpath(".//w:instrText | .//w:fldSimple", namespaces=NS)
    count = 0
    for field in fields:
        if field.xpath("ancestor::mc:Fallback", namespaces=NS):
            continue
        instruction = field.text if field.tag == qn(W, "instrText") else field.get(qn(W, "instr"))
        if re.search(r"\bPAGE\b", instruction or "", re.I):
            count += 1
    return count


def inspect_active_docx_headers(source: zipfile.ZipFile, names: set[str], root: etree._Element, title: str) -> dict:
    """Inspect only the headers that can appear on a rendered document page."""
    rels = relationship_map(source)
    settings = etree.fromstring(source.read("word/settings.xml")) if "word/settings.xml" in names else None
    odd_even_headers = bool(settings is not None and settings.xpath(".//w:evenAndOddHeaders", namespaces=NS))
    active_header_ids: list[str] = []
    active_footer_ids: list[str] = []
    for section_props in root.xpath(".//w:sectPr", namespaces=NS):
        types = {"default"}
        if odd_even_headers:
            types.add("even")
        if section_props.xpath("./w:titlePg", namespaces=NS):
            types.add("first")
        for header_type in types:
            active_header_ids.extend(
                section_props.xpath("./w:headerReference[@w:type=$type]/@r:id", namespaces=NS, type=header_type)
            )
            active_footer_ids.extend(
                section_props.xpath("./w:footerReference[@w:type=$type]/@r:id", namespaces=NS, type=header_type)
            )
    active_header_names = {rels[header_id] for header_id in active_header_ids if header_id in rels}
    active_footer_names = {rels[footer_id] for footer_id in active_footer_ids if footer_id in rels}
    headers, header_page_counts, footer_page_counts = [], [], []
    for name in sorted(active_header_names):
        if name not in names:
            continue
        header = etree.fromstring(source.read(name))
        header_text = clean_text("".join(header.xpath(".//w:t/text()", namespaces=NS)))
        headers.append(header_text)
        header_page_counts.append(page_field_count_in_story_xml(header))
    for name in sorted(active_footer_names):
        if name not in names:
            continue
        footer = etree.fromstring(source.read(name))
        footer_page_counts.append(page_field_count_in_story_xml(footer))
    active_count = max(len(headers), len(footer_page_counts))
    header_page_counts.extend([0] * max(0, active_count - len(header_page_counts)))
    footer_page_counts.extend([0] * max(0, active_count - len(footer_page_counts)))
    global_header_compliant = bool(headers) and all(value for value in headers) and all(
        header_count + footer_count == 1
        for header_count, footer_count in zip(header_page_counts, footer_page_counts)
    )
    return {
        "header_texts": headers,
        "header_title": all(title in value for value in headers),
        "header_version": all("V1.0" in value for value in headers),
        "header_page_field": all(count == 1 for count in header_page_counts),
        "page_field_count": sum(header_page_counts) + sum(footer_page_counts),
        "header_page_field_counts": header_page_counts,
        "footer_page_field_counts": footer_page_counts,
        "footer_page_field_count": sum(footer_page_counts),
        "header_compliant": global_header_compliant,
    }


def inspect_docx_header_only(path: Path, title: str) -> dict:
    with zipfile.ZipFile(path) as source:
        names = set(source.namelist())
        root = etree.fromstring(source.read("word/document.xml"))
        return inspect_active_docx_headers(source, names, root, title)


def inspect_docx(path: Path, title: str, check_login: bool) -> dict:
    with zipfile.ZipFile(path) as source:
        names = set(source.namelist())
        root = etree.fromstring(source.read("word/document.xml"))
        direct_paragraphs = root.xpath("/w:document/w:body/w:p", namespaces=NS)
        paragraphs = root.xpath("/w:document/w:body//w:p", namespaces=NS)
        paragraph_rows = [
            {
                "index": i + 1,
                "text": paragraph_text(p),
                "props": paragraph_props(p),
                "sentence_font_mismatches": paragraph_sentence_font_mismatches(p),
            }
            for i, p in enumerate(direct_paragraphs)
        ]
        paragraph_texts = [paragraph_text(p) for p in paragraphs]
        body_text = "\n".join("".join(p.xpath(".//w:t/text()", namespaces=NS)) for p in paragraphs)
        header_audit = inspect_active_docx_headers(source, names, root, title)
        rels = relationship_map(source)
        images = []
        seen_occurrences = set()
        for p_index, paragraph in enumerate(paragraphs, 1):
            nearby = clean_text(" ".join(paragraph_texts[max(0, p_index - 3):min(len(paragraph_texts), p_index + 1)]))
            rids = paragraph.xpath(".//a:blip/@r:embed | .//v:imagedata/@r:id", namespaces=NS)
            for rid in rids:
                target = rels.get(rid)
                occurrence = (p_index, rid, target)
                if not target or target not in names or occurrence in seen_occurrences:
                    continue
                seen_occurrences.add(occurrence)
                data = source.read(target)
                try:
                    with Image.open(io.BytesIO(data)) as opened:
                        image = opened.convert("RGB")
                        width, height = image.size
                        dhash = dhash_image(image)
                        style_rgb = interface_style_signature(image)
                except Exception:
                    width = height = 0
                    dhash = None
                    style_rgb = None
                images.append({
                    "order": len(images) + 1,
                    "paragraph": p_index,
                    "nearby_text": nearby,
                    "relationship_id": rid,
                    "media_path": target,
                    "sha256": hashlib.sha256(data).hexdigest(),
                    "bytes": len(data),
                    "width": width,
                    "height": height,
                    "dhash": dhash,
                    "style_rgb": style_rgb,
                })

        substantial = [
            row for row in images
            if row["bytes"] >= 10_000 and row["width"] >= 320 and row["height"] >= 180
        ]
        login_screen_review = inspect_login_screens_in_order(source, substantial) if check_login else {"status": "not_checked", "matches": []}
        style_candidates = []
        if len(substantial) >= 2:
            ratios = [row["width"] / row["height"] for row in substantial]
            median_ratio = statistics.median(ratios)
            signatures = [row["style_rgb"] for row in substantial if row["style_rgb"] is not None]
            median_rgb = tuple(statistics.median(channel) for channel in zip(*signatures)) if signatures else None
            for row, ratio in zip(substantial, ratios):
                ratio_delta = abs(math.log(max(ratio, 0.01) / max(median_ratio, 0.01)))
                palette_distance = None
                if median_rgb is not None and row["style_rgb"] is not None:
                    palette_distance = math.sqrt(sum((a - b) ** 2 for a, b in zip(row["style_rgb"], median_rgb)))
                # A single ratio or palette deviation is common for dialogs and
                # mobile-shaped screens. Flag only screenshots that differ on
                # both dimensions from the project's dominant visual style.
                if ratio_delta >= 0.34 and palette_distance is not None and palette_distance >= 95:
                    style_candidates.append({
                        "order": row["order"],
                        "paragraph": row["paragraph"],
                        "reasons": [],
                        "ratio": round(ratio, 3),
                        "palette_distance": round(palette_distance, 1) if palette_distance is not None else None,
                    })
            if style_candidates:
                mismatch_count = len(style_candidates)
                for candidate in style_candidates:
                    candidate["reasons"] = [f"有{mismatch_count}个图片与项目主流截图的配色和大体均明显不一致"]
            # Screenshots often share the same export size even when they come
            # from two unrelated products. Detect a clear light/dark theme
            # split as a separate, review-worthy palette inconsistency.
            themed = [(row, sum(row["style_rgb"]) / 3) for row in substantial if row["style_rgb"] is not None]
            if len(themed) >= 6:
                median_brightness = statistics.median(value for _, value in themed)
                light = [(row, value) for row, value in themed if value >= median_brightness + 35]
                dark = [(row, value) for row, value in themed if value <= median_brightness - 35]
                if len(light) >= 3 and len(dark) >= 3:
                    outliers = light if len(light) <= len(dark) else dark
                    existing = {candidate["order"] for candidate in style_candidates}
                    for row, value in outliers:
                        if row["order"] not in existing:
                            style_candidates.append({
                                "order": row["order"], "paragraph": row["paragraph"],
                                "reasons": [], "ratio": round(row["width"] / row["height"], 3),
                                "palette_distance": round(abs(value - median_brightness), 1),
                            })
                    theme_count = len(outliers)
                    for candidate in style_candidates:
                        if not candidate["reasons"]:
                            candidate["reasons"] = [f"发现深色和浅色两套主配色，其中{theme_count}个图片偏离项目主流配色"]
        login_evidence = [
            {
                "order": row["order"],
                "paragraph": row["paragraph"],
                "nearby_text": row["nearby_text"][:500],
                "media_path": row["media_path"],
            }
            for row in substantial
            if check_login and re.search(r"登录|登陆|账号|帐号|用户名|密码|验证码|sign\s*in|log\s*in", row["nearby_text"], re.I)
        ]
        login_candidates = []
        if check_login and substantial:
            first = substantial[0]
            first_bytes = source.read(first["media_path"])
            with Image.open(io.BytesIO(first_bytes)) as opened:
                first_image = opened.convert("RGB")
            for candidate in substantial[1:]:
                exact = candidate["sha256"] == first["sha256"]
                distance = (first["dhash"] ^ candidate["dhash"]).bit_count()
                ratio_delta = abs((first["width"] / first["height"]) - (candidate["width"] / candidate["height"]))
                if exact or (distance <= 6 and ratio_delta <= 0.15):
                    candidate_bytes = source.read(candidate["media_path"])
                    with Image.open(io.BytesIO(candidate_bytes)) as opened:
                        rms = normalized_rms(first_image, opened.convert("RGB"))
                    if exact or distance <= 3 or rms <= 24:
                        login_candidates.append({
                            "first_order": first["order"],
                            "candidate_order": candidate["order"],
                            "candidate_paragraph": candidate["paragraph"],
                            "exact": exact,
                            "dhash_distance": distance,
                            "rms": round(rms, 2),
                            "first_media": first["media_path"],
                            "candidate_media": candidate["media_path"],
                        })

    legacy_blocks = []
    for i, row in enumerate(paragraph_rows):
        if not is_cover_title(row["text"], title):
            continue
        for j in range(i + 1, min(len(paragraph_rows), i + 7)):
            candidate_text = paragraph_rows[j]["text"]
            if not candidate_text:
                continue
            if is_manual_label_candidate(candidate_text):
                legacy_blocks.append({
                    "title_paragraph": row["index"],
                    "label_paragraph": paragraph_rows[j]["index"],
                    "label": candidate_text,
                    "title_text": row["text"],
                    "accepted_label": LABEL_RE.fullmatch(candidate_text) is not None,
                    "title_props": row["props"],
                    "label_props": paragraph_rows[j]["props"],
                })
            # Once visible non-label content starts, a later label belongs to
            # the body and must not be paired with the earlier title as a
            # cover. This also prevents the repair stage from deleting it as a
            # supposed duplicate cover fragment.
            break
    # A supplied "software name + 说明书/手册" page is already a cover.  Keep
    # it instead of inserting another cover merely because its label is an
    # older naming variant.
    blocks = list(legacy_blocks)
    cover = next((block for block in blocks if block["title_paragraph"] <= 20), blocks[0] if blocks else None)
    duplicate_blocks = [
        block for block in legacy_blocks
        if cover is None or (block["title_paragraph"], block["label_paragraph"]) != (cover["title_paragraph"], cover["label_paragraph"])
    ]
    title_rows = [row for row in paragraph_rows if row["index"] <= 30 and is_cover_title(row["text"], title)]
    if len(title_rows) > 1:
        for row in title_rows[1:]:
            duplicate_blocks.append({
                "title_paragraph": row["index"],
                "label_paragraph": None,
                "label": "仅有软件名称的重复封面",
                "accepted_label": False,
                "title_props": row["props"],
                "label_props": {},
                "isolated_title_only": True,
            })
    paired_label_indexes = {block["label_paragraph"] for block in legacy_blocks}
    residual_manual_labels = [
        row for row in paragraph_rows
        if row["index"] <= 30
        and is_manual_label_candidate(row["text"])
        and row["index"] not in paired_label_indexes
    ]
    style_ok = False
    layout_ok = False
    if cover:
        title_props, label_props = cover["title_props"], cover["label_props"]
        fonts = set(title_props["fonts"] + label_props["fonts"])
        fonts_ok = bool(fonts) and all(any(expected.lower() in value.lower() for expected in SIMSUN_NAMES) for value in fonts)
        sizes_ok = (
            len(title_props["sizes_half_points"]) == 1
            and title_props["sizes_half_points"] == label_props["sizes_half_points"]
        )
        style_ok = fonts_ok and sizes_ok and title_props["bold"] is True and label_props["bold"] is True
        vertically_centered = bool(root.xpath(".//w:sectPr/w:vAlign[@w:val='center']", namespaces=NS))
        layout_ok = (
            vertically_centered
            and title_props.get("alignment") == "center"
            and label_props.get("alignment") == "center"
            and label_props.get("spacing_before_points", 0) >= 36
        )
    body_eligible_rows = [
        row for row in paragraph_rows
        if (len(row["text"]) >= 2 or is_image_caption(row["text"], row["props"]))
        and compact_text(row["text"]) != compact_text(title)
        and not is_manual_label_candidate(row["text"])
        and not BODY_ORDINAL_PREFIX_RE.match(row["text"])
        and (row["props"].get("outline_level") is None or row["props"].get("outline_level") >= 9)
    ]
    body_candidates = [
        row for row in body_eligible_rows
        if row["props"].get("fonts") and row["props"].get("sizes_half_points")
    ]
    first_body_style_mismatch = False
    body_text_style_mismatch = False
    if len(body_candidates) >= 2:
        def body_signature(row):
            return (
                tuple(row["props"]["fonts"]), tuple(row["props"]["sizes_half_points"]),
                row["props"].get("bold") is True, row["props"].get("italic") is True,
            )
        signatures = [body_signature(row) for row in body_candidates]
        common_signature = max(set(signatures), key=signatures.count)
        first_body_style_mismatch = body_signature(body_candidates[0]) != common_signature
        body_text_style_mismatch = any(signature != common_signature for signature in signatures)
    # OOXML can inherit a paragraph's font from its Word style, leaving no
    # direct rFonts entry. Let Word resolve those effective styles when a
    # normal-body paragraph still carries direct formatting evidence.
    body_text_style_needs_word_check = (
        len(body_eligible_rows) >= 2
        and any(not row["props"].get("fonts") or not row["props"].get("sizes_half_points") for row in body_eligible_rows)
        and any(
            row["props"].get("fonts") or row["props"].get("sizes_half_points")
            or row["props"].get("bold") is not None or row["props"].get("italic") is not None
            for row in body_eligible_rows
        )
    )
    inline_body_mismatches = [
        {"paragraph": row["index"], **finding}
        for row in body_eligible_rows
        for finding in row["sentence_font_mismatches"]
    ]
    inline_heading_mismatches = [
        {"paragraph": row["index"], **finding}
        for row in paragraph_rows
        if row not in body_eligible_rows
        for finding in row["sentence_font_mismatches"]
    ]
    first_post_cover_heading_needs_word_check = False
    if cover:
        after_cover = [
            row for row in paragraph_rows
            if row["index"] > cover["label_paragraph"]
            and row["text"]
            and not is_cover_title(row["text"], title)
            and not is_manual_label_candidate(row["text"])
        ]
        if after_cover:
            candidate = after_cover[0]
            if len(candidate["text"]) <= 80 and not re.search(r"[。！？；;!?]", candidate["text"]):
                references = [
                    row for row in after_cover[1:]
                    if len(row["text"]) <= 80
                    and not re.search(r"[。！？；:：;!?]", row["text"])
                    and (BODY_ORDINAL_PREFIX_RE.match(row["text"]) or row["props"].get("bold") is True)
                    and row["props"].get("fonts")
                    and row["props"].get("sizes_half_points")
                ]
                if len(references) >= 2:
                    def heading_signature(row):
                        return (
                            tuple(row["props"]["fonts"]), tuple(row["props"]["sizes_half_points"]),
                            row["props"].get("bold") is True, row["props"].get("italic") is True,
                        )
                    reference_signatures = [heading_signature(row) for row in references]
                    common_heading_signature = max(set(reference_signatures), key=reference_signatures.count)
                    first_post_cover_heading_needs_word_check = heading_signature(candidate) != common_heading_signature
    return {
        "valid_zip": True,
        "paragraphs_front": paragraph_rows,
        "cover_blocks": blocks,
        "legacy_cover_blocks": legacy_blocks,
        "duplicate_cover_blocks": duplicate_blocks,
        "residual_manual_labels": residual_manual_labels,
        "cover_found": cover is not None,
        "cover_label": cover["label"] if cover else None,
        "cover_title_text": cover["title_text"] if cover else None,
        "cover_title_has_version": bool(cover and COVER_VERSION_SUFFIX_RE.search(cover["title_text"])),
        "cover_style_explicit_ok": style_ok,
        "cover_layout_explicit_ok": layout_ok,
        **header_audit,
        "first_body_style_mismatch": first_body_style_mismatch,
        "body_text_style_mismatch": body_text_style_mismatch,
        "body_text_style_needs_word_check": body_text_style_needs_word_check,
        "inline_body_font_mismatches": inline_body_mismatches,
        "inline_heading_font_mismatches": inline_heading_mismatches,
        "first_post_cover_heading_needs_word_check": first_post_cover_heading_needs_word_check,
        "cover_uses_separate_first_header": bool(root.xpath(".//w:sectPr/w:titlePg", namespaces=NS)),
        "mentions_login": "登录" in body_text or "登陆" in body_text,
        "body_text": body_text,
        "image_count": len(images),
        "substantial_image_count": len(substantial),
        "substantial_images": [
            {key: row[key] for key in ("order", "paragraph", "nearby_text", "media_path", "width", "height", "style_rgb")}
            for row in substantial
        ],
        "interface_style_candidates": style_candidates,
        "login_evidence": login_evidence,
        "login_screen_review": login_screen_review,
        "images": images,
        "login_candidates": login_candidates,
    }


def login_signal_map(text: str) -> dict[str, bool]:
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
    return {
        "登录动作": bool(LOGIN_ACTION_RE.search(text)),
        "用户名": bool(LOGIN_USER_RE.search(text)),
        "密码": bool(LOGIN_PASSWORD_RE.search(text)),
        "验证码或辅助项": bool(LOGIN_EXTRA_RE.search(text)),
    }


def has_reliable_login_context(text: str) -> bool:
    signals = login_signal_map(text)
    return (signals["用户名"] and signals["密码"]) or (signals["登录动作"] and signals["用户名"])


def is_reliable_login_screen(screen_text: str, nearby_text: str = "") -> tuple[bool, dict[str, bool], dict[str, bool]]:
    screen_signals = login_signal_map(screen_text)
    combined_signals = login_signal_map(f"{nearby_text} {screen_text}")
    reliable = sum(screen_signals.values()) >= 2 or (
        sum(screen_signals.values()) >= 1 and sum(combined_signals.values()) >= 2
    )
    return reliable, screen_signals, combined_signals


def windows_ocr_available() -> bool:
    return os.name == "nt" and Path(__file__).with_name("windows_ocr.ps1").is_file()


def windows_ocr_text(image_path: Path) -> str:
    if not windows_ocr_available():
        return ""
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(Path(__file__).with_name("windows_ocr.ps1")), str(image_path)],
        text=True, encoding="utf-8", errors="replace", stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
        timeout=30,
    )
    return result.stdout if result.returncode == 0 else ""


def ocr_backend_available() -> bool:
    return bool(shutil.which("tesseract") or shutil.which("tesseract.exe") or windows_ocr_available())


def ocr_image_text(image: Image.Image) -> str:
    executable = shutil.which("tesseract") or shutil.which("tesseract.exe")
    with tempfile.TemporaryDirectory(prefix="softcopy-ocr-") as temporary:
        image_path = Path(temporary) / "screen.png"
        if not executable:
            image.save(image_path, format="PNG")
            original = windows_ocr_text(image_path)
            if sum(login_signal_map(original).values()) >= 2:
                return original
            prepared = ImageOps.autocontrast(ImageOps.grayscale(image))
            if max(prepared.size) < 1600:
                ratio = 1600 / max(prepared.size)
                prepared = prepared.resize(
                    (round(prepared.width * ratio), round(prepared.height * ratio)), Image.Resampling.LANCZOS
                )
            prepared.save(image_path, format="PNG")
            refined = windows_ocr_text(image_path)
            return "\n".join(dict.fromkeys(line.strip() for line in (original + "\n" + refined).splitlines() if line.strip()))

        def read(current: Image.Image, psm: int) -> str:
            current.save(image_path, format="PNG")
            result = subprocess.run(
                [executable, str(image_path), "stdout", "-l", "chi_sim+eng", "--psm", str(psm)],
                text=True, encoding="utf-8", errors="replace", stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                timeout=20,
            )
            return result.stdout if result.returncode == 0 else ""

        original = read(image, 11)
        if sum(login_signal_map(original).values()) >= 2:
            return original
        prepared = ImageOps.autocontrast(ImageOps.grayscale(image))
        if max(prepared.size) < 1600:
            ratio = 1600 / max(prepared.size)
            prepared = prepared.resize(
                (round(prepared.width * ratio), round(prepared.height * ratio)), Image.Resampling.LANCZOS
            )
        refined = read(prepared, 6)
    return "\n".join(dict.fromkeys(line.strip() for line in (original + "\n" + refined).splitlines() if line.strip()))


LOGIN_FOLLOWUP_IMAGE_COUNT = 6


def inspect_login_screens_in_order(source: zipfile.ZipFile, images: list[dict]) -> dict:
    """Scan in order, then inspect the next six screenshots after first login."""
    if not ocr_backend_available():
        return {"status": "unavailable", "matches": []}
    matches = []
    first_match_index: int | None = None
    for image_index, row in enumerate(images):
        if first_match_index is not None and image_index > first_match_index + LOGIN_FOLLOWUP_IMAGE_COUNT:
            break
        try:
            with Image.open(io.BytesIO(source.read(row["media_path"]))) as opened:
                ocr_text = ocr_image_text(opened.convert("RGB"))
        except Exception:
            ocr_text = ""
        reliable, screen_signals, combined_signals = is_reliable_login_screen(
            ocr_text, str(row.get("nearby_text", ""))
        )
        # The screenshot must contribute at least one signal. Nearby Word text
        # is only corroboration, so a body paragraph mentioning login cannot
        # be mistaken for a login screen by itself.
        if reliable:
            matches.append({
                "order": row["order"], "paragraph": row["paragraph"],
                "source": (
                    "OCR" if sum(screen_signals.values()) >= 2
                    else "附近说明文字（OCR无有效文字）" if not any(screen_signals.values())
                    else "OCR + 附近说明文字"
                ),
                "signals": [key for key, present in combined_signals.items() if present],
            })
            if len(matches) == 2:
                return {"status": "multiple", "matches": matches}
            first_match_index = image_index
            continue
    return {"status": "missing" if not matches else "single", "matches": matches}


def inspect_pdf(path: Path, title: str, code: bool = False) -> dict:
    reader = PdfReader(str(path))
    pages = len(reader.pages)
    first_text = reader.pages[0].extract_text() or "" if pages else ""
    last_text = reader.pages[-1].extract_text() or "" if pages else ""
    first_lines = [clean_text(line) for line in first_text.splitlines() if clean_text(line)]
    last_lines = [clean_text(line) for line in last_text.splitlines() if clean_text(line)]

    def has_page_number(lines: list[str], number: int) -> bool:
        expected = str(number)
        return any(
            re.fullmatch(rf"(?:第\s*)?{re.escape(expected)}\s*(?:页)?", line)
            or re.search(rf"第\s*{re.escape(expected)}\s*页", line)
            # Word frequently exports the right-aligned PAGE field onto the
            # same extracted line as the left header text: “软件名 V1.0 1”.
            or re.search(rf"V\s*1\.0\s+{re.escape(expected)}\s*$", line, re.I)
            for line in lines
        )
    result = {
        "pages": pages,
        "first_text": first_text,
        "last_text": last_text,
        "first_has_title": compact_text(title) in compact_text(first_text),
        "last_has_title": compact_text(title) in compact_text(last_text),
        "first_has_version": bool(re.search(r"V\s*1\.0", first_text, re.I)),
        "last_has_version": bool(re.search(r"V\s*1\.0", last_text, re.I)),
        "first_has_page_number": has_page_number(first_lines, 1),
        "last_has_page_number": has_page_number(last_lines, pages),
        "title_on_one_extracted_line": any(compact_text(title) in compact_text(line) for line in first_lines),
    }
    if code:
        combined = first_text + "\n" + last_text
        lines = combined.splitlines()
        code_markers = sum(combined.count(mark) for mark in ("{", "}", "(", ")", ";", "=", "def ", "class ", "public ", "function "))
        balances = {pair: combined.count(pair[0]) - combined.count(pair[1]) for pair in ("()", "[]", "{}")}
        indented = sum(bool(re.match(r"^[ \t]{2,}\S", line)) for line in lines)
        result["format_metrics"] = {
            "code_markers": code_markers,
            "indented_lines": indented,
            "bracket_balance": balances,
            "looks_like_code": code_markers >= 6 and len(lines) >= 10,
        }
    return result


def cover_page_is_clean(first_text: str, title: str) -> bool:
    residual = []
    for raw_line in first_text.splitlines():
        line = clean_text(raw_line)
        if not line:
            continue
        if title in line and re.search(r"V\s*1\.0", line, re.I):
            continue
        for label in MANUAL_LABELS:
            line = line.replace(label, "")
        line = re.sub(r"^[\-—–\s]*\d+[\-—–\s]*$", "", line)
        line = re.sub(r"[\-—–|｜\s]", "", line)
        if line:
            residual.append(line)
    joined = "".join(residual)
    return joined in {"", title}


def cover_title_is_one_line(first_text: str, title: str) -> bool:
    return any(
        title in clean_text(line) and not re.search(r"V\s*1\.0", line, re.I)
        for line in first_text.splitlines()
    )


def locate_package_dirs(base: Path) -> list[Path]:
    candidates = set()
    for txt in base.rglob("*.txt"):
        if any(part.startswith(".") for part in txt.relative_to(base).parts):
            continue
        parent = txt.parent
        files = list(parent.iterdir())
        if any(p.suffix.lower() in {".docx", ".pdf"} for p in files if p.is_file()):
            candidates.add(parent)
    if not candidates:
        candidates = {p for p in base.iterdir() if p.is_dir()}
    return sorted(candidates, key=lambda p: str(p.relative_to(base)))


def choose_files(folder: Path) -> dict:
    files = [
        p for p in folder.iterdir()
        if p.is_file()
        and not p.name.startswith("~$")
        and ".codex-new-" not in p.name
        and not p.name.endswith(".working.pdf")
    ]
    txts = sorted(p for p in files if p.suffix.lower() == ".txt")
    title = txts[0].stem if len(txts) == 1 else ""
    manual_words = (*MANUAL_LABELS, "说明", "手册", "文档")
    code_words = ("代码", "源程序", "源代码", "程序代码")

    def ranked(candidates: list[Path]) -> list[Path]:
        return sorted(candidates, key=lambda path: (
            not (title and title in path.stem),
            len(path.stem),
            path.name.casefold(),
        ))

    docx_all = [p for p in files if p.suffix.lower() == ".docx"]
    pdf_all = [p for p in files if p.suffix.lower() == ".pdf"]
    code_docxs = ranked([p for p in docx_all if any(word in p.stem for word in code_words)])
    code_pdfs = ranked([p for p in pdf_all if any(word in p.stem for word in code_words)])
    docxs = ranked([p for p in docx_all if p not in code_docxs and any(word in p.stem for word in manual_words)])
    manual_pdfs = ranked([p for p in pdf_all if p not in code_pdfs and any(word in p.stem for word in manual_words)])
    if not docxs and len([p for p in docx_all if p not in code_docxs]) == 1:
        docxs = [p for p in docx_all if p not in code_docxs]
    if not manual_pdfs and len([p for p in pdf_all if p not in code_pdfs]) == 1:
        manual_pdfs = [p for p in pdf_all if p not in code_pdfs]
    return {
        "txt": txts[0] if txts else None,
        "docx": docxs[0] if docxs else None,
        "code_docx": code_docxs[0] if code_docxs else None,
        "code_pdf": code_pdfs[0] if code_pdfs else None,
        "manual_pdf": manual_pdfs[0] if manual_pdfs else None,
        "multiple": {
            "txt": [str(p) for p in txts[1:]],
            "docx": [str(p) for p in docxs[1:]],
            "code_docx": [str(p) for p in code_docxs[1:]],
            "code_pdf": [str(p) for p in code_pdfs[1:]],
            "manual_pdf": [str(p) for p in manual_pdfs[1:]],
        },
    }


def add_code_pdf_audit(row: dict, path: Path, title: str, cached_code: dict | None = None) -> None:
    """Attach deterministic first/last-page code checks to a package row."""
    try:
        code = cached_code if cached_code is not None else inspect_pdf(path, title, code=True)
        row["code_pdf"] = {k: v for k, v in code.items() if k not in {"first_text", "last_text"}}
        if not 70 <= code["pages"] <= 130:
            row["candidates"].append({"file": path.name, "location": "总页数", "message": f"共{code['pages']}页，偏离约100页"})
        if not code["format_metrics"]["looks_like_code"]:
            row["candidates"].append({"file": path.name, "location": "第一页和最后一页", "message": "文本特征不像常规代码，需看首尾页拼图"})
        if not (code["first_has_title"] and code["last_has_title"]):
            row["candidates"].append({"file": path.name, "location": "第一页/最后一页页眉", "message": "软件名称未同时出现在首尾页"})
        if not (code["first_has_version"] and code["last_has_version"]):
            row["candidates"].append({"file": path.name, "location": "第一页/最后一页页眉", "message": "V1.0未同时出现在首尾页"})
        if not (code["first_has_page_number"] and code["last_has_page_number"]):
            row["candidates"].append({"file": path.name, "location": "第一页/最后一页页眉或页脚", "message": "未同时提取到首尾页码，需结合拼图确认"})
    except Exception as exc:
        row["clear_errors"].append({"file": path.name, "location": "PDF", "message": f"解析失败：{exc}"})


def inspect_package(
    folder: Path,
    base: Path,
    hashes: HashCache,
    cached_doc: dict | None = None,
    cached_code: dict | None = None,
    code_header_repairs: set[str] | None = None,
) -> dict:
    started = time.perf_counter()
    files = choose_files(folder)
    title = files["txt"].stem if files["txt"] else folder.name
    row = {
        "name": title,
        "folder": str(folder.relative_to(base)),
        "files": {},
        "clear_errors": [],
        "candidates": [],
        "actions": [],
        "code_actions": [],
    }
    labels = (("txt", "TXT"), ("code_pdf", "代码PDF"), ("docx", "说明DOCX"), ("manual_pdf", "说明PDF"))
    for key, label in labels:
        path = files[key]
        if path:
            row["files"][key] = {"name": path.name, "sha256": hashes.sha256(path), "bytes": path.stat().st_size}
        else:
            row["clear_errors"].append({"file": label, "location": "文件夹", "message": f"缺少{label}"})
    for key, extras in files["multiple"].items():
        if extras:
            row["candidates"].append({"file": key, "location": "文件夹", "message": "存在多个候选文件"})

    doc = None
    if files["docx"]:
        try:
            doc = cached_doc if cached_doc is not None else inspect_docx(files["docx"], title, False)
        except (zipfile.BadZipFile, KeyError, etree.XMLSyntaxError) as exc:
            row["clear_errors"].append({"file": files["docx"].name, "location": "DOCX结构", "message": f"DOCX无法解析：{exc}"})
    if files["txt"]:
        txt = inspect_txt(files["txt"], title)
        row["txt"] = {k: v for k, v in txt.items() if k != "text"}
        row["clear_errors"].extend({"file": files["txt"].name, **item} for item in txt["clear_errors"])
        row["candidates"].extend({"file": files["txt"].name, **item} for item in txt["candidates"])
    else:
        txt = {"text": ""}
    forced_code_repair = row["folder"] in (code_header_repairs or set()) or title in (code_header_repairs or set())
    if files["code_pdf"]:
        add_code_pdf_audit(row, files["code_pdf"], title, cached_code)
        code = row.get("code_pdf", {})
        code_docx_header = None
        if files["code_docx"]:
            try:
                code_docx_header = inspect_docx_header_only(files["code_docx"], title)
            except (zipfile.BadZipFile, KeyError, etree.XMLSyntaxError) as exc:
                row["candidates"].append({"file": files["code_docx"].name, "location": "代码DOCX页眉", "message": f"代码DOCX页眉无法读取，将重新整理：{exc}"})
        header_mismatch = not all((
            code.get("first_has_title"), code.get("last_has_title"),
            code.get("first_has_version"), code.get("last_has_version"),
            code.get("first_has_page_number"), code.get("last_has_page_number"),
        ))
        docx_header_mismatch = code_docx_header is not None and not code_docx_header.get("header_compliant", False)
        if header_mismatch or docx_header_mismatch or forced_code_repair:
            if files["code_docx"]:
                row["files"]["code_docx"] = {"name": files["code_docx"].name, "sha256": hashes.sha256(files["code_docx"]), "bytes": files["code_docx"].stat().st_size}
                row["code_actions"] = ["NORMALIZE_CODE_HEADER", "EXPORT_CODE_PDF"]
            else:
                row["clear_errors"].append({"file": files["code_pdf"].name, "location": "代码DOCX", "message": "代码PDF需要从DOCX重新导出，但未找到代码DOCX；未直接修改PDF"})
    elif files["code_docx"]:
        row["files"]["code_docx"] = {"name": files["code_docx"].name, "sha256": hashes.sha256(files["code_docx"]), "bytes": files["code_docx"].stat().st_size}
        row["code_actions"] = ["NORMALIZE_CODE_HEADER", "EXPORT_CODE_PDF"]
    if doc is None:
        row["seconds"] = round(time.perf_counter() - started, 3)
        return row
    body_text = str(doc.get("body_text", ""))
    if doc.get("body_text_style_mismatch") or doc.get("body_text_style_needs_word_check") or doc.get("inline_body_font_mismatches"):
        row["actions"].append("NORMALIZE_BODY_TEXT_STYLE")
    if doc.get("duplicate_cover_blocks"):
        row["actions"].append("REMOVE_DUPLICATE_COVER_TEXT")
    if doc.get("residual_manual_labels"):
        row["actions"].append("REMOVE_DUPLICATE_COVER_TEXT")
    if doc.get("first_post_cover_heading_needs_word_check"):
        row["actions"].append("NORMALIZE_FIRST_POST_COVER_HEADING")
    if re.search(rf'["“”「」『』]\s*{re.escape(title)}\s*["“”「」『』]', body_text):
        row["actions"].append("REMOVE_TITLE_DOUBLE_QUOTES")
    if re.search(r'["“”「」『』]', body_text):
        row["actions"].append("REMOVE_ALL_DOUBLE_QUOTES")
    if re.search(r'(?:19|20)\d{2}\s*年', body_text):
        row["actions"].append("REPLACE_YEARS_WITH_CURRENT")
    row["actions"] = list(dict.fromkeys(row["actions"]))
    row["manual_docx"] = {k: v for k, v in doc.items() if k not in {"body_text", "paragraphs_front", "images"}}
    for finding in doc.get("inline_heading_font_mismatches", []):
        row["candidates"].append({
            "file": files["docx"].name,
            "location": f"说明书第{finding['paragraph']}段标题",
            "message": f"同一句中文文字混用 {finding['style_count']} 种字体格式：{finding['text']}；需统一后复核",
        })
    for candidate in doc.get("interface_style_candidates", []):
        row["candidates"].append({
            "file": files["docx"].name,
            "location": f"正文第{candidate['paragraph']}段附近/第{candidate['order']}张主要图片",
            "message": "；".join(candidate["reasons"]) + "，需查看该项目界面组图确认是否来自同一软件",
        })
    if not doc["cover_found"]:
        row["actions"].append("ADD_COVER")
    if doc["cover_found"] and (not doc["cover_style_explicit_ok"] or not doc.get("cover_layout_explicit_ok", False)):
        row["actions"].append("NORMALIZE_COVER_STYLE")
    if doc.get("cover_title_has_version"):
        # 封面名称只允许纯软件名称；版本号只出现在页眉。封面标题带“名称 V1.0”
        # 时同样判为不合格，需去掉版本后缀并整理封面格式。
        row["actions"].append("REMOVE_COVER_VERSION")
        row["actions"].append("NORMALIZE_COVER_STYLE")
    if not doc.get("header_compliant"):
        row["actions"].append("NORMALIZE_HEADER")
    title_topic, title_score = dominant_topic(title)
    body_topic, body_score = dominant_topic(doc["body_text"])
    if title_topic and body_topic and title_topic != body_topic and title_score >= 1 and body_score >= 5:
        target = row["clear_errors"] if body_score >= 8 else row["candidates"]
        target.append({
            "file": files["docx"].name,
            "location": "正文主题",
            "message": f"软件名称偏向“{title_topic}”，说明正文持续偏向“{body_topic}”，内容明显不对应" if body_score >= 8 else f"软件名称偏向“{title_topic}”，说明内容偏向“{body_topic}”，需人工确认",
        })
    if files["manual_pdf"]:
        try:
            manual_pdf = inspect_pdf(files["manual_pdf"], title)
            row["manual_pdf"] = {k: v for k, v in manual_pdf.items() if k not in {"first_text", "last_text"}}
            row["manual_pdf"]["cover_page_clean"] = cover_page_is_clean(manual_pdf["first_text"], title)
            if not manual_pdf["first_has_title"]:
                row["actions"].append("ENSURE_COVER_FIRST_PAGE")
                row["actions"].append("EXPORT_PDF")
            elif doc and doc["cover_found"] and not row["manual_pdf"]["cover_page_clean"]:
                row["actions"].append("ADD_COVER")
                row["actions"].append("REMOVE_DUPLICATE_COVER_TEXT")
            if doc and doc["cover_found"] and not cover_title_is_one_line(manual_pdf["first_text"], title):
                row["actions"].append("NORMALIZE_COVER_STYLE")
        except Exception as exc:
            row["clear_errors"].append({"file": files["manual_pdf"].name, "location": "PDF", "message": f"解析失败：{exc}"})
            if files["docx"]:
                row["actions"].append("EXPORT_PDF")
    if files["docx"] and files["manual_pdf"] and files["docx"].stat().st_mtime_ns > files["manual_pdf"].stat().st_mtime_ns:
        row["actions"].append("EXPORT_PDF")
    if row["actions"] and "EXPORT_PDF" not in row["actions"]:
        row["actions"].append("EXPORT_PDF")
    row["actions"] = list(dict.fromkeys(row["actions"]))
    row["seconds"] = round(time.perf_counter() - started, 3)
    return row


def poppler_executable() -> Path | None:
    # The GUI pipeline may run under a different Python (e.g. Anaconda on PATH)
    # than the runtime that bundles Poppler, so also search the known runtime
    # location instead of relying only on sys.executable's neighbors.
    candidates = [
        Path(sys.executable).resolve().parents[1] / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe",
        Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies" / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    found = shutil.which("pdftoppm") or shutil.which("pdftoppm.exe")
    return Path(found) if found else None


def render_page_via_pdfium(path: Path, page_index: int, width: int) -> Image.Image:
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(path))
    try:
        page = document[page_index]
        try:
            page_width, _ = page.get_size()
            bitmap = page.render(scale=max(0.15, width / page_width))
            return bitmap.to_pil().convert("RGB")
        finally:
            page.close()
    finally:
        document.close()


def render_document_via_pdfium(path: Path, width: int) -> list[Image.Image]:
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(path))
    images = []
    try:
        for page in document:
            page_width, _ = page.get_size()
            bitmap = page.render(scale=max(0.15, width / page_width))
            images.append(bitmap.to_pil().convert("RGB"))
            page.close()
    finally:
        document.close()
    return images


def render_pdf_page(path: Path, page_index: int, width: int) -> Image.Image:
    if os.environ.get("SOFTCOPY_USE_PDFIUM") == "1":
        try:
            return render_page_via_pdfium(path, page_index, width)
        except Exception:
            pass
    # Poppler is the default because it tolerates older Word PDFs that can crash PDFium.
    executable = poppler_executable()
    if executable is None:
        # No Poppler reachable; fall back to the in-process renderer instead of
        # launching a non-existent pdftoppm.
        return render_page_via_pdfium(path, page_index, width)
    with tempfile.TemporaryDirectory(prefix="softcopy-render-") as temp:
        prefix = Path(temp) / "page"
        page_number = page_index + 1
        subprocess.run(
            [str(executable), "-f", str(page_number), "-l", str(page_number), "-scale-to-x", str(width), "-scale-to-y", "-1", "-singlefile", "-jpeg", "-jpegopt", "quality=86", str(path), str(prefix)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return Image.open(prefix.with_suffix(".jpg")).convert("RGB").copy()


def render_pdf_pages(path: Path, width: int) -> list[Image.Image]:
    """Render a whole manual with one Poppler process and low-I/O JPEG intermediates."""
    executable = poppler_executable()
    if executable is None:
        return render_document_via_pdfium(path, width)
    with tempfile.TemporaryDirectory(prefix="softcopy-manual-") as temp:
        prefix = Path(temp) / "page"
        subprocess.run(
            [str(executable), "-scale-to-x", str(width), "-scale-to-y", "-1", "-jpeg", "-jpegopt", "quality=82", str(path), str(prefix)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        def page_number(candidate: Path) -> int:
            match = re.search(r"-(\d+)\.jpg$", candidate.name, re.I)
            return int(match.group(1)) if match else 0

        images = []
        for candidate in sorted(Path(temp).glob("page-*.jpg"), key=page_number):
            with Image.open(candidate) as opened:
                images.append(opened.convert("RGB").copy())
        return images


def _unused_pdfium_reference(path: Path, page_index: int, width: int) -> Image.Image:
    """Kept out of the normal path; documents the optional in-process renderer."""
    try:
        import pypdfium2 as pdfium

        document = pdfium.PdfDocument(str(path))
        page = document[page_index]
        page_width, _ = page.get_size()
        bitmap = page.render(scale=max(0.15, width / page_width))
        image = bitmap.to_pil().convert("RGB")
        page.close()
        document.close()
        return image
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc


def make_manual_sheet(package: dict, base: Path, out_dir: Path) -> str | None:
    if package.get("skip_remaining"):
        return None
    info = package.get("files", {}).get("manual_pdf")
    if not info:
        return None
    pdf = base / package["folder"] / info["name"]
    count = package.get("manual_pdf", {}).get("pages") or len(PdfReader(str(pdf)).pages)
    thumbs = render_pdf_pages(pdf, 350)
    if not thumbs:
        return None
    cols, label_h, heading_h = 4, 30, 58
    cell_w = 370
    cell_h = max(image.height for image in thumbs) + label_h + 12
    rows = math.ceil(len(thumbs) / cols)
    canvas = Image.new("RGB", (cell_w * cols, heading_h + cell_h * rows), "#d7dbe0")
    draw = ImageDraw.Draw(canvas)
    draw.text((14, 12), f"{package['name']}｜说明书共 {count} 页", fill="black", font=font(25, True))
    for index, image in enumerate(thumbs):
        x = (index % cols) * cell_w + 10
        y = heading_h + (index // cols) * cell_h
        draw.text((x, y + 2), f"第 {index + 1} 页", fill="black", font=font(17))
        canvas.paste(image, (x, y + label_h))
    out_dir.mkdir(parents=True, exist_ok=True)
    output = out_dir / f"manual_{safe_name(package['name'])}.jpg"
    canvas.save(output, quality=88, optimize=True)
    return str(output)


def make_manual_first_two_sheet(package: dict, base: Path, out_dir: Path) -> str | None:
    """Create the user pre-audit preview containing only manual PDF pages 1 and 2."""
    if package.get("skip_remaining"):
        return None
    info = package.get("files", {}).get("manual_pdf")
    if not info:
        return None
    pdf = base / package["folder"] / info["name"]
    page_count = package.get("manual_pdf", {}).get("pages") or len(PdfReader(str(pdf)).pages)
    images = [render_pdf_page(pdf, index, 720) for index in range(min(2, page_count))]
    if not images:
        return None
    gap, pad, heading_h, label_h = 28, 24, 62, 38
    cell_w = max(image.width for image in images)
    cell_h = max(image.height for image in images)
    canvas = Image.new(
        "RGB",
        (pad * 2 + cell_w * len(images) + gap * (len(images) - 1), heading_h + label_h + cell_h + pad),
        "#d7dbe0",
    )
    draw = ImageDraw.Draw(canvas)
    draw.text((pad, 13), f"{package['name']} | 说明PDF前两页", fill="black", font=font(25, True))
    for index, image in enumerate(images):
        x = pad + index * (cell_w + gap)
        draw.text((x, heading_h), f"第 {index + 1} 页", fill="black", font=font(18, True))
        canvas.paste(image, (x, heading_h + label_h))
    out_dir.mkdir(parents=True, exist_ok=True)
    output = out_dir / f"manual_first_two_{safe_name(package['name'])}.jpg"
    canvas.save(output, quality=90, optimize=True)
    return str(output)


def make_batch_overview(rows: list[tuple[int, str, str]], output: Path, title: str, item_width: int = 1050) -> str | None:
    """Combine ordered per-project sheets into one numbered two-column review board."""
    items = []
    for index, name, path in rows:
        if not path or not Path(path).is_file():
            continue
        with Image.open(path) as opened:
            image = opened.convert("RGB")
        height = round(image.height * item_width / image.width)
        items.append((index, name, image.resize((item_width, height), Image.Resampling.LANCZOS)))
    if not items:
        return None
    cols, gap, pad, heading_h, item_heading_h = 2, 28, 32, 76, 56
    rows_count = math.ceil(len(items) / cols)
    row_heights = [
        max(items[row * cols + col][2].height for col in range(cols) if row * cols + col < len(items))
        for row in range(rows_count)
    ]
    canvas = Image.new(
        "RGB",
        (pad * 2 + cols * item_width + gap, heading_h + pad + sum(height + item_heading_h for height in row_heights) + gap * (rows_count - 1)),
        "#eef0f3",
    )
    draw = ImageDraw.Draw(canvas)
    draw.text((pad, 16), title, fill="black", font=font(30, True))
    y = heading_h
    for row, row_height in enumerate(row_heights):
        x = pad
        for col in range(cols):
            item_index = row * cols + col
            if item_index >= len(items):
                break
            number, name, image = items[item_index]
            draw.rectangle((x - 2, y - 2, x + item_width + 2, y + item_heading_h + image.height + 2), fill="white", outline="#a5aab2", width=2)
            draw.text((x + 12, y + 10), f"{number}. {name}", fill="black", font=font(25, True))
            canvas.paste(image, (x, y + item_heading_h))
            x += item_width + gap
        y += row_height + item_heading_h + gap
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, quality=88, optimize=True)
    return str(output)


def make_code_board(packages: list[dict], base: Path, output: Path) -> str | None:
    cards = []
    for package in packages:
        if package.get("skip_remaining"):
            continue
        info = package.get("files", {}).get("code_pdf")
        if not info:
            continue
        path = base / package["folder"] / info["name"]
        pages = package.get("code_pdf", {}).get("pages", 0)
        if pages < 1:
            continue
        cards.append((package["name"], 1, render_pdf_page(path, 0, 620)))
        cards.append((package["name"], pages, render_pdf_page(path, pages - 1, 620)))
    if not cards:
        return None
    cols, label_h = 2, 48
    cell_w = 650
    cell_h = max(image.height for _, _, image in cards) + label_h + 12
    rows = math.ceil(len(cards) / cols)
    canvas = Image.new("RGB", (cell_w * cols, cell_h * rows), "#d7dbe0")
    draw = ImageDraw.Draw(canvas)
    for index, (title, page_number, image) in enumerate(cards):
        x = (index % cols) * cell_w + 10
        y = (index // cols) * cell_h
        draw.text((x, y + 8), f"{title}｜第 {page_number} 页", fill="black", font=font(21, True))
        canvas.paste(image, (x, y + label_h))
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, quality=90, optimize=True)
    return str(output)


def extract_docx_media(path: Path, media_names: list[str]) -> list[Image.Image]:
    images = []
    with zipfile.ZipFile(path) as source:
        for name in media_names:
            with Image.open(io.BytesIO(source.read(name))) as opened:
                image = opened.convert("RGB")
                images.append(image.copy())
    return images


def make_login_gate_board(packages: list[dict], base: Path, output: Path) -> str | None:
    cards = []
    for package in packages:
        media = package.get("login_gate_media", [])
        info = package.get("files", {}).get("docx")
        if not media or not info:
            continue
        path = base / package["folder"] / info["name"]
        images = extract_docx_media(path, [row["media_path"] for row in media])
        for row, image in zip(media, images):
            cards.append((package["name"], row, image))
    if not cards:
        return None
    cols, cell_w, cell_h, label_h = 2, 660, 560, 72
    rows = math.ceil(len(cards) / cols)
    canvas = Image.new("RGB", (cell_w * cols, cell_h * rows), "#d7dbe0")
    draw = ImageDraw.Draw(canvas)
    for index, (title, row, image) in enumerate(cards):
        x = (index % cols) * cell_w + 10
        y = (index // cols) * cell_h
        draw.text((x, y + 8), f"{title}｜第{row['order']}张主要图片｜段落{row['paragraph']}", fill="black", font=font(20, True))
        context = clean_text(row.get("nearby_text", ""))[:46]
        if context:
            draw.text((x, y + 40), context, fill="#333333", font=font(15))
        canvas.paste(image, (x, y + label_h))
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, quality=90, optimize=True)
    return str(output)


def make_interface_sheet(package: dict, base: Path, out_dir: Path) -> str | None:
    media = package.get("manual_docx", {}).get("substantial_images", [])
    info = package.get("files", {}).get("docx")
    if not media or not info:
        return None
    path = base / package["folder"] / info["name"]
    images = extract_docx_media(path, [row["media_path"] for row in media])
    candidate_orders = {
        row["order"] for row in package.get("manual_docx", {}).get("interface_style_candidates", [])
    }
    login_review = package.get("manual_docx", {}).get("login_screen_review", {})
    login_orders = [
        int(match["order"])
        for match in login_review.get("matches", [])
        if str(match.get("order", "")).isdigit()
    ] if login_review.get("status") == "multiple" else []
    login_ranks = {order: index + 1 for index, order in enumerate(login_orders)}
    cols, heading_h, label_h, gap, pad = min(3, len(images)), 62, 44, 18, 18
    cell_w = max(image.width for image in images) + pad * 2
    cell_h = max(image.height for image in images) + label_h + pad * 2
    rows = math.ceil(len(images) / cols)
    canvas = Image.new("RGB", (cols * cell_w + (cols - 1) * gap, heading_h + rows * cell_h + (rows - 1) * gap), "#d7dbe0")
    draw = ImageDraw.Draw(canvas)
    heading = f"{package['name']}｜界面截图共 {len(images)} 张"
    if login_orders:
        heading += "｜警报：发现多个登录页"
    draw.text((14, 12), heading, fill="#b00020" if login_orders else "black", font=font(25, True))
    for index, (row, image) in enumerate(zip(media, images)):
        x = (index % cols) * (cell_w + gap) + pad
        y = heading_h + (index // cols) * (cell_h + gap)
        login_rank = login_ranks.get(row["order"])
        flagged = row["order"] in candidate_orders or login_rank is not None
        label = f"第{row['order']}张｜段落{row['paragraph']}｜{row['width']}×{row['height']}"
        if login_rank is not None:
            label += f"｜重点：登录页 {login_rank}/{len(login_orders)}"
        draw.text((x, y + 7), label, fill="#b00020" if flagged else "black", font=font(17, flagged))
        paste_x = x + (cell_w - pad * 2 - image.width) // 2
        paste_y = y + label_h
        canvas.paste(image, (paste_x, paste_y))
        if flagged:
            border_width = 7 if login_rank is not None else 4
            draw.rectangle((paste_x - 3, paste_y - 3, paste_x + image.width + 3, paste_y + image.height + 3), outline="#b00020", width=border_width)
    out_dir.mkdir(parents=True, exist_ok=True)
    folder_key = hashlib.sha1(str(package.get("folder", "")).encode("utf-8")).hexdigest()[:10]
    output = out_dir / f"interfaces_{safe_name(package['name'])}_{folder_key}.png"
    canvas.save(output, format="PNG", compress_level=3)
    return str(output)


def make_interface_overview_groups(rows: list[tuple[int, str, str]], out_dir: Path) -> list[str]:
    """Combine five ordered project sheets per group without resampling any project pixels."""
    available = [(index, name, path) for index, name, path in rows if path and Path(path).is_file()]
    if not available:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    outputs = []
    for group_no, start in enumerate(range(0, len(available), INTERFACE_GROUP_SIZE), 1):
        group = available[start:start + INTERFACE_GROUP_SIZE]
        pending = [group]
        part_no = 0
        while pending:
            part = pending.pop(0)
            opened = []
            for index, name, path in part:
                with Image.open(path) as source:
                    opened.append((index, name, source.convert("RGB").copy()))
            cols = min(2, len(opened))
            item_heading_h, title_h, gap, pad = 56, 76, 24, 24
            row_count = math.ceil(len(opened) / cols)
            col_widths = [max(opened[i][2].width for i in range(col, len(opened), cols)) for col in range(cols)]
            row_heights = [max(opened[i][2].height for i in range(row * cols, min((row + 1) * cols, len(opened)))) for row in range(row_count)]
            width = pad * 2 + sum(col_widths) + gap * (cols - 1)
            height = title_h + pad + sum(value + item_heading_h for value in row_heights) + gap * (row_count - 1)
            if (width > MAX_REVIEW_DIMENSION or height > MAX_REVIEW_DIMENSION or width * height > MAX_REVIEW_PIXELS) and len(part) > 1:
                midpoint = math.ceil(len(part) / 2)
                pending[0:0] = [part[:midpoint], part[midpoint:]]
                continue
            part_no += 1
            canvas = Image.new("RGB", (width, height), "#eef0f3")
            draw = ImageDraw.Draw(canvas)
            first_index, last_index = part[0][0], part[-1][0]
            suffix = f"｜分片 {part_no}" if len(part) != len(group) or pending else ""
            draw.text((pad, 16), f"程序界面审核组 {group_no}｜项目 {first_index}-{last_index}{suffix}", fill="black", font=font(30, True))
            y = title_h
            for row_index, row_height in enumerate(row_heights):
                x = pad
                for col in range(cols):
                    item_index = row_index * cols + col
                    if item_index >= len(opened):
                        break
                    number, name, image = opened[item_index]
                    draw.rectangle((x - 2, y - 2, x + image.width + 2, y + item_heading_h + image.height + 2), fill="white", outline="#a5aab2", width=2)
                    draw.text((x + 10, y + 10), f"{number}. {name}", fill="black", font=font(25, True))
                    canvas.paste(image, (x, y + item_heading_h))
                    x += col_widths[col] + gap
                y += row_height + item_heading_h + gap
            output = out_dir / f"interfaces_overview_{first_index:03d}_{last_index:03d}_part{part_no}.png"
            canvas.save(output, format="PNG", compress_level=3)
            outputs.append(str(output))
    return outputs


def make_login_board(packages: list[dict], base: Path, output: Path) -> str | None:
    cards = []
    for package in packages:
        candidates = package.get("manual_docx", {}).get("login_candidates", [])
        info = package.get("files", {}).get("docx")
        if not candidates or not info:
            continue
        path = base / package["folder"] / info["name"]
        for candidate in candidates:
            images = extract_docx_media(path, [candidate["first_media"], candidate["candidate_media"]])
            cards.append((package["name"], candidate, images[0], images[1]))
    if not cards:
        return None
    cell_w, cell_h = 1340, 560
    canvas = Image.new("RGB", (cell_w, cell_h * len(cards)), "#d7dbe0")
    draw = ImageDraw.Draw(canvas)
    for index, (title, candidate, first, second) in enumerate(cards):
        y = index * cell_h
        label = (
            f"{title}｜第一张 vs 第{candidate['candidate_order']}张｜"
            f"exact={candidate['exact']} dHash={candidate['dhash_distance']} RMS={candidate['rms']}"
        )
        draw.text((12, y + 10), label, fill="black", font=font(21, True))
        canvas.paste(first, (20, y + 58))
        canvas.paste(second, (680, y + 58))
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, quality=90, optimize=True)
    return str(output)


def preview_package(folder: Path, base: Path, hashes: HashCache) -> dict:
    """Read only explanation-DOCX data needed by the image approval gate."""
    files = choose_files(folder)
    title = files["txt"].stem if files["txt"] else folder.name
    row = {
        "name": title,
        "folder": str(folder.relative_to(base)),
        "files": {},
        "manual_docx": {},
        "clear_errors": [],
        "candidates": [],
    }
    if files["code_pdf"]:
        path = files["code_pdf"]
        row["files"]["code_pdf"] = {"name": path.name, "sha256": hashes.sha256(path), "bytes": path.stat().st_size}
        add_code_pdf_audit(row, path, title)
    else:
        row["clear_errors"].append({"file": "代码PDF", "location": "文件夹", "message": "缺少代码PDF"})
    if files["docx"]:
        path = files["docx"]
        row["files"]["docx"] = {"name": path.name, "sha256": hashes.sha256(path), "bytes": path.stat().st_size}
        try:
            row["manual_docx"] = inspect_docx(path, title, True)
        except (zipfile.BadZipFile, KeyError, etree.XMLSyntaxError):
            pass
    if files["manual_pdf"]:
        path = files["manual_pdf"]
        row["files"]["manual_pdf"] = {"name": path.name, "sha256": hashes.sha256(path), "bytes": path.stat().st_size}
    return row
    return row


def preview_tree(base: Path, work: Path, hashes: HashCache, workers: int, visuals: bool, force: bool) -> dict:
    """Build the code-first preflight and cache the later interface approval board."""
    work.mkdir(parents=True, exist_ok=True)
    folders = locate_package_dirs(base)
    source_rows = []
    for folder in folders:
        files = choose_files(folder)
        txt = files.get("txt")
        docx = files.get("docx")
        code_pdf = files.get("code_pdf")
        source_rows.append({
            "folder": str(folder.relative_to(base)),
            "files": {
                key: {
                    "path": str(files[key].relative_to(base)),
                    "sha256": hashes.sha256(files[key]),
                } if files.get(key) else None
                for key in ("txt", "docx", "code_docx", "code_pdf", "manual_pdf")
            },
        })
    fingerprint = hashlib.sha256(json.dumps(source_rows, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
    preview_path = work / "preview.json"
    meta = load_json(work / "preview_meta.json", {})
    if not force and meta.get("version") == VERSION and meta.get("fingerprint") == fingerprint and preview_path.exists():
        cached = load_json(preview_path, {})
        required = list(cached.get("visuals", {}).get("interface_overviews", []))
        required.extend(cached.get("visuals", {}).get("interface_sheets", []))
        if not visuals or all(path and Path(path).is_file() for path in required):
            cached["cache_hit"] = True
            return cached

    packages = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = [pool.submit(preview_package, folder, base, hashes) for folder in folders]
        for future in concurrent.futures.as_completed(futures):
            packages.append(future.result())
    packages.sort(key=lambda row: row["folder"])
    visuals_out = {"interface_sheets": [], "interface_overviews": [], "code_edges": None, "pre_review_projects": []}
    if visuals:
        visual_dir = work / "visuals"
        with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, min(workers, 4))) as pool:
            values = list(pool.map(lambda row: make_interface_sheet(row, base, visual_dir), packages))
        visuals_out["interface_sheets"] = sorted(value for value in values if value)
        interface_by_folder = {
            row["folder"]: value for row, value in zip(packages, values) if value
        }
        visuals_out["pre_review_projects"] = [
            {
                "index": index,
                "program": row["name"],
                "folder": row["folder"],
                "interface_sheet": interface_by_folder.get(row["folder"]),
            }
            for index, row in enumerate(packages, 1)
        ]
        visuals_out["interface_overviews"] = make_interface_overview_groups(
            [(row["index"], row["program"], row["interface_sheet"]) for row in visuals_out["pre_review_projects"]],
            visual_dir,
        )
    result = {
        "version": VERSION,
        "created_at": now_iso(),
        "base": str(base),
        "fingerprint": fingerprint,
        "package_count": len(packages),
        "packages": [{"index": index, **row} for index, row in enumerate(packages, 1)],
        "visuals": visuals_out,
        "cache_hit": False,
    }
    atomic_json(preview_path, result)
    atomic_json(work / "preview_meta.json", {"version": VERSION, "fingerprint": fingerprint, "at": now_iso()})
    return result


def select_preview_packages(preview: dict, work: Path, include_names: list[str], include_indexes: list[int]) -> dict | None:
    packages = list(preview.get("packages", []))
    selection_path = work / "program_selection.json"
    if include_names or include_indexes:
        selected_names = set(include_names)
        known_names = {row["name"] for row in packages}
        unknown = sorted(selected_names - known_names)
        if unknown:
            raise RuntimeError("未找到获准项目：" + "、".join(unknown))
        selected_folders = {row["folder"] for row in packages if row["name"] in selected_names}
        for index in include_indexes:
            if index < 1 or index > len(packages):
                raise RuntimeError(f"项目序号超出范围：{index}；有效范围为 1-{len(packages)}")
            selected_folders.add(packages[index - 1]["folder"])
        decision = {
            "version": VERSION,
            "preview_fingerprint": preview.get("fingerprint"),
            "continue": [row["name"] for row in packages if row["folder"] in selected_folders],
            "continue_folders": [row["folder"] for row in packages if row["folder"] in selected_folders],
            "rejected": [row["name"] for row in packages if row["folder"] not in selected_folders],
            "created_at": now_iso(),
        }
        atomic_json(selection_path, decision)
        return decision
    decision = load_json(selection_path, None)
    if decision and decision.get("preview_fingerprint") == preview.get("fingerprint"):
        return decision
    return None


def tree_fingerprint(base: Path, hashes: HashCache, folders: set[str] | None = None) -> tuple[str, list[dict]]:
    rows = []
    candidates = (base / folder for folder in sorted(folders)) if folders is not None else (base,)
    paths = []
    for candidate in candidates:
        paths.extend(p for p in candidate.rglob("*") if p.is_file() and p.suffix.lower() in RELEVANT_SUFFIXES)
    for path in sorted(paths):
        rows.append({"path": str(path.relative_to(base)), "size": path.stat().st_size, "sha256": hashes.sha256(path)})
    digest = hashlib.sha256(json.dumps(rows, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
    return digest, rows


def audit_tree(
    base: Path,
    work: Path,
    hashes: HashCache,
    workers: int,
    visuals: bool,
    force: bool,
    selected_folders: set[str] | None = None,
    preview_packages: dict[str, dict] | None = None,
    preview_visuals: dict | None = None,
) -> dict:
    work.mkdir(parents=True, exist_ok=True)
    tree_digest, files = tree_fingerprint(base, hashes, selected_folders)
    code_review = load_json(work.parent / "code_reviewed.json", {}) or {}
    code_header_repairs = {str(value) for value in code_review.get("code_header_repairs", [])}
    decision_digest = hashlib.sha256(json.dumps({"code_headers": sorted(code_header_repairs)}, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
    fingerprint = hashlib.sha256(f"{tree_digest}\0{decision_digest}".encode("utf-8")).hexdigest()
    audit_path = work / "audit.json"
    meta = load_json(work / "audit_meta.json", {})
    if not force and meta.get("version") == VERSION and meta.get("fingerprint") == fingerprint and audit_path.exists():
        cached = load_json(audit_path, {})
        visual_values = cached.get("visuals", {})
        required = []
        for value in visual_values.values():
            if isinstance(value, list):
                required.extend(item for item in value if isinstance(item, str) and item)
                for item in value:
                    if isinstance(item, dict):
                        required.extend(path for path in item.values() if isinstance(path, str) and Path(path).suffix.lower() in {".jpg", ".png"})
            elif isinstance(value, str) and value:
                required.append(value)
        if not visuals or all(Path(path).exists() for path in required):
            cached["cache_hit"] = True
            return cached

    folders = locate_package_dirs(base)
    if selected_folders is not None:
        folders = [folder for folder in folders if str(folder.relative_to(base)) in selected_folders]
    packages = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = {
            pool.submit(
                inspect_package,
                folder,
                base,
                hashes,
                (preview_packages or {}).get(str(folder.relative_to(base)), {}).get("manual_docx"),
                (preview_packages or {}).get(str(folder.relative_to(base)), {}).get("code_pdf"),
                code_header_repairs,
            ): folder
            for folder in folders
        }
        for future in concurrent.futures.as_completed(futures):
            packages.append(future.result())
    packages.sort(key=lambda row: row["folder"])
    visuals_out = {"code_edges": None, "manual_sheets": [], "interface_sheets": [], "interface_overviews": [], "pre_review_projects": []}
    if visuals:
        visual_dir = work / "visuals"
        # Manual content correspondence is checked from complete DOCX text.
        # Avoid rendering every PDF page; final layout uses only processed pages 1-2.
        reusable_interfaces = {
            row.get("folder"): row.get("interface_sheet")
            for row in (preview_visuals or {}).get("pre_review_projects", [])
            if row.get("folder") and row.get("interface_sheet") and Path(row["interface_sheet"]).is_file()
        }
        interface_by_folder = dict(reusable_interfaces)
        missing_interface_packages = [row for row in packages if row["folder"] not in reusable_interfaces]
        if missing_interface_packages:
            with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, min(workers, 4))) as pool:
                values = list(pool.map(lambda row: make_interface_sheet(row, base, visual_dir), missing_interface_packages))
            interface_by_folder.update({
                row["folder"]: value for row, value in zip(missing_interface_packages, values) if value
            })
        visuals_out["interface_sheets"] = sorted(interface_by_folder.values())
        visuals_out["pre_review_projects"] = [
            {
                "index": index,
                "program": row["name"],
                "folder": row["folder"],
                "interface_sheet": interface_by_folder.get(row["folder"]),
            }
            for index, row in enumerate(packages, 1)
        ]
        visuals_out["interface_overviews"] = list((preview_visuals or {}).get("interface_overviews", []))
        visuals_out["code_edges"] = (preview_visuals or {}).get("code_edges")
        if not visuals_out["code_edges"] or not Path(visuals_out["code_edges"]).is_file():
            visuals_out["code_edges"] = make_code_board(packages, base, visual_dir / "code_edges.jpg")

    repair_plan = {
        "version": VERSION,
        "base": str(base),
        "created_at": now_iso(),
        "packages": [
            {
                "name": row["name"],
                "folder": row["folder"],
                "docx": row.get("files", {}).get("docx", {}).get("name"),
                "pdf": row.get("files", {}).get("manual_pdf", {}).get("name"),
                "actions": row["actions"],
            }
            for row in packages if row["actions"] and row.get("files", {}).get("docx")
        ],
    }
    code_repair_plan = {
        "version": VERSION,
        "base": str(base),
        "created_at": now_iso(),
        "packages": [
            {
                "name": row["name"],
                "folder": row["folder"],
                "docx": row.get("files", {}).get("code_docx", {}).get("name"),
                "pdf": row.get("files", {}).get("code_pdf", {}).get("name"),
                "actions": row.get("code_actions", []),
            }
            for row in packages if row.get("code_actions")
        ],
    }
    result = {
        "version": VERSION,
        "created_at": now_iso(),
        "base": str(base),
        "fingerprint": fingerprint,
        "file_count": len(files),
        "package_count": len(packages),
        "packages": packages,
        "repair_plan": repair_plan,
        "code_repair_plan": code_repair_plan,
        "visuals": visuals_out,
        "cache_hit": False,
    }
    atomic_json(audit_path, result)
    atomic_json(work / "repair_plan.json", repair_plan)
    atomic_json(work / "code_repair_plan.json", code_repair_plan)
    atomic_json(work / "audit_meta.json", {"version": VERSION, "fingerprint": fingerprint, "at": now_iso()})
    return result


def copy_tree_once(source: Path, destination: Path, work: Path) -> None:
    marker = work / "copy_state.json"
    if destination.exists() and any(destination.iterdir()):
        state = load_json(marker, {})
        if state.get("source") != str(source.resolve()) or state.get("destination") != str(destination.resolve()):
            raise RuntimeError("修改版已存在且不是本次流水线创建的目录；为防止覆盖，请使用新的修改目录。")
        return
    destination.parent.mkdir(parents=True, exist_ok=True)
    if os.name == "nt" and shutil.which("robocopy"):
        result = subprocess.run(
            ["robocopy", str(source), str(destination), "/E", "/COPY:DAT", "/DCOPY:T", "/R:1", "/W:1", "/MT:8", "/NFL", "/NDL", "/NJH", "/NJS", "/NP"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if result.returncode > 7:
            raise RuntimeError(f"复制修改版失败，robocopy退出码：{result.returncode}")
    else:
        shutil.copytree(source, destination, dirs_exist_ok=False)
    atomic_json(marker, {"source": str(source.resolve()), "destination": str(destination.resolve()), "at": now_iso()})


def sync_missing_files(source: Path, destination: Path, folders) -> None:
    """Copy files from source into an existing folder only when missing, so a
    reused 修改版 tree stays complete without overwriting prior repairs."""
    for relative in folders:
        src_dir = source / relative
        dst_dir = destination / relative
        if not src_dir.is_dir() or not dst_dir.is_dir():
            continue
        for src_file in src_dir.iterdir():
            if not src_file.is_file():
                continue
            name = src_file.name
            if name.startswith("~$") or name.startswith(".") or ".backup-" in name or name.endswith((".tmp", ".temp")):
                continue
            dst_file = dst_dir / src_file.name
            if not dst_file.exists():
                dst_file.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src_file, dst_file)


def copy_selected_tree_once(source: Path, destination: Path, work: Path, packages: list[dict]) -> None:
    """Create or safely extend the modified tree without overwriting prior repairs."""
    marker = work / "copy_state.json"
    folders = [str(row["folder"]) for row in packages]
    expected = {
        "source": str(source.resolve()),
        "destination": str(destination.resolve()),
        "folders": folders,
    }
    if destination.exists() and any(destination.iterdir()):
        state = load_json(marker, {})
        requested = set(folders)
        if state.get("source") != expected["source"] or state.get("destination") != expected["destination"]:
            # The modified tree already exists but this work directory has no
            # matching record. This happens when re-processing from a different
            # work directory (e.g. "从第二轮开始" on already-processed material).
            # Reuse the existing project folders and copy only the missing ones;
            # a completely unrelated 修改版 (no requested project present) is
            # still refused so foreign data is never clobbered.
            previous = {
                str(relative)
                for relative in requested
                if (destination / relative).is_dir()
            }
            if not previous:
                raise RuntimeError("修改版已存在但缺少同源处理记录；请指定新的 --modified 目录")
            for relative in sorted(requested - previous):
                src = source / relative
                dst = destination / relative
                if dst.exists():
                    raise RuntimeError(f"新增项目目录已存在但不在处理记录中：{dst}")
                if not src.is_dir():
                    raise RuntimeError(f"待处理项目目录不存在：{src}")
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copytree(src, dst, copy_function=shutil.copy2)
            sync_missing_files(source, destination, previous)
            atomic_json(marker, {**expected, "folders": sorted(previous | requested), "at": now_iso()})
            return
        previous = set(state.get("folders", []))
        for relative in sorted(requested - previous):
            src = source / relative
            dst = destination / relative
            if dst.exists():
                raise RuntimeError(f"新增项目目录已存在但不在处理记录中：{dst}")
            if not src.is_dir():
                raise RuntimeError(f"待处理项目目录不存在：{src}")
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copytree(src, dst, copy_function=shutil.copy2)
        sync_missing_files(source, destination, previous)
        atomic_json(marker, {**expected, "folders": sorted(previous | requested), "at": now_iso()})
        return
    destination.mkdir(parents=True, exist_ok=True)
    for relative in folders:
        src = source / relative
        dst = destination / relative
        if not src.is_dir():
            raise RuntimeError(f"获准项目目录不存在：{src}")
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(src, dst, copy_function=shutil.copy2)
    atomic_json(marker, {**expected, "at": now_iso()})


def select_audit_packages(audit: dict, work: Path, include_names: list[str], include_indexes: list[int]) -> tuple[dict, dict | None]:
    packages = list(audit.get("packages", []))
    selection_path = work / "program_selection.json"
    if include_names or include_indexes:
        selected_names = set(include_names)
        known_names = {row["name"] for row in packages}
        unknown = sorted(selected_names - known_names)
        if unknown:
            raise RuntimeError("未找到获准项目：" + "、".join(unknown))
        selected_folders = {row["folder"] for row in packages if row["name"] in selected_names}
        for index in include_indexes:
            if index < 1 or index > len(packages):
                raise RuntimeError(f"项目序号超出范围：{index}；有效范围为 1-{len(packages)}")
            selected_folders.add(packages[index - 1]["folder"])
        decision = {
            "version": VERSION,
            "continue": [row["name"] for row in packages if row["folder"] in selected_folders],
            "continue_folders": [row["folder"] for row in packages if row["folder"] in selected_folders],
            "rejected": [row["name"] for row in packages if row["folder"] not in selected_folders],
            "created_at": now_iso(),
        }
        atomic_json(selection_path, decision)
    else:
        decision = load_json(selection_path, None)
    if not decision:
        return audit, None
    approved_folders = set(decision.get("continue_folders", []))
    if not approved_folders:
        approved_names = set(decision.get("continue", []))
        approved_folders = {row["folder"] for row in packages if row["name"] in approved_names}
    filtered = dict(audit)
    filtered["packages"] = [row for row in packages if row["folder"] in approved_folders]
    filtered["package_count"] = len(filtered["packages"])
    plan = dict(audit.get("repair_plan", {}))
    plan["packages"] = [row for row in plan.get("packages", []) if row.get("folder") in approved_folders]
    filtered["repair_plan"] = plan
    code_plan = dict(audit.get("code_repair_plan", {}))
    code_plan["packages"] = [row for row in code_plan.get("packages", []) if row.get("folder") in approved_folders]
    filtered["code_repair_plan"] = code_plan
    return filtered, decision


def list_word_pids() -> set[int]:
    if os.name != "nt":
        return set()
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    return {int(value) for value in re.findall(r'"WINWORD\.EXE","(\d+)"', result.stdout or "", re.I)}


def terminate_process_tree(process: subprocess.Popen, word_pid_path: Path, candidate_word_pids: set[int] | None = None) -> None:
    if os.name == "nt":
        word_pid = None
        try:
            word_pid = int(word_pid_path.read_text(encoding="utf-8-sig").strip())
        except (OSError, ValueError):
            pass
        targets = {word_pid} if word_pid else set(candidate_word_pids or set())
        for target in targets:
            subprocess.run(
                ["taskkill", "/PID", str(target), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    else:
        process.kill()


def close_owned_word(word_pid_path: Path) -> None:
    """Close only the Word process recorded by this workflow."""
    try:
        word_pid = int(word_pid_path.read_text(encoding="utf-8-sig").strip())
    except (OSError, ValueError):
        return
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(word_pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )


def libreoffice_executable() -> Path | None:
    candidates = [
        shutil.which("soffice.exe"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    return None


def fallback_code_pdf_export(modified: Path, package: dict, report_path: Path, timeout_seconds: int = CODE_FALLBACK_TIMEOUT_SECONDS) -> int | None:
    """Export a code DOCX with LibreOffice only after Word timed out."""
    soffice = libreoffice_executable()
    if soffice is None:
        return None
    folder = modified / str(package.get("folder", ""))
    docx = folder / str(package.get("docx", ""))
    pdf = folder / str(package.get("pdf", ""))
    if not docx.is_file():
        return None
    output_dir = report_path.parent / "libreoffice_export" / safe_name(str(package.get("name", "code")))
    output_dir.mkdir(parents=True, exist_ok=True)
    profile_uri = (output_dir / "profile").resolve().as_uri()
    command = [
        str(soffice), "--headless", "--norestore", "--nodefault", "--nolockcheck",
        "--nofirststartwizard", f"-env:UserInstallation={profile_uri}",
        "--convert-to", "pdf", "--outdir", str(output_dir), str(docx),
    ]
    try:
        result = subprocess.run(command, text=True, encoding="utf-8", errors="replace", stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=timeout_seconds, check=False)
    except (OSError, subprocess.TimeoutExpired):
        return None
    generated = output_dir / (docx.stem + ".pdf")
    if result.returncode != 0 or not generated.is_file() or generated.stat().st_size == 0:
        return None
    temporary_target = pdf.with_name(pdf.name + ".easysoftware-copy.tmp")
    try:
        temporary_target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(generated, temporary_target)
        os.replace(temporary_target, pdf)
        pages = len(PdfReader(str(pdf)).pages)
        generated.unlink(missing_ok=True)
        return pages
    except (OSError, ValueError):
        temporary_target.unlink(missing_ok=True)
        return None


def backup_word_pair(modified: Path, row: dict, backup_dir: Path) -> tuple[Path, Path, bool]:
    folder = modified / str(row.get("folder", ""))
    docx = folder / str(row.get("docx", ""))
    pdf_name = str(row.get("pdf", ""))
    pdf = folder / pdf_name if pdf_name else docx.with_suffix(".pdf")
    if not docx.is_file():
        raise RuntimeError(f"Word源文档不存在：{docx}")
    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(docx, backup_dir / docx.name)
    had_pdf = pdf.is_file()
    if had_pdf:
        shutil.copy2(pdf, backup_dir / pdf.name)
    return docx, pdf, had_pdf


def restore_word_pair(docx: Path, pdf: Path, backup_dir: Path, had_pdf: bool) -> None:
    shutil.copy2(backup_dir / docx.name, docx)
    backup_pdf = backup_dir / pdf.name
    if had_pdf and backup_pdf.is_file():
        shutil.copy2(backup_pdf, pdf)
    elif not had_pdf and pdf.exists():
        pdf.unlink()
    lock = docx.with_name("~$" + docx.name[2:])
    if lock.exists():
        lock.unlink()


def invoke_word_pipeline(
    modified: Path,
    plan_path: Path,
    report_path: Path,
    script_path: Path,
    timeout_seconds: int = 300,
    progress: ProgressReporter | None = None,
    module: str = "",
) -> dict:
    plan = load_json(plan_path, {})
    packages = [row for row in plan.get("packages", []) if row.get("actions")]
    if not packages:
        report = {"version": VERSION, "packages": [], "message": "没有需要Word修复的文档"}
        atomic_json(report_path, report)
        return report

    # A single broken code document can make Word hang during PDF export.
    # Isolate code documents so one timeout never blocks the remaining batch.
    if module == "代码处理" and len(packages) > 1:
        item_dir = report_path.parent / "code_items"
        item_dir.mkdir(parents=True, exist_ok=True)
        combined = {"version": VERSION, "created_at": now_iso(), "base": str(modified), "packages": []}
        for index, package in enumerate(packages, 1):
            if progress:
                progress.update(module, index - 1, str(package.get("name", "")))
            item_plan = item_dir / f"plan_{index:03d}.json"
            item_report = item_dir / f"report_{index:03d}.json"
            atomic_json(item_plan, {**plan, "packages": [package]})
            item_report_data = invoke_word_pipeline(
                modified, item_plan, item_report, script_path,
                timeout_seconds=timeout_seconds, progress=None, module=module,
            )
            combined["packages"].extend(item_report_data.get("packages", []))
            if progress:
                progress.update(module, index, str(package.get("name", "")))
        atomic_json(report_path, combined)
        return combined

    shell = shutil.which("powershell.exe") or shutil.which("pwsh") or shutil.which("powershell")
    if not shell:
        raise RuntimeError("未找到PowerShell，无法运行Word批处理")
    timeout = max(timeout_seconds, timeout_seconds * len(packages))
    pid_path = report_path.with_suffix(".word.pid")
    backups: list[tuple[dict, Path, Path, Path, bool]] = []
    for index, row in enumerate(packages, 1):
        backup_dir = report_path.parent / "timeout_backups" / f"{index:03d}_{safe_name(str(row.get('name', 'unnamed')))}"
        docx, pdf, had_pdf = backup_word_pair(modified, row, backup_dir)
        backups.append((row, docx, pdf, backup_dir, had_pdf))
    command = [
        shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script_path),
        "-BaseDir", str(modified), "-PlanPath", str(plan_path), "-ReportPath", str(report_path),
        "-PidPath", str(pid_path),
    ]
    if progress:
        command.extend([
            "-ProgressPath", str(progress.path), "-ProgressTotal", str(len(packages)),
            "-ProgressPhase", module or "Word处理",
        ])
    with word_automation_lock(progress, module):
        word_pids_before = list_word_pids()
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        try:
            output, _ = process.communicate(timeout=timeout + 45)
            output = decode_process_output(output)
            completed = subprocess.CompletedProcess(command, process.returncode, output)
        except subprocess.TimeoutExpired as exc:
            terminate_process_tree(process, pid_path, list_word_pids() - word_pids_before)
            close_owned_word(pid_path)
            fallback_pages = None
            if module == "代码处理" and len(packages) == 1:
                fallback_pages = fallback_code_pdf_export(modified, packages[0], report_path)
            if fallback_pages:
                package = packages[0]
                report = {"version": VERSION, "created_at": now_iso(), "base": str(modified), "packages": [{
                    "name": package.get("name", ""), "folder": package.get("folder", ""),
                    "status": "OK", "changed": True,
                    "actions_requested": package.get("actions", []),
                    "actions_applied": ["NORMALIZE_CODE_HEADER", "EXPORT_CODE_PDF_LIBREOFFICE"],
                    "failures": [], "fallback": "Word导出超时，已从当前DOCX使用LibreOffice导出PDF",
                    "word_pages": fallback_pages, "seconds": timeout,
                }]}
                atomic_json(report_path, report)
                return report
            report = {"version": VERSION, "created_at": now_iso(), "base": str(modified), "packages": [
                {"name": row.get("name", ""), "folder": row.get("folder", ""), "status": "PRESERVED_FOR_REVIEW", "changed": True,
                 "actions_requested": row.get("actions", []), "actions_applied": [],
                 "failures": [f"Word批处理超时（>{timeout}秒），已保留当前修复结果供审核"], "word_pages": 0, "seconds": timeout}
                for row, *_ in backups
            ], "raw_output": decode_process_output(exc.stdout)[-3000:]}
            atomic_json(report_path, report)
            return report
    close_owned_word(pid_path)
    report = load_json(report_path, None)
    if report and report.get("packages"):
        report["raw_output"] = (completed.stdout or "")[-3000:]
        atomic_json(report_path, report)
        return report
    report = {"version": VERSION, "created_at": now_iso(), "base": str(modified), "packages": [
        {"name": row.get("name", ""), "folder": row.get("folder", ""), "status": "PRESERVED_FOR_REVIEW", "changed": True,
         "actions_requested": row.get("actions", []), "actions_applied": [],
         "failures": ["Word批处理未生成有效报告，已保留当前修复结果：" + (completed.stdout or "")[-500:]], "word_pages": 0, "seconds": 0}
        for row, *_ in backups
    ], "raw_output": (completed.stdout or "")[-3000:]}
    atomic_json(report_path, report)
    return report


def docx_body_digest(path: Path, title: str) -> str:
    with zipfile.ZipFile(path) as source:
        root = etree.fromstring(source.read("word/document.xml"))
    # A cover insertion legitimately shifts TOC page numbers. Exclude only
    # paragraphs styled as TOC entries; keep all ordinary body paragraphs in
    # the preservation digest.
    parts = []
    in_toc = False
    blank_after_toc = 0
    for paragraph in root.xpath("/w:document/w:body//w:p", namespaces=NS):
        style = "".join(paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)).casefold()
        paragraph_text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()
        if paragraph_text in {"目录", "目錄"}:
            in_toc = True
            blank_after_toc = 0
            continue
        if in_toc:
            if not paragraph_text:
                blank_after_toc += 1
                if blank_after_toc >= 2:
                    in_toc = False
                continue
            blank_after_toc = 0
            continue
        if style.startswith("toc"):
            continue
        parts.extend(paragraph.xpath(".//w:t/text()", namespaces=NS))
    text = "".join(parts)
    text = text.replace(title, "")
    for label in (*MANUAL_LABELS, "说明书"):
        spaced_label = r"\s*".join(re.escape(character) for character in label)
        text = re.sub(spaced_label, "", text)
    text = re.sub(r'["“”「」『』]', "", text)
    text = re.sub(r'(?:19|20)\d{2}(?:[\s,，、]*(?:19|20)\d{2})*\s*年', "目前", text)
    normalized = re.sub(r"\s+", "", text)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def docx_document_text_digest(path: Path) -> str:
    """Hash main-document text only; header-only repair must preserve it exactly."""
    with zipfile.ZipFile(path) as source:
        root = etree.fromstring(source.read("word/document.xml"))
    text = "".join(root.xpath("/w:document/w:body//w:t/text()", namespaces=NS))
    return hashlib.sha256(re.sub(r"\s+", "", text).encode("utf-8")).hexdigest()


def make_changed_manual_qa_board(cards: list[tuple[str, int, Image.Image]], output: Path) -> str | None:
    if not cards:
        return None
    pages_by_title: dict[str, list[Image.Image]] = {}
    title_order: list[str] = []
    for title, _page_number, image in cards:
        if title not in pages_by_title:
            pages_by_title[title] = []
            title_order.append(title)
        pages_by_title[title].append(image)
    columns, title_height, cell_width = 2, 44, 540
    cell_height = max(image.height for _, _, image in cards) + title_height + 10
    canvas = Image.new("RGB", (columns * cell_width, len(title_order) * cell_height), "#d7dbe0")
    draw = ImageDraw.Draw(canvas)
    for row, title in enumerate(title_order):
        top = row * cell_height
        draw.text((8, top + 8), title, fill="black", font=font(19, True))
        for column, image in enumerate(pages_by_title[title][:2]):
            canvas.paste(image, (column * cell_width + 8, top + title_height))
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, quality=90, optimize=True)
    return str(output)


def make_changed_manual_qa_board_legacy(cards: list[tuple[str, int, Image.Image]], output: Path) -> str | None:
    if not cards:
        return None
    cols, label_h = 3, 44
    cell_w = 540
    cell_h = max(image.height for _, _, image in cards) + label_h + 10
    rows = math.ceil(len(cards) / cols)
    canvas = Image.new("RGB", (cols * cell_w, rows * cell_h), "#d7dbe0")
    draw = ImageDraw.Draw(canvas)
    for index, (title, page_number, image) in enumerate(cards):
        x = (index % cols) * cell_w + 8
        y = (index // cols) * cell_h
        draw.text((x, y + 8), f"{title}｜第 {page_number} 页", fill="black", font=font(19, True))
        canvas.paste(image, (x, y + label_h))
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, quality=90, optimize=True)
    return str(output)


def verify_result(modified: Path, audit: dict, word_report: dict, code_report: dict, work: Path, hashes: HashCache, workers: int, manual_only: bool = False) -> dict:
    word_report_by_folder = {str(row.get("folder", "")): row for row in word_report.get("packages", [])}
    source_base = Path(audit.get("base", ""))
    checked = []
    qa_cards = []
    for source in audit.get("packages", []):
        folder_key = str(source.get("folder", ""))
        word_row = word_report_by_folder.get(folder_key, {
            "name": source.get("name", ""),
            "folder": folder_key,
            "actions_requested": [],
            "failures": [],
            "word_pages": 0,
        })
        name = word_row.get("name") or source.get("name", "")
        folder = modified / folder_key
        files = choose_files(folder)
        failures = list(word_row.get("failures", []))
        if files["docx"]:
            try:
                doc = inspect_docx(files["docx"], name, False)
                if not doc["cover_found"]:
                    failures.append("DOCX最终未检测到封面标题块")
                if not doc["header_compliant"]:
                    failures.append("DOCX最终缺少页眉或动态页码字段")
                source_docx_name = source.get("files", {}).get("docx", {}).get("name")
                if source_docx_name:
                    source_docx = source_base / source.get("folder", "") / source_docx_name
                    if source_docx.exists() and docx_body_digest(source_docx, name) != docx_body_digest(files["docx"], name):
                        failures.append("修复前后正文摘要不一致，存在非封面文字变化")
            except Exception as exc:
                failures.append(f"DOCX最终解析失败：{exc}")
        else:
            failures.append("修改版缺少说明DOCX")
        if files["manual_pdf"]:
            try:
                pdf = inspect_pdf(files["manual_pdf"], name)
                if not pdf["first_has_title"]:
                    failures.append("说明PDF第一页缺少软件名称")
                if not (pdf["first_has_version"] and pdf["first_has_page_number"]):
                    failures.append("说明PDF第一页缺少V1.0或页码")
                if not (pdf["last_has_title"] and pdf["last_has_version"] and pdf["last_has_page_number"]):
                    failures.append("说明PDF末页缺少规范页眉名称、V1.0或页码")
                if pdf["pages"] < 1:
                    failures.append("说明PDF页数为0")
                word_pages = int(word_row.get("word_pages") or 0)
                if word_pages and pdf["pages"] != word_pages:
                    failures.append(f"Word页数{word_pages}与说明PDF页数{pdf['pages']}不一致")
                for page_index in range(min(2, pdf["pages"])):
                    qa_cards.append((name, page_index + 1, render_pdf_page(files["manual_pdf"], page_index, 510)))
            except Exception as exc:
                failures.append(f"说明PDF最终解析失败：{exc}")
        else:
            failures.append("修改版缺少说明PDF")
        checked.append({"name": name, "folder": str(source.get("folder", word_row.get("folder", ""))), "path": str(folder), "ok": not failures, "failures": failures, "word": word_row})
    # Inspect every selected code PDF, not only documents that needed a Word
    # repair. This guarantees the final code evidence board is never blank just
    # because the original PDF was already compliant.
    code_report_by_folder = {str(row.get("folder", "")): row for row in code_report.get("packages", [])}
    code_checked = []
    code_cards = []
    for source in ([] if manual_only else audit.get("packages", [])):
        folder_key = str(source.get("folder", ""))
        name = str(source.get("name", ""))
        code_row = code_report_by_folder.get(folder_key, {})
        folder = modified / folder_key
        files = choose_files(folder)
        failures = list(code_row.get("failures", []))
        if files.get("code_docx"):
            try:
                code_header = inspect_docx_header_only(files["code_docx"], name)
                if not code_header["header_compliant"]:
                    failures.append("代码DOCX最终缺少页眉或动态页码字段")
            except Exception as exc:
                failures.append(f"代码DOCX页眉最终解析失败：{exc}")
            source_name = source.get("files", {}).get("code_docx", {}).get("name")
            if source_name:
                source_docx = source_base / source.get("folder", "") / source_name
                if source_docx.exists() and docx_document_text_digest(source_docx) != docx_document_text_digest(files["code_docx"]):
                    failures.append("代码DOCX页眉修复前后正文摘要不一致")
        else:
            failures.append("修改版缺少代码DOCX，需从DOCX重新导出代码PDF")
        if files.get("code_pdf"):
            try:
                pdf = inspect_pdf(files["code_pdf"], name, code=True)
                if not (pdf["first_has_title"] and pdf["last_has_title"]):
                    failures.append("代码PDF首尾页未同时提取到软件名称")
                if not (pdf["first_has_version"] and pdf["last_has_version"]):
                    failures.append("代码PDF首尾页未同时提取到V1.0")
                if not (pdf["first_has_page_number"] and pdf["last_has_page_number"]):
                    failures.append("代码PDF首尾页未同时提取到对应页码")
                word_pages = int(code_row.get("word_pages") or 0)
                if word_pages and pdf["pages"] != word_pages:
                    failures.append(f"代码Word页数{word_pages}与PDF页数{pdf['pages']}不一致")
                for page_index in list(dict.fromkeys((0, pdf["pages"] - 1))):
                    code_cards.append((name, page_index + 1, render_pdf_page(files["code_pdf"], page_index, 510)))
            except Exception as exc:
                failures.append(f"代码PDF最终解析失败：{exc}")
        else:
            failures.append("修改版缺少代码PDF")
        code_checked.append({"name": name, "folder": folder_key, "path": str(folder), "ok": not failures, "failures": failures, "word": code_row, "kind": "code"})
    # Keep only the repaired manual cover evidence in the second review. The
    # interface images were already judged in the first review and are omitted.
    qa_visual = make_changed_manual_qa_board(qa_cards, work / "qa_visuals" / "modified_manual_first_two_qa.jpg")
    code_qa_visual = make_changed_manual_qa_board(code_cards, work / "qa_visuals" / "modified_code_edges_qa.jpg")
    result = {
        "version": VERSION,
        "created_at": now_iso(),
        "modified": str(modified),
        "checked": checked + code_checked,
        "all_ok": all(row["ok"] for row in checked + code_checked),
        "modified_manual_qa": qa_visual,
        "modified_code_edges_qa": code_qa_visual,
        "processed_manual_first_two_sheets": [],
        "processed_manual_first_two_overview": None,
    }
    atomic_json(work / "qa.json", result)
    return result


def concise_summary(audit: dict, word_report: dict | None, code_report: dict | None, qa: dict | None, work: Path, txt_report: dict | None = None) -> dict:
    clear = []
    candidates = []
    for package in audit.get("packages", []):
        for item in package.get("clear_errors", []):
            clear.append({"program": package["name"], "folder": package["folder"], **item})
        for item in package.get("candidates", []):
            candidates.append({"program": package["name"], "folder": package["folder"], **item})
    modified = []
    txt_by_folder = {str(row.get("folder", "")): row for row in (txt_report or {}).get("packages", [])}
    word_by_folder = {str(row.get("folder", "")): row for row in (word_report or {}).get("packages", [])}
    code_by_folder = {str(row.get("folder", "")): row for row in (code_report or {}).get("packages", [])}
    if word_report:
        for row in word_report.get("packages", []):
            if row.get("changed"):
                modified.append({"program": row.get("name"), "actions": row.get("actions_applied", []), "font_size": row.get("font_size")})
    if code_report:
        for row in code_report.get("packages", []):
            if row.get("changed"):
                modified.append({"program": row.get("name"), "actions": row.get("actions_applied", []), "file": "代码DOCX/PDF"})
    qa_failures = []
    if qa:
        for row in qa.get("checked", []):
            for message in row.get("failures", []):
                qa_failures.append({"program": row["name"], "folder": row.get("folder", ""), "file": "修改后DOCX/PDF", "location": "最终质量门禁", "message": message})
    # Once final QA exists, the global report must contain only defects that
    # still exist in the repaired files. Initial findings remain in the per-
    # project comparison fields and audit.json, but are not reported as live
    # errors again.
    if qa is not None:
        clear = []
        candidates = []
    qa_by_folder: dict[str, list[dict]] = {}
    if qa:
        for row in qa.get("checked", []):
            qa_by_folder.setdefault(str(row.get("folder", "")), []).append(row)

    def original_messages(package: dict, file_key: str) -> list[str]:
        filename = str(package.get("files", {}).get(file_key, {}).get("name", ""))
        return [str(item.get("message", "")) for item in package.get("clear_errors", []) + package.get("candidates", []) if str(item.get("file", "")) == filename]

    def describe_actions(actions: list[str], labels: dict[str, str], fallback: str) -> str:
        values = [labels.get(str(action), str(action)) for action in actions]
        return "、".join(values) if values else fallback

    code_labels = {"NORMALIZE_CODE_HEADER": "修复页眉和页码", "EXPORT_CODE_PDF": "重新导出完整代码 PDF", "EXPORT_CODE_PDF_LIBREOFFICE": "Word超时后从DOCX备用导出代码 PDF"}
    code_labels["ADD_CODE_PAGE_NUMBERS"] = "为代码 PDF 全部页面添加连续页码"
    manual_labels = {
        "ADD_COVER": "补充封面", "REMOVE_DUPLICATE_COVER_TEXT": "清理重复封面",
        "REPLACE_INVALID_COVER": "删除原封面并重新生成标准封面",
        "NORMALIZE_COVER_STYLE": "整理封面格式", "NORMALIZE_HEADER": "修复页眉和页码",
        "NORMALIZE_BODY_TEXT_STYLE": "统一全部普通正文字体格式", "REMOVE_TITLE_DOUBLE_QUOTES": "删除标题双引号",
        "NORMALIZE_FIRST_POST_COVER_HEADING": "统一封面后首个标题格式",
        "REMOVE_COVER_VERSION": "封面名称去掉版本号",
        "TRIM_BODY_LEADING_BLANKS": "删除正文开头空行",
        "EXPORT_PDF": "重新导出说明书 PDF",
        "REMOVE_ALL_DOUBLE_QUOTES": "删除说明书全文双引号",
        "REPLACE_YEARS_WITH_CURRENT": "年份替换为目前",
    }
    manual_visible_actions = {
        "ADD_COVER", "REMOVE_DUPLICATE_COVER_TEXT", "REPLACE_INVALID_COVER",
        "NORMALIZE_COVER_STYLE", "NORMALIZE_HEADER", "NORMALIZE_BODY_TEXT_STYLE",
        "NORMALIZE_FIRST_POST_COVER_HEADING",
        "REMOVE_COVER_VERSION", "TRIM_BODY_LEADING_BLANKS",
        "REMOVE_TITLE_DOUBLE_QUOTES", "REMOVE_ALL_DOUBLE_QUOTES",
        "REPLACE_YEARS_WITH_CURRENT", "EXPORT_PDF",
    }
    project_results = []
    for index, package in enumerate(audit.get("packages", []), 1):
        folder = str(package.get("folder", ""))
        files = package.get("files", {})
        txt_row = txt_by_folder.get(folder, {})
        word_row = word_by_folder.get(folder, {})
        code_row = code_by_folder.get(folder, {})
        final_rows = qa_by_folder.get(folder, [])
        manual_final = next((row for row in final_rows if row.get("kind") != "code"), {})
        code_final = next((row for row in final_rows if row.get("kind") == "code"), {})
        code_original = original_messages(package, "code_pdf")
        manual_original = original_messages(package, "docx") + original_messages(package, "manual_pdf")
        has_txt = bool(files.get("txt"))
        has_code = bool(files.get("code_pdf"))
        has_manual = bool(files.get("docx") or files.get("manual_pdf"))
        project_results.append({
            "index": index,
            "name": package.get("name", ""),
            "folder": folder,
            "txt_original_valid": txt_row.get("original_valid", has_txt and not bool(package.get("txt", {}).get("clear_errors", []))),
            "txt_repairs": txt_row.get("repairs", {}),
            "txt_final": "最终检查通过" if txt_row and not txt_row.get("error") else (str(txt_row["error"]) if txt_row.get("error") else "本次未处理"),
            "code_original": "原代码检查通过" if has_code and not code_original else ("；".join(code_original) if code_original else "未提供代码 PDF"),
            "code_actions": describe_actions(list(code_row.get("actions_applied", [])), code_labels, "原代码无需修复" if has_code else "未提供代码 PDF"),
            "code_final": "最终检查通过" if code_final.get("ok", has_code) else "；".join(code_final.get("failures", [])) or "未生成最终检查结果",
            "manual_original": "原说明书检查通过" if has_manual and not manual_original else ("；".join(manual_original) if manual_original else "未提供说明书"),
            "manual_actions": describe_actions([action for action in word_row.get("actions_applied", []) if action in manual_visible_actions], manual_labels, "封面、页眉和页码无需修改" if has_manual else "未提供说明书"),
            "manual_final": "封面、页眉和页码检查通过" if manual_final.get("ok", has_manual) else "；".join(str(message) for message in manual_final.get("failures", [])) or "说明书最终检查未通过",
        })
    result = {
        "version": VERSION,
        "packages": audit.get("package_count", 0),
        "login_review_required": [],
        "modified": modified,
        "project_results": project_results,
        "clear_errors": clear + qa_failures,
        "visual_candidates": candidates,
        "visuals": {
            **audit.get("visuals", {}),
            "modified_manual_qa": qa.get("modified_manual_qa") if qa else None,
            "modified_code_edges_qa": qa.get("modified_code_edges_qa") if qa else None,
            "processed_manual_first_two_sheets": qa.get("processed_manual_first_two_sheets", []) if qa else [],
            "processed_manual_first_two_overview": qa.get("processed_manual_first_two_overview") if qa else None,
        },
        "work_dir": str(work),
    }
    atomic_json(work / "summary.json", result)
    return result


def default_work_dir(batch_root: Path) -> Path:
    digest = hashlib.sha1(str(batch_root.resolve()).lower().encode("utf-8")).hexdigest()[:16]
    return Path(tempfile.gettempdir()) / "codex-softcopyright" / digest


def main() -> int:
    parser = argparse.ArgumentParser(description="程序优先、低Token的软著材料批量审查与修复")
    parser.add_argument("batch_root", type=Path, help="批次目录：压缩包（RAR/ZIP/7z 等）或已解压的软著材料")
    parser.add_argument("--stage", choices=("preview", "audit", "repair", "verify", "all"), default="all")
    parser.add_argument("--extracted", type=Path, help="解压版目录；默认 BATCH_ROOT/解压版")
    parser.add_argument("--modified", type=Path, help="修改版目录；默认 BATCH_ROOT/修改版")
    parser.add_argument("--work-dir", type=Path, help="缓存和报告目录；默认系统临时目录中的稳定缓存")
    parser.add_argument("--input-archive", action="append", type=Path, default=[], help="仅处理指定压缩包；可重复")
    parser.add_argument("--use-existing-extracted", action="store_true", help="直接使用用户选择的已解压材料目录，不重复解压")
    parser.add_argument("--workers", type=int, default=max(2, min(6, (os.cpu_count() or 4) // 2)))
    parser.add_argument("--include", action="append", default=[], help="获准继续处理的软件名称；可重复")
    parser.add_argument("--include-index", action="append", type=int, default=[], help="获准继续处理的预览序号；可重复")
    parser.add_argument("--no-visuals", action="store_true", help="跳过拼图；仅用于程序回归测试")
    parser.add_argument("--force", action="store_true", help="忽略审计缓存")
    parser.add_argument("--manual-only", action="store_true", help="仅处理并验证说明书，跳过代码材料")
    args = parser.parse_args()

    batch_root = args.batch_root.resolve()
    extracted = (args.extracted or batch_root / "解压版").resolve()
    modified = (args.modified or batch_root / "修改版").resolve()
    work = (args.work_dir or default_work_dir(batch_root)).resolve()
    work.mkdir(parents=True, exist_ok=True)
    hashes = HashCache(work / "hash_cache.json")
    timing = {}
    audit = None
    word_report = None
    code_report = None
    qa = None
    txt_report = None
    progress = ProgressReporter(work)

    try:
        if args.stage != "verify":
            progress.update("extract/cache")
            started = time.perf_counter()
            if args.use_existing_extracted:
                if not extracted.is_dir() or not any(extracted.iterdir()):
                    raise RuntimeError(f"已解压材料目录不存在或为空：{extracted}")
                atomic_json(work / "extraction.json", {
                    "version": VERSION,
                    "source": str(extracted),
                    "target": str(extracted),
                    "mode": "USER_SELECTED_EXTRACTED",
                    "at": now_iso(),
                })
            else:
                ensure_extracted(batch_root, extracted, work, hashes, args.workers, args.input_archive)
            timing["extract_or_reuse"] = round(time.perf_counter() - started, 3)

        if args.stage == "preview":
            progress.update("interface image extraction")
            started = time.perf_counter()
            preview = preview_tree(extracted, work / "preview", hashes, args.workers, not args.no_visuals, args.force)
            progress.total = int(preview.get("package_count", 0))
            progress.update("首轮审核准备", 0, "")
            timing["preview_interfaces"] = round(time.perf_counter() - started, 3)
            timing["total_recorded"] = round(sum(timing.values()), 3)
            atomic_json(work / "timing.json", {"version": VERSION, "at": now_iso(), "stages": timing})
            hashes.save()
            progress.update("waiting for image approval", status="WAITING")
            print(json.dumps({
                "status": "NEEDS_PROJECT_SELECTION",
                "packages": preview["package_count"],
                "project_selection": str(work / "program_selection.json"),
                "modified": 0,
                "visuals": preview["visuals"],
                "timing": timing,
            }, ensure_ascii=False))
            return 0

        if args.stage in {"audit", "repair", "all"}:
            progress.update("interface image extraction")
            preview = preview_tree(extracted, work / "preview", hashes, args.workers, not args.no_visuals, args.force)
            selection = select_preview_packages(preview, work, args.include, args.include_index)
            if selection is None:
                timing["total_recorded"] = round(sum(timing.values()), 3)
                atomic_json(work / "timing.json", {"version": VERSION, "at": now_iso(), "stages": timing})
                hashes.save()
                progress.update("waiting for image approval", status="WAITING")
                print(json.dumps({
                    "status": "NEEDS_PROJECT_SELECTION",
                    "packages": preview["package_count"],
                    "project_selection": str(work / "program_selection.json"),
                    "modified": 0,
                    "visuals": preview["visuals"],
                    "timing": timing,
                }, ensure_ascii=False))
                return 0
            if not selection.get("continue_folders"):
                print(json.dumps({"status": "NOTHING_SELECTED", "packages": 0, "selection": selection}, ensure_ascii=False))
                return 0
            started = time.perf_counter()
            progress.total = len(selection["continue_folders"])
            progress.update("audit selected projects")
            audit = audit_tree(
                extracted,
                work / "audit",
                hashes,
                args.workers,
                not args.no_visuals,
                args.force,
                set(selection["continue_folders"]),
                {row["folder"]: row for row in preview.get("packages", [])},
                preview.get("visuals", {}),
            )
            timing["selected_audit_and_visuals"] = round(time.perf_counter() - started, 3)
        elif args.stage == "verify":
            audit = load_json(work / "audit" / "audit.json", None)
            if not audit:
                raise RuntimeError("没有可复用的审计结果，请先运行 --stage audit")

        if args.stage == "audit":
            summary = concise_summary(audit, None, None, None, work)
            timing["total_recorded"] = round(sum(timing.values()), 3)
            atomic_json(work / "timing.json", {"version": VERSION, "at": now_iso(), "stages": timing})
            hashes.save()
            print(json.dumps({
                "status": "OK" if not summary["clear_errors"] else "HAS_ERRORS",
                "packages": summary["packages"],
                "modified": 0,
                "clear_errors": len(summary["clear_errors"]),
                "visual_candidates": len(summary["visual_candidates"]),
                "summary": str(work / "summary.json"),
                "visuals": summary["visuals"],
                "timing": timing,
            }, ensure_ascii=False))
            return 0

        selected_plan_path = work / "audit" / "repair_plan_selected.json"
        atomic_json(selected_plan_path, audit["repair_plan"])
        selected_code_plan_path = work / "audit" / "code_repair_plan_selected.json"
        atomic_json(selected_code_plan_path, audit.get("code_repair_plan", {"version": VERSION, "packages": []}))

        if args.stage in {"repair", "all"}:
            started = time.perf_counter()
            progress.update("复制项目", 0, "")
            copy_selected_tree_once(extracted, modified, work, audit["packages"])
            timing["copy_or_reuse"] = round(time.perf_counter() - started, 3)
            started = time.perf_counter()
            txt_report = repair_selected_txt(modified, audit["packages"], work, progress)
            timing["txt_repair"] = round(time.perf_counter() - started, 3)
            started = time.perf_counter()
            progress.update("说明书处理", 0, "")
            plan_path = selected_plan_path
            script_path = Path(__file__).with_name("word_manual_pipeline_stable.ps1")
            word_report = invoke_word_pipeline(modified, plan_path, work / "word_report.json", script_path, progress=progress, module="说明书处理")
            timing["word_repair_and_export"] = round(time.perf_counter() - started, 3)
            if args.manual_only:
                code_report = {"version": VERSION, "packages": [], "skipped": True}
                atomic_json(work / "code_header_report.json", code_report)
            else:
                started = time.perf_counter()
                progress.update("代码处理", 0, "")
                code_plan = audit.get("code_repair_plan", {"version": VERSION, "packages": []})
                word_packages = []
                missing_docx = []
                for package in code_plan.get("packages", []):
                    if package.get("docx"):
                        word_packages.append(package)
                    else:
                        missing_docx.append({
                            "name": package.get("name", ""), "folder": package.get("folder", ""),
                            "status": "PRESERVED_FOR_REVIEW", "changed": False,
                            "actions_requested": package.get("actions", []), "actions_applied": [],
                            "failures": ["代码PDF缺少配套DOCX，按规则未直接修改PDF"],
                            "word_pages": 0, "seconds": 0,
                        })
                atomic_json(selected_code_plan_path, {**code_plan, "packages": word_packages})
                code_script_path = Path(__file__).with_name("repair_code_headers.ps1")
                code_report = invoke_word_pipeline(modified, selected_code_plan_path, work / "code_header_report.json", code_script_path, timeout_seconds=CODE_WORD_TIMEOUT_SECONDS, progress=progress, module="代码处理")
                code_report["packages"] = code_report.get("packages", []) + missing_docx
                atomic_json(work / "code_header_report.json", code_report)
                timing["code_header_repair_and_export"] = round(time.perf_counter() - started, 3)
        elif args.stage == "verify":
            word_report = load_json(work / "word_report.json", None)
            if word_report is None:
                raise RuntimeError("没有Word修复报告，请先运行 --stage repair")
            code_report = load_json(work / "code_header_report.json", {"version": VERSION, "packages": []})

        if args.stage in {"verify", "all"}:
            started = time.perf_counter()
            progress.update("final QA")
            qa = verify_result(modified, audit, word_report or {"packages": []}, code_report or {"packages": []}, work, hashes, args.workers, args.manual_only)
            timing["verify"] = round(time.perf_counter() - started, 3)

        summary = concise_summary(audit, word_report, code_report, qa, work, txt_report)
        timing["total_recorded"] = round(sum(timing.values()), 3)
        atomic_json(work / "timing.json", {"version": VERSION, "at": now_iso(), "stages": timing})
        hashes.save()
        progress.update("final result review", status="WAITING")
        print(json.dumps({
            "status": "OK" if not summary["clear_errors"] else "HAS_ERRORS",
            "packages": summary["packages"],
            "login_review_required": summary["login_review_required"],
            "modified": len(summary["modified"]),
            "clear_errors": len(summary["clear_errors"]),
            "visual_candidates": len(summary["visual_candidates"]),
            "summary": str(work / "summary.json"),
            "visuals": summary["visuals"],
            "timing": timing,
        }, ensure_ascii=False))
        return 0
    except Exception as exc:
        hashes.save()
        progress.update("failed", status="FAILED", error=str(exc))
        atomic_json(work / "fatal_error.json", {"version": VERSION, "at": now_iso(), "error": str(exc)})
        print(json.dumps({"status": "FAILED", "error": str(exc), "work_dir": str(work)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

"""Offline Chinese-to-English translation for source-code comments."""
from __future__ import annotations

import atexit
import json
import os
import re
import socket
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path


MODEL_NAME = "QwenPaw-Flash-2B-Q4_K_M.gguf"
HAN_RUN = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]+")
_SERVER: subprocess.Popen | None = None
_BASE_URL: str | None = None
_SERVER_STDERR = None
_PATH_MAPPINGS: list[str] = []
_PATH_MUTEX = None
_SERVER_LOCK = threading.RLock()


def _runtime(app_root: Path) -> tuple[Path, Path]:
    configured = os.environ.get("EASYSOFTWARE_LLAMA_HOME")
    candidates = [Path(configured)] if configured else []
    candidates.extend([
        app_root / "llama",
        app_root / "models" / "llama",
        Path(r"D:\all\ai_agent_about\llama"),
    ])
    for directory in candidates:
        server = directory / "llama-server.exe"
        model_candidates = [
            directory / MODEL_NAME,
            directory / "models" / MODEL_NAME,
            directory.parent / "model" / MODEL_NAME,
            directory.parent / "models" / MODEL_NAME,
            Path(r"D:\all\ai_agent_about\model") / MODEL_NAME,
        ]
        model = next((path for path in model_candidates if path.is_file()), None)
        if server.is_file() and model:
            return server, model
    raise RuntimeError(
        "找不到离线翻译模型。请将 llama.cpp 与 QwenPaw 模型放入软件 models/llama，"
        "或设置 EASYSOFTWARE_LLAMA_HOME。"
    )


def _stop_server() -> None:
    global _SERVER, _BASE_URL, _SERVER_STDERR, _PATH_MAPPINGS, _PATH_MUTEX
    with _SERVER_LOCK:
        try:
            if _SERVER and _SERVER.poll() is None:
                _SERVER.terminate()
                try:
                    _SERVER.wait(timeout=8)
                except subprocess.TimeoutExpired:
                    _SERVER.kill()
                    _SERVER.wait(timeout=5)
        finally:
            _SERVER = None
            _BASE_URL = None
            if _SERVER_STDERR is not None:
                try:
                    _SERVER_STDERR.close()
                except OSError:
                    pass
                _SERVER_STDERR = None
            mappings = _PATH_MAPPINGS
            _PATH_MAPPINGS = []
            for drive in reversed(mappings):
                try:
                    subprocess.run(
                        ["subst", drive, "/D"], stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL, check=False, timeout=5,
                    )
                except Exception:
                    pass
            try:
                _release_path_mutex()
            except Exception:
                pass


def _release_path_mutex() -> None:
    global _PATH_MUTEX
    handle = _PATH_MUTEX
    _PATH_MUTEX = None
    if handle is None:
        return
    import ctypes

    kernel32 = ctypes.windll.kernel32
    try:
        kernel32.ReleaseMutex(handle)
    finally:
        kernel32.CloseHandle(handle)


def _acquire_path_mutex() -> None:
    """Serialize temporary SUBST allocation across this program's processes."""
    global _PATH_MUTEX
    import ctypes

    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
    kernel32.WaitForSingleObject.restype = ctypes.c_uint32
    kernel32.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.ReleaseMutex.argtypes = [ctypes.c_void_p]
    handle = kernel32.CreateMutexW(None, False, "Local\\EasySoftwareLlamaPathBridge")
    if not handle:
        raise OSError("无法创建本地模型路径映射锁")
    result = kernel32.WaitForSingleObject(handle, 10000)
    if result not in (0, 0x80):  # WAIT_OBJECT_0 / WAIT_ABANDONED
        kernel32.CloseHandle(handle)
        raise RuntimeError("本地模型路径映射正被另一个 EasySoftware 进程占用")
    _PATH_MUTEX = handle


def _get_logical_drives() -> int:
    import ctypes

    return ctypes.windll.kernel32.GetLogicalDrives()


def _ascii_runtime_paths(server: Path, model: Path) -> tuple[Path, Path]:
    """Return ASCII SUBST aliases for Unicode Windows runtime/model paths."""
    global _PATH_MAPPINGS
    if os.name != "nt" or all(str(path).isascii() for path in (server, model, server.parent)):
        return server, model

    _acquire_path_mutex()
    try:
        used_drives = _get_logical_drives()
        aliases: dict[str, str] = {}
        mapped_paths: list[Path] = []
        errors: list[str] = []
        for path in (server, model):
            parent = os.path.normcase(os.path.abspath(path.parent))
            drive = aliases.get(parent)
            if drive is None:
                for letter in reversed("DEFGHIJKLMNOPQRSTUVWXYZ"):
                    if used_drives & (1 << (ord(letter) - ord("A"))):
                        continue
                    candidate = f"{letter}:"
                    result = subprocess.run(
                        ["subst", candidate, str(path.parent)], capture_output=True,
                        text=True, check=False, timeout=10,
                    )
                    if result.returncode == 0:
                        drive = candidate
                        aliases[parent] = drive
                        _PATH_MAPPINGS.append(drive)
                        used_drives |= 1 << (ord(letter) - ord("A"))
                        break
                    errors.append((result.stderr or result.stdout or "").strip())
                if drive is None:
                    detail = next((item for item in reversed(errors) if item), "没有可用盘符")
                    raise RuntimeError(f"无法为本地翻译模型创建临时路径映射：{detail}")
            mapped_paths.append(Path(f"{drive}\\{path.name}"))
        if not all(str(path).isascii() for path in mapped_paths):
            raise RuntimeError("临时模型路径仍包含非 ASCII 字符")
        return mapped_paths[0], mapped_paths[1]
    except BaseException:
        _stop_server()
        raise


def _server_error_tail() -> str:
    global _SERVER_STDERR
    if _SERVER_STDERR is None:
        return ""
    _SERVER_STDERR.flush()
    _SERVER_STDERR.seek(0)
    content = _SERVER_STDERR.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    return str(content)[-2000:].strip()


atexit.register(_stop_server)


def _ensure_server(app_root: Path) -> str:
    global _SERVER, _BASE_URL, _SERVER_STDERR
    with _SERVER_LOCK:
        if _SERVER and _SERVER.poll() is None and _BASE_URL:
            return _BASE_URL
        _stop_server()
        server, model = _runtime(app_root)
        server_path, model_path = _ascii_runtime_paths(server, model)
        try:
            with socket.socket() as sock:
                sock.bind(("127.0.0.1", 0))
                port = sock.getsockname()[1]
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            _SERVER_STDERR = tempfile.TemporaryFile(mode="w+b")
            _SERVER = subprocess.Popen(
                [str(server_path), "--model", str(model_path), "--host", "127.0.0.1", "--port", str(port),
                 "--ctx-size", "4096", "--n-gpu-layers", "99", "--alias", "local-comment-translator"],
                cwd=str(server_path.parent), stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                stderr=_SERVER_STDERR, creationflags=flags,
            )
        except BaseException:
            _stop_server()
            raise
        _BASE_URL = f"http://127.0.0.1:{port}"
        deadline = time.monotonic() + 150
        while time.monotonic() < deadline:
            if _SERVER.poll() is not None:
                detail = _server_error_tail()
                _stop_server()
                suffix = f"；运行日志：{detail}" if detail else "；请检查模型文件和显存/内存"
                raise RuntimeError("本地翻译模型启动失败" + suffix)
            try:
                with urllib.request.urlopen(_BASE_URL + "/health", timeout=2) as response:
                    if response.status == 200:
                        return _BASE_URL
            except (OSError, urllib.error.URLError):
                time.sleep(1)
        detail = _server_error_tail()
        _stop_server()
        suffix = f"；运行日志：{detail}" if detail else ""
        raise RuntimeError("本地翻译模型启动超时" + suffix)


def translate_comment(text: str, *, app_root: Path, target_chars: int | None = None) -> str:
    """Translate one comment locally. No external network requests are made."""
    if "\n" in text or "\r" in text:
        translated_lines: list[str] = []
        lines = text.splitlines(keepends=True)
        total_content = max(1, sum(len(line.strip()) for line in lines))
        for line in lines:
            ending = "\r\n" if line.endswith("\r\n") else "\n" if line.endswith("\n") else "\r" if line.endswith("\r") else ""
            content = line[:-len(ending)] if ending else line
            leading = content[:len(content) - len(content.lstrip())]
            trailing = content[len(content.rstrip()):]
            core = content[len(leading):len(content) - len(trailing) if trailing else len(content)]
            if core and any("\u3400" <= ch <= "\u9fff" for ch in core):
                line_target = max(4, round(target_chars * len(core.strip()) / total_content)) if target_chars else None
                core = translate_comment(core, app_root=app_root, target_chars=line_target)
            translated_lines.append(leading + core + trailing + ending)
        return "".join(translated_lines)
    base_url = _ensure_server(app_root)
    if target_chars:
        length_note = (
            f" Rewrite the English result to no more than {target_chars} characters, keeping only essential meaning."
            if target_chars < len(text) * 3
            else f" Aim for about {target_chars} characters by adding only useful details supported by the comment."
        )
    else:
        length_note = " Keep it concise and faithful."
    prompt = (
        "Translate this source-code comment from Chinese to clear English. Preserve identifiers, "
        "paths, commands, numbers, markup, syntax, and technical meaning. Translate Chinese text inside "
        "example string literals too, while keeping their quotes and code syntax. Output only the translated comment, "
        "with no quotes or explanation. Do not invent behavior."
        + length_note
        + "\n\nComment:\n" + text
    )
    body = json.dumps({
        "model": "local-comment-translator",
        "messages": [
            {"role": "system", "content": "You are a source-code comment translator. Always translate all Chinese into English. Output only English text and never explain."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.1,
        "max_tokens": max(128, min(2048, (target_chars or len(text) * 4) + 96)),
        "chat_template_kwargs": {"enable_thinking": False},
        "stream": False,
    }).encode("utf-8")
    request = urllib.request.Request(
        base_url + "/v1/chat/completions", data=body,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(request, timeout=120) as response:
        result = json.loads(response.read().decode("utf-8"))
    translated = str(result["choices"][0]["message"]["content"]).strip().strip("` \t\r\n\"'")
    if HAN_RUN.search(translated):
        retry = prompt + "\nEvery Chinese word must be translated into English; do not leave any Chinese characters in the output."
        body = json.dumps({
            "model": "local-comment-translator",
            "messages": [
                {"role": "system", "content": "Translate, do not answer the comment. Convert every Chinese character and Chinese phrase to English. The final output must contain zero Chinese characters."},
                {"role": "user", "content": retry},
            ],
            "temperature": 0,
            "max_tokens": max(192, min(2048, (target_chars or len(text) * 4) + 128)),
            "chat_template_kwargs": {"enable_thinking": False},
            "stream": False,
        }).encode("utf-8")
        request = urllib.request.Request(base_url + "/v1/chat/completions", data=body, headers={"Content-Type": "application/json"}, method="POST")
        with urllib.request.urlopen(request, timeout=120) as response:
            result = json.loads(response.read().decode("utf-8"))
        translated = str(result["choices"][0]["message"]["content"]).strip().strip("` \t\r\n\"'")
    if not translated or HAN_RUN.search(translated):
        pieces: list[str] = []
        cursor = 0
        for match in HAN_RUN.finditer(translated):
            pieces.append(translated[cursor:match.start()])
            segment_prompt = (
                "Translate this Chinese phrase into English for a source-code comment. "
                "Output only the English phrase. Phrase: " + match.group(0)
                + "\nComment context: " + text[:500]
            )
            segment_body = json.dumps({
                "model": "local-comment-translator",
                "messages": [
                    {"role": "system", "content": "You are a Chinese-to-English translator. Never repeat Chinese characters. Output only English."},
                    {"role": "user", "content": segment_prompt},
                ],
                "temperature": 0,
                "max_tokens": 96,
                "chat_template_kwargs": {"enable_thinking": False},
                "stream": False,
            }).encode("utf-8")
            segment_request = urllib.request.Request(base_url + "/v1/chat/completions", data=segment_body, headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(segment_request, timeout=120) as response:
                segment_result = json.loads(response.read().decode("utf-8"))
            english = str(segment_result["choices"][0]["message"]["content"]).strip().strip("` \t\r\n\"'")
            if not english or HAN_RUN.search(english):
                raise RuntimeError("本地模型无法完整翻译该代码注释")
            pieces.append(english)
            cursor = match.end()
        pieces.append(translated[cursor:])
        translated = "".join(pieces)
        if not translated or HAN_RUN.search(translated):
            raise RuntimeError("本地模型无法完整翻译该代码注释")
    return translated

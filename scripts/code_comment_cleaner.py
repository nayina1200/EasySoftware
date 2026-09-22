from __future__ import annotations

import ast
import hashlib
import html
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import textwrap
import time
import tokenize
import zipfile
import contextlib
import ctypes
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
HAN = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\U00020000-\U0002fa1f\U00030000-\U0003134f]")
CODE_DOCX = re.compile(r"代码\s*\.docx$", re.IGNORECASE)
CODE_PDF = re.compile(r"代码\s*\.pdf$", re.IGNORECASE)
RULE_VERSION = "code-comments-v2"
WORD_AUTOMATION_MUTEX = "Local\\EasySoftware.WordAutomation.v1"


@dataclass
class Finding:
    paragraph: int
    kind: str
    original: str
    result: str
    reason: str = ""


def contains_han(value: str) -> bool:
    return bool(HAN.search(value))


def _regex_can_start(previous: str) -> bool:
    return not previous or previous in "=([{!,:;?&|+-*%^~<>"


def _inside_css_url(line_prefix: str) -> bool:
    opened = [match.start() for match in re.finditer(r"url\s*\(", line_prefix, re.IGNORECASE)]
    return bool(opened and line_prefix.rfind(")") < opened[-1])


def _python_floor_division(text: str, position: int, line_end: int) -> tuple[bool, bool]:
    """Return (is_operator, is_ambiguous) for a possible Python // token."""
    line_start = text.rfind("\n", 0, position) + 1
    line_text = text[line_start:line_end]
    column = position - line_start
    before = line_text[:column].rstrip()
    after = line_text[column + 2:]
    if not before or not re.search(r"[\w\]\)]$", before):
        return False, False
    try:
        tokens = list(tokenize.generate_tokens(io.StringIO(line_text).readline))
    except (IndentationError, tokenize.TokenError):
        return False, False
    operator = next((token for token in tokens if token.start == (1, column) and token.type == tokenize.OP and token.string in {"//", "//="}), None)
    if operator is None:
        return False, False
    python_comment = next((token for token in tokens if token.type == tokenize.COMMENT and token.start[1] > column), None)
    code_part = line_text[:python_comment.start[1]] if python_comment else line_text
    context = text[max(0, line_start - 1200):min(len(text), line_end + 1200)]
    javascript_context = bool(re.search(r"(?:\b(?:const|let|var|function|new)\b|=>|\bthis\.)", context))
    if python_comment is None and javascript_context:
        return False, False
    try:
        ast.parse(textwrap.dedent(code_part))
    except SyntaxError:
        return False, False
    if python_comment is None and contains_han(after):
        return False, True
    return True, False


def _decrement_operator(text: str, position: int, line_end: int) -> bool:
    line_start = text.rfind("\n", 0, position) + 1
    before = text[line_start:position]
    after = text[position + 2:line_end]
    # Conventional postfix forms: count--; i-- /* ... */; value--)
    if re.search(r"[A-Za-z0-9_\]\)]\s*$", before) and re.match(r"\s*(?:[;,.?\])}:]|//|/\*)", after):
        return True
    # Conventional prefix forms: --count. A separating space is ambiguous with
    # an SQL comment and is intentionally not classified as an operator.
    if (not before.strip() or re.search(r"[=({[,!?:;+*/%&|^~<>-]\s*$", before)) and re.match(r"[A-Za-z_]\w*", after):
        return True
    return False


def _is_python_hash_comment(text: str, position: int, line_end: int) -> bool:
    line_start = text.rfind("\n", 0, position) + 1
    before = text[line_start:position]
    after = text[position + 1:line_end]
    if not before.strip():
        # A leading '#name' can be a CSS id selector. Only concrete selector
        # syntax on this line is strong enough to protect it; otherwise Python's
        # legal no-space comment form wins.
        if not after or after[:1].isspace():
            return True
        return not bool(re.search(r"[\{,]", after))
    if _inside_css_url(before):
        return False
    # CSS declarations may contain hashes that are not valid hex colours too
    # (for example generated identifiers). They are values, not comments.
    if re.fullmatch(r"\s*[-\w]+\s*:\s*", before):
        return False
    if after and not after[:1].isspace():
        # div#id is strong CSS selector evidence. An assignment or expression
        # followed by #text is a Python comment even without spaces.
        return not bool(re.fullmatch(r"\s*[A-Za-z_][\w.-]*", before))
    return True


def _code_files(root: Path) -> tuple[list[Path], list[Path]]:
    files = [path for path in root.rglob("*") if path.is_file()]
    docx_files = sorted((path for path in files if CODE_DOCX.search(path.name)), key=lambda path: str(path).casefold())
    pdf_files = sorted((path for path in files if CODE_PDF.search(path.name)), key=lambda path: str(path).casefold())
    return docx_files, pdf_files


def _pair_key(path: Path) -> tuple[str, str]:
    return str(path.parent).casefold(), path.stem.strip().casefold()


def _document_rows(root: Path, inspect: bool) -> list[dict]:
    docx_files, pdf_files = _code_files(root)
    pdf_by_key = {_pair_key(path): path for path in pdf_files}
    docx_keys = {_pair_key(path) for path in docx_files}
    rows: list[dict] = []
    for docx in docx_files:
        pdf = pdf_by_key.get(_pair_key(docx))
        if pdf is None:
            rows.append({"docx": str(docx), "status": "SKIPPED_MISSING_PDF", "deletions": 0, "error": "缺少同名代码 PDF，未处理", "findings": []})
        elif inspect:
            rows.append(inspect_docx(docx))
        else:
            rows.append({"docx": str(docx), "pdf": str(pdf), "status": "READY", "deletions": 0, "findings": []})
    for pdf in pdf_files:
        if _pair_key(pdf) not in docx_keys:
            rows.append({"docx": "", "pdf": str(pdf), "status": "SKIPPED_MISSING_DOCX", "deletions": 0, "error": "缺少同名代码 DOCX，未处理", "findings": []})
    if not docx_files and not pdf_files:
        rows.append({"docx": "", "status": "MISSING_CODE_DOCX", "deletions": 0, "error": "未找到代码 DOCX，未处理", "findings": []})
    return rows


def find_comment_ranges(text: str) -> tuple[list[tuple[int, int, str]], list[Finding]]:
    """Find closed comments containing Han characters without touching literals.

    The lexer deliberately accepts the union of Java/JavaScript/CSS, Python,
    HTML and SQL comments. Ambiguous unclosed constructs are retained and
    reported instead of being guessed.
    """
    ranges: list[tuple[int, int, str]] = []
    uncertain: list[Finding] = []
    i, length = 0, len(text)
    line = 1
    while i < length:
        ch = text[i]
        if ch == "\n":
            line += 1
            i += 1
            continue

        # Python triple-quoted strings/docstrings are protected as literals.
        if text.startswith("'''", i) or text.startswith('\"\"\"', i):
            delimiter = text[i:i + 3]
            end = text.find(delimiter, i + 3)
            if end < 0:
                uncertain.append(Finding(line, "uncertain", text[i:i + 160], "retained", "未闭合的三引号内容"))
                break
            line += text.count("\n", i, end + 3)
            i = end + 3
            continue

        # Quoted and template strings. Escapes protect the following character.
        if ch in "'\"`":
            delimiter = ch
            start_line = line
            i += 1
            closed = False
            while i < length:
                if text[i] == "\\":
                    i += 2
                    continue
                if text[i] == "\n":
                    line += 1
                    if delimiter != "`":
                        break
                if text[i] == delimiter:
                    i += 1
                    closed = True
                    break
                i += 1
            if not closed and delimiter == "`":
                uncertain.append(Finding(start_line, "uncertain", text[max(0, i - 160):i], "retained", "未闭合的模板字符串"))
            continue

        # JavaScript regular-expression literals are protected. This is only
        # entered in expression-start positions, so division remains ordinary.
        if ch == "/" and not text.startswith(("//", "/*"), i):
            previous = text[:i].rstrip()[-1:] 
            if _regex_can_start(previous):
                j, in_class = i + 1, False
                while j < length and text[j] != "\n":
                    if text[j] == "\\":
                        j += 2
                        continue
                    if text[j] == "[":
                        in_class = True
                    elif text[j] == "]":
                        in_class = False
                    elif text[j] == "/" and not in_class:
                        j += 1
                        while j < length and text[j].isalpha():
                            j += 1
                        i = j
                        break
                    j += 1
                else:
                    i += 1
                continue

        start = i
        kind = ""
        end = -1
        if text.startswith("<!--", i):
            kind = "html"
            end_marker = "-->"
            marker_length = 4
        elif text.startswith("/*", i):
            kind = "block"
            end_marker = "*/"
            marker_length = 2
        else:
            end_marker = ""
            marker_length = 0

        if end_marker:
            close = text.find(end_marker, i + marker_length)
            if close < 0:
                uncertain.append(Finding(line, "uncertain", text[i:i + 200], "retained", f"未闭合的{kind}注释"))
                break
            end = close + len(end_marker)
        elif text.startswith("//", i):
            line_prefix = text[text.rfind("\n", 0, i) + 1:i]
            if _inside_css_url(line_prefix):
                close = text.find(")", i + 2)
                i = length if close < 0 else close + 1
                continue
            line_end = text.find("\n", i)
            line_end = length if line_end < 0 else line_end
            is_floor_division, is_ambiguous = _python_floor_division(text, i, line_end)
            if is_floor_division:
                i += 3 if text.startswith("//=", i) else 2
                continue
            if is_ambiguous:
                uncertain.append(Finding(line, "uncertain", text[i:line_end], "retained", "// 可能是注释或 Python 整除运算符"))
                i = line_end
                continue
            kind = "line"
            end = line_end
        elif ch == "#":
            tail_end = text.find("\n", i)
            tail_end = length if tail_end < 0 else tail_end
            if _is_python_hash_comment(text, i, tail_end):
                kind = "python"
                end = tail_end
        elif text.startswith("--", i):
            before = text[text.rfind("\n", 0, i) + 1:i]
            tail_end = text.find("\n", i)
            tail_end = length if tail_end < 0 else tail_end
            after = text[i + 2:tail_end]
            css_custom_property = (not before.strip() and bool(re.match(r"[^\s:]+\s*:", after))) or bool(re.search(r"var\s*\(\s*$", before, re.IGNORECASE))
            if _decrement_operator(text, i, tail_end):
                i += 2
                continue
            if not css_custom_property:
                kind = "sql"
                end = tail_end

        if end >= 0:
            value = text[start:end]
            if contains_han(value):
                ranges.append((start, end, kind))
            line += text.count("\n", start, end)
            i = end
            continue
        i += 1
    return ranges, uncertain


def _paragraphs(root: etree._Element) -> list[etree._Element]:
    return root.xpath("//w:body//w:p", namespaces=NS)


def _paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def _set_paragraph_text(paragraph: etree._Element, value: str) -> None:
    nodes = paragraph.xpath(".//w:t", namespaces=NS)
    if not nodes:
        return
    remaining = value
    for node in nodes:
        capacity = len(node.text or "")
        node.text, remaining = remaining[:capacity], remaining[capacity:]
        if node.text[:1].isspace() or node.text[-1:].isspace():
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    if remaining:
        nodes[-1].text = (nodes[-1].text or "") + remaining


def _must_retain_paragraph(paragraph: etree._Element) -> bool:
    if paragraph.xpath("./w:pPr/w:sectPr", namespaces=NS):
        return True
    parent = paragraph.getparent()
    return bool(parent is not None and parent.tag == f"{{{W_NS}}}tc" and len(parent.xpath("./w:p", namespaces=NS)) <= 1)


def analyze_document_xml(xml: bytes) -> tuple[etree._Element, list[Finding], int]:
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    root = etree.fromstring(xml, parser)
    paragraphs = _paragraphs(root)
    texts = [_paragraph_text(p) for p in paragraphs]
    joined = "\n".join(texts)
    ranges, uncertain = find_comment_ranges(joined)
    findings: list[Finding] = list(uncertain)
    offsets: list[int] = []
    cursor = 0
    for value in texts:
        offsets.append(cursor)
        cursor += len(value) + 1
    by_paragraph: dict[int, list[tuple[int, int, str]]] = {}
    for start, end, kind in ranges:
        for index, (offset, value) in enumerate(zip(offsets, texts)):
            local_start = max(0, start - offset)
            local_end = min(len(value), end - offset)
            if local_start < local_end:
                by_paragraph.setdefault(index, []).append((local_start, local_end, kind))
    changed = 0
    for index, paragraph in enumerate(paragraphs):
        original = texts[index]
        spans = sorted(by_paragraph.get(index, []), reverse=True)
        cleaned = original
        removed_values = []
        for start, end, kind in spans:
            removed_values.append(cleaned[start:end])
            cleaned = cleaned[:start] + cleaned[end:]
        cleaned = cleaned.rstrip() if spans else cleaned
        if cleaned != original:
            changed += len(spans)
            findings.append(Finding(index + 1, "deleted", " | ".join(reversed(removed_values)), cleaned, "注释含汉字"))
            if not cleaned.strip():
                parent = paragraph.getparent()
                if parent is not None:
                    if _must_retain_paragraph(paragraph):
                        _set_paragraph_text(paragraph, "")
                        continue
                    previous = paragraph.getprevious()
                    following = paragraph.getnext()
                    parent.remove(paragraph)
                    if (
                        previous is not None and following is not None
                        and previous.tag == f"{{{W_NS}}}p" and following.tag == f"{{{W_NS}}}p"
                        and not _paragraph_text(previous).strip() and not _paragraph_text(following).strip()
                    ):
                        parent.remove(following)
                continue
            _set_paragraph_text(paragraph, cleaned)
    return root, sorted(findings, key=lambda item: item.paragraph), changed


def inspect_docx(path: Path) -> dict:
    try:
        with zipfile.ZipFile(path) as archive:
            xml = archive.read("word/document.xml")
        _, findings, changed = analyze_document_xml(xml)
        return {"docx": str(path), "status": "READY", "deletions": changed, "findings": [asdict(item) for item in findings]}
    except Exception as exc:
        return {"docx": str(path), "status": "FAILED", "deletions": 0, "error": str(exc), "findings": []}


def clean_docx(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
        root, findings, changed = analyze_document_xml(xml)
        output_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        members = [(item, output_xml if item.filename == "word/document.xml" else archive.read(item.filename)) for item in archive.infolist()]
    if changed:
        fd, temp_name = tempfile.mkstemp(prefix=path.stem + "-", suffix=".docx", dir=path.parent)
        os.close(fd)
        temp = Path(temp_name)
        try:
            with zipfile.ZipFile(temp, "w") as output:
                for item, content in members:
                    output.writestr(item, content)
            with zipfile.ZipFile(temp) as check:
                check.testzip()
                etree.fromstring(check.read("word/document.xml"))
            os.replace(temp, path)
        finally:
            temp.unlink(missing_ok=True)
    return {"docx": str(path), "status": "CLEANED", "deletions": changed, "findings": [asdict(item) for item in findings]}


def _find_7z(app_root: Path) -> Path:
    candidates = [app_root / "tools/7z/7z.exe", Path(shutil.which("7z.exe") or ""), Path(shutil.which("7z") or "")]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise RuntimeError("未找到 7-Zip，无法解压 RAR")


def _validate_archive_members(seven_zip: Path, archive: Path) -> None:
    result = subprocess.run([str(seven_zip), "l", "-slt", str(archive)], text=True, encoding="utf-8", errors="replace", stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if result.returncode != 0:
        raise RuntimeError(f"无法读取 RAR 目录：{archive.name}")
    member_listing = result.stdout.split("----------", 1)[-1] if "----------" in result.stdout else ""
    for match in re.finditer(r"(?m)^Path = (.+)$", member_listing):
        value = match.group(1).strip()
        if not value:
            continue
        normalized = value.replace("\\", "/")
        if normalized.startswith("/") or re.match(r"^[A-Za-z]:", value) or ".." in Path(normalized).parts:
            raise ValueError(f"压缩包包含不安全路径：{value}")


def _validate_extracted_tree(target: Path) -> None:
    root = target.resolve()
    for path in target.rglob("*"):
        attributes = getattr(path.lstat(), "st_file_attributes", 0)
        if path.is_symlink() or attributes & 0x400:
            raise ValueError(f"解压结果包含链接或重解析点：{path}")
        resolved = path.resolve()
        if resolved != root and root not in resolved.parents:
            raise ValueError(f"解压结果越界：{path}")


def _archive_key(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def extract_cached(archive: Path, work: Path, app_root: Path) -> Path:
    key = _archive_key(archive)
    target = work / "extracted" / key
    marker = target / ".complete"
    if marker.is_file():
        _validate_archive_members(_find_7z(app_root), archive)
        _validate_extracted_tree(target)
        return target
    temporary = target.with_name(target.name + ".tmp")
    if temporary.exists():
        shutil.rmtree(temporary)
    temporary.mkdir(parents=True)
    seven_zip = _find_7z(app_root)
    _validate_archive_members(seven_zip, archive)
    command = [str(seven_zip), "x", "-y", f"-o{temporary}", str(archive)]
    result = subprocess.run(command, text=True, encoding="utf-8", errors="replace", stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if result.returncode != 0:
        shutil.rmtree(temporary, ignore_errors=True)
        raise RuntimeError(f"RAR 解压失败：{archive.name}\n{result.stdout[-1000:]}")
    _validate_extracted_tree(temporary)
    target.parent.mkdir(parents=True, exist_ok=True)
    os.replace(temporary, target)
    marker.write_text(json.dumps({"archive": str(archive), "sha256": key}, ensure_ascii=False), encoding="utf-8")
    return target


def material_root(extracted: Path) -> Path:
    code_docs, _ = _code_files(extracted)
    if not code_docs:
        return extracted
    common = Path(os.path.commonpath([str(path.parent) for path in code_docs]))
    if len({path.parent for path in code_docs}) == 1 and common.parent != extracted.parent:
        return common.parent
    return common


def discover_archives(source: Path) -> list[Path]:
    if source.is_file():
        if source.suffix.casefold() != ".rar":
            raise ValueError("代码注释清理阶段当前只接受 RAR")
        return [source]
    return sorted((path for path in source.rglob("*.rar") if "_已处理_" not in str(path)), key=lambda p: str(p).casefold())


def _write_report(path: Path, report: dict) -> None:
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    rows = []
    for archive in report.get("archives", []):
        if archive.get("error"):
            rows.append(f"<article><h3>{html.escape(archive.get('archive', ''))}</h3><p>压缩包处理失败：{html.escape(archive.get('error', ''))}</p></article>")
        for document in archive.get("documents", []):
            document_name = document.get("docx") or document.get("pdf") or "代码文件"
            findings = "".join(
                f"<li><b>第 {item['paragraph']} 段</b> [{html.escape(item['kind'])}] "
                f"{html.escape(item.get('original', ''))} → {html.escape(item.get('result', ''))} "
                f"{html.escape(item.get('reason', ''))}</li>" for item in document.get("findings", [])
            )
            rows.append(f"<article><h3>{html.escape(document_name)}</h3><p>状态：{html.escape(document.get('status', ''))}；删除 {document.get('deletions', 0)} 处</p><ul>{findings}</ul></article>")
    state = "执行结果" if report.get("applied") else "执行前预览"
    document = f"""<!doctype html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\"><title>代码注释清理{state}</title>
<style>body{{font:15px/1.6 'Microsoft YaHei',sans-serif;max-width:1200px;margin:30px auto;color:#17202a}}article{{border:1px solid #d9e2ec;padding:14px;margin:12px 0}}h1{{color:#165dff}}li{{white-space:pre-wrap;margin:6px 0}}.note{{background:#fff7e6;padding:12px}}</style></head><body><h1>代码中文注释清理—{state}</h1><p class=\"note\">只处理名称以“代码.docx”结尾的文件；TXT 和说明文件未修改。不确定项保留并列在下方。</p>{''.join(rows)}</body></html>"""
    path.with_suffix(".html").write_text(document, encoding="utf-8")


@contextlib.contextmanager
def _word_lock(timeout_seconds: int = 180):
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
    try:
        result = kernel32.WaitForSingleObject(handle, timeout_seconds * 1000)
        if result not in (0, 0x80):
            raise TimeoutError("等待 Word 处理队列超时")
        acquired = True
        yield
    finally:
        if acquired:
            kernel32.ReleaseMutex(handle)
        kernel32.CloseHandle(handle)


def _kill_owned_processes(process: subprocess.Popen, pid_path: Path) -> None:
    if os.name != "nt":
        process.kill()
        return
    subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
    try:
        word_pid = int(pid_path.read_text(encoding="utf-8-sig").strip())
    except (OSError, ValueError):
        return
    subprocess.run(["taskkill", "/PID", str(word_pid), "/T", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)


def _export_pdf(docx: Path, app_root: Path, work: Path) -> tuple[bool, str]:
    pdf = docx.with_suffix(".pdf")
    temp_pdf = work / "pdf" / (hashlib.sha1(str(docx).encode("utf-8")).hexdigest() + ".pdf")
    result_path = temp_pdf.with_suffix(".json")
    pid_path = temp_pdf.with_suffix(".pid")
    temp_pdf.parent.mkdir(parents=True, exist_ok=True)
    script = app_root / "scripts/export_modified_pdf_item.ps1"
    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-DocxPath", str(docx), "-TemporaryPdfPath", str(temp_pdf), "-ResultPath", str(result_path), "-PidPath", str(pid_path)]
    with _word_lock():
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        try:
            process.communicate(timeout=30)
            completed_returncode = process.returncode
        except subprocess.TimeoutExpired:
            _kill_owned_processes(process, pid_path)
            process.wait(timeout=10)
            completed_returncode = None
    method = "Microsoft Word"
    if completed_returncode != 0 or not temp_pdf.is_file():
        soffice = next((Path(value) for value in (shutil.which("soffice.exe"), r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe") if value and Path(value).is_file()), None)
        if soffice is None:
            return False, "Word 导出失败且未安装 LibreOffice；已保留原 PDF"
        outdir = temp_pdf.parent / (temp_pdf.stem + "-lo")
        outdir.mkdir(parents=True, exist_ok=True)
        lo = subprocess.run([str(soffice), "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(docx)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=90, check=False)
        generated = outdir / (docx.stem + ".pdf")
        if lo.returncode != 0 or not generated.is_file():
            return False, "Word/LibreOffice 导出都失败；已保留原 PDF"
        shutil.copy2(generated, temp_pdf)
        method = "LibreOffice"
    try:
        from pypdf import PdfReader
        if temp_pdf.stat().st_size <= 0 or len(PdfReader(str(temp_pdf)).pages) <= 0:
            raise ValueError("无有效页")
        target_temp = pdf.with_name(pdf.name + ".easysoftware.tmp")
        shutil.copy2(temp_pdf, target_temp)
        os.replace(target_temp, pdf)
        return True, method
    except Exception as exc:
        return False, f"新 PDF 验证失败：{exc}；已保留原 PDF"


def run(
    source: Path, work: Path, app_root: Path, apply: bool = False,
    skip_preview: bool = False, progress=None,
) -> dict:
    archives = discover_archives(source.resolve())
    if not archives:
        raise ValueError("未找到 RAR 压缩包")
    report = {"version": RULE_VERSION, "created_at": datetime.now().astimezone().isoformat(), "source": str(source.resolve()), "applied": apply, "archives": []}
    prepared = []
    failures = 0
    if progress:
        progress.update("准备项目清单", 0, active=True, failures=0)
    for archive in archives:
        archive_row = {"archive": str(archive), "status": "READY", "documents": [], "output": None}
        try:
            extracted = extract_cached(archive, work, app_root)
            root = material_root(extracted)
            archive_row["documents"] = _document_rows(root, inspect=False)
            prepared.append((archive, archive_row, root))
        except Exception as exc:
            archive_row["status"] = "FAILED"
            archive_row["error"] = str(exc)
            archive_row["documents"] = [{"docx": "", "status": "FAILED", "deletions": 0, "error": str(exc), "findings": []}]
            prepared.append((archive, archive_row, None))
        report["archives"].append(archive_row)

    total = sum(len(row["documents"]) for _, row, _ in prepared)
    if progress:
        progress.set_total(total)
        progress.update("项目清单已就绪", 0, active=False, failures=0)

    completed = 0
    for archive, archive_row, root in prepared:
        archive_start = completed
        if root is None:
            failures += 1
            completed += 1
            if progress:
                progress.update("读取RAR失败", completed, "无法读取项目", status="FAILED", active=False, failures=failures, archive=archive.name, error=archive_row.get("error"))
            continue
        try:
            if apply:
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output = archive.parent / f"{archive.stem}_已处理_{stamp}"
                suffix = 2
                while output.exists():
                    output = archive.parent / f"{archive.stem}_已处理_{stamp}_{suffix}"
                    suffix += 1
                shutil.copytree(root, output)
                applied_documents = []
                for discovered in _document_rows(output, inspect=False):
                    project_name = Path(discovered.get("docx") or discovered.get("pdf") or "未识别项目").stem
                    if progress:
                        progress.update("清理代码注释", completed, project_name, active=True, failures=failures, archive=archive.name)
                    if discovered["status"] != "READY":
                        applied_documents.append(discovered)
                        failures += 1
                        completed += 1
                        if progress:
                            progress.update("项目跳过", completed, project_name, status="FAILED", active=False, failures=failures, archive=archive.name, error=discovered.get("error"))
                        continue
                    docx = Path(discovered["docx"])
                    try:
                        row = clean_docx(docx)
                        if row["deletions"]:
                            if progress:
                                progress.update("导出代码PDF", completed, project_name, active=True, failures=failures, archive=archive.name)
                            success, detail = _export_pdf(docx, app_root, work)
                            row["pdf_status"] = "REPLACED" if success else "PRESERVED"
                            row["pdf_detail"] = detail
                            if not success:
                                row["status"] = "CLEANED_PDF_FAILED"
                                failures += 1
                        else:
                            row["pdf_status"] = "UNCHANGED"
                            row["pdf_detail"] = "未删除注释，无需重新导出"
                    except Exception as exc:
                        row = {"docx": str(docx), "status": "FAILED", "deletions": 0, "error": str(exc), "findings": []}
                        failures += 1
                    applied_documents.append(row)
                    completed += 1
                    if progress:
                        item_failed = row.get("status") in {"FAILED", "CLEANED_PDF_FAILED"}
                        progress.update("项目处理完成" if not item_failed else "项目处理失败", completed, project_name, status="FAILED" if item_failed else "RUNNING", active=False, failures=failures, archive=archive.name, error=row.get("error") or (row.get("pdf_detail") if item_failed else None))
                archive_row["documents"] = applied_documents
                archive_row["output"] = str(output)
            else:
                inspected_documents = []
                for discovered in archive_row["documents"]:
                    project_name = Path(discovered.get("docx") or discovered.get("pdf") or "未识别项目").stem
                    if progress:
                        progress.update("扫描代码注释", completed, project_name, active=True, failures=failures, archive=archive.name)
                    row = inspect_docx(Path(discovered["docx"])) if discovered["status"] == "READY" else discovered
                    if row.get("status") != "READY":
                        failures += 1
                    inspected_documents.append(row)
                    completed += 1
                    if progress:
                        item_failed = row.get("status") != "READY"
                        progress.update("预览扫描完成" if not item_failed else "预览扫描失败", completed, project_name, status="FAILED" if item_failed else "RUNNING", active=False, failures=failures, archive=archive.name, error=row.get("error"))
                archive_row["documents"] = inspected_documents
            archive_row["status"] = "PARTIAL" if any(item.get("status") not in {"READY", "CLEANED"} for item in archive_row["documents"]) else "OK"
        except Exception as exc:
            archive_row["status"] = "FAILED"
            archive_row["error"] = str(exc)
            left = max(0, len(archive_row["documents"]) - (completed - archive_start))
            failures += left
            completed += left
            if progress:
                progress.update("压缩包处理失败", completed, status="FAILED", active=False, failures=failures, archive=archive.name, error=str(exc))
    report_path = work / ("code_comment_cleanup_result.json" if apply else "code_comment_cleanup_preview.json")
    report_path.parent.mkdir(parents=True, exist_ok=True)
    _write_report(report_path, report)
    deletions = sum(doc.get("deletions", 0) for arc in report["archives"] for doc in arc["documents"])
    report_failures = sum(doc.get("status") in {"FAILED", "CLEANED_PDF_FAILED", "SKIPPED_MISSING_PDF", "SKIPPED_MISSING_DOCX", "MISSING_CODE_DOCX"} for arc in report["archives"] for doc in arc["documents"])
    failures = max(failures, report_failures)
    final_status = ("CODE_COMMENT_CLEANUP_PARTIAL" if failures else "CODE_COMMENT_CLEANUP_OK") if apply else "NEEDS_CODE_COMMENT_REVIEW"
    if progress:
        progress.update("代码注释清理完成" if apply else "代码注释预览完成", total, status="DONE" if not failures else "FAILED", active=False, failures=failures)
    return {
        "status": final_status,
        "report": str(report_path.with_suffix(".html")), "report_json": str(report_path),
        "deletions": deletions, "failures": failures,
        "outputs": [arc["output"] for arc in report["archives"] if arc.get("output")],
        "work_dir": str(work), "skip_preview": skip_preview,
    }

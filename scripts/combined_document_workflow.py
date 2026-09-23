from __future__ import annotations

import html
import io
import json
import os
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import archive_types
import code_comment_cleaner
import manual_content_stage
import process_batch


VERSION = 1
PREVIEW_STATUS = "NEEDS_DOCUMENT_PROCESSING_REVIEW"
DOCUMENT_STAGES = ("manual", "code", "txt_placeholder")


def _selected_stages(stages: str | list[str] | tuple[str, ...] | None) -> tuple[str, ...]:
    if stages is None:
        return DOCUMENT_STAGES
    requested = stages.split(",") if isinstance(stages, str) else stages
    selected = {str(stage).strip() for stage in requested}
    selected.discard("")
    unknown = selected.difference(DOCUMENT_STAGES)
    if unknown:
        raise ValueError(f"未知文档处理阶段：{', '.join(sorted(unknown))}")
    if not selected.intersection({"manual", "code"}):
        raise ValueError("至少选择一个可执行阶段：manual 或 code")
    return tuple(stage for stage in DOCUMENT_STAGES if stage in selected)


def _publish_progress(
    progress: Any,
    phase: str,
    current: int,
    name: str = "",
    *,
    status: str = "RUNNING",
    failures: int = 0,
    error: str | None = None,
    active: bool | None = None,
) -> None:
    if progress is None:
        return
    details = {"status": status, "failures": failures}
    if error:
        details["error"] = error
    if active is not None:
        details["active"] = active
    target = progress.update if hasattr(progress, "update") else progress
    try:
        target(phase, current, name, **details)
    except TypeError:
        target(phase, current, name)


class _StageProgress:
    def __init__(self, progress: Any, offset: int, failures: int = 0):
        self.progress = progress
        self.offset = offset
        self.failures = failures

    def update(self, phase: str, current: int = 0, name: str = "", **details: Any) -> None:
        _publish_progress(
            self.progress,
            phase,
            self.offset + current,
            name,
            status=str(details.get("status", "RUNNING")),
            failures=max(self.failures, int(details.get("failures", 0))),
            error=details.get("error"),
            active=details.get("active"),
        )


def _material_root(source: Path, work: Path, app_root: Path) -> tuple[Path, str]:
    source = source.resolve()
    if source.is_file():
        if not archive_types.is_archive_file(source):
            raise ValueError("材料输入需要是压缩包（RAR/ZIP/7z 等）或已解压的材料文件夹")
        extracted = code_comment_cleaner.extract_cached(source, work, app_root)
        root = code_comment_cleaner.material_root(extracted)
        nested_state_path = work / "combined_nested_archives.json"
        try:
            prior = json.loads(nested_state_path.read_text(encoding="utf-8")) if nested_state_path.is_file() else []
        except (OSError, json.JSONDecodeError):
            prior = []
        hashes = process_batch.HashCache(work / "combined_hashes.json")
        nested = process_batch.expand_nested_archives(
            root,
            process_batch.find_7z(),
            max(1, min(4, os.cpu_count() or 1)),
            hashes,
            prior_results=prior,
        )
        hashes.save()
        nested_state_path.write_text(json.dumps(nested, ensure_ascii=False, indent=2), encoding="utf-8")
        return code_comment_cleaner.material_root(root), source.stem
    if source.is_dir():
        if "_已处理_" in source.name:
            raise ValueError("不能把已处理目录再次作为输入")
        return source, source.name
    raise FileNotFoundError(f"输入不存在：{source}")


def _code_inventory(root: Path, *, inspect: bool) -> list[dict[str, Any]]:
    docx_files, pdf_files = code_comment_cleaner._code_files(root)
    pdf_by_key = {code_comment_cleaner._pair_key(path): path for path in pdf_files}
    docx_keys = {code_comment_cleaner._pair_key(path) for path in docx_files}
    rows: list[dict[str, Any]] = []
    for docx in docx_files:
        pdf = pdf_by_key.get(code_comment_cleaner._pair_key(docx))
        if pdf is None:
            rows.append({
                "docx": str(docx), "pdf": "", "status": "SKIPPED_MISSING_PDF",
                "deletions": 0, "error": "缺少同名代码 PDF，未处理", "findings": [],
            })
            continue
        row = code_comment_cleaner.inspect_docx(docx) if inspect else {
            "docx": str(docx), "status": "READY", "deletions": 0, "findings": [],
        }
        row["pdf"] = str(pdf)
        rows.append(row)
    for pdf in pdf_files:
        if code_comment_cleaner._pair_key(pdf) not in docx_keys:
            rows.append({
                "docx": "", "pdf": str(pdf), "status": "SKIPPED_MISSING_DOCX",
                "deletions": 0, "error": "缺少同名代码 DOCX，未处理", "findings": [],
            })
    if not docx_files and not pdf_files:
        rows.append({
            "docx": "", "pdf": "", "status": "MISSING_CODE_DOCX",
            "deletions": 0, "error": "未找到代码 DOCX，未处理", "findings": [],
        })
    return rows


def _run_code_stage(
    root: Path,
    app_root: Path,
    work: Path,
    *,
    apply: bool,
    progress: Any,
    offset: int,
    failures: int,
) -> tuple[dict[str, Any], int]:
    rows = _code_inventory(root, inspect=False)
    results: list[dict[str, Any]] = []
    stage_failures = 0
    for index, discovered in enumerate(rows, 1):
        item_name = Path(discovered.get("docx") or discovered.get("pdf") or "代码材料").stem
        current = offset + index
        try:
            if discovered["status"] != "READY":
                row = discovered
            elif not apply:
                row = code_comment_cleaner.inspect_docx(Path(discovered["docx"]))
                row["pdf"] = discovered["pdf"]
            else:
                row = code_comment_cleaner.clean_docx(Path(discovered["docx"]))
                row["pdf"] = discovered["pdf"]
                if row["deletions"]:
                    _publish_progress(progress, "导出代码PDF", current - 1, item_name, failures=failures + stage_failures)
                    with process_batch.word_automation_lock(progress, "代码处理"):
                        exported, detail = code_comment_cleaner._export_pdf(Path(discovered["docx"]), app_root, work)
                    row["pdf_status"] = "REPLACED" if exported else "PRESERVED"
                    row["pdf_detail"] = detail
                    if not exported:
                        row["status"] = "CLEANED_PDF_FAILED"
                else:
                    row["pdf_status"] = "UNCHANGED"
                    row["pdf_detail"] = "未删除注释，无需重新导出"
        except Exception as exc:
            row = {
                "docx": discovered.get("docx", ""), "pdf": discovered.get("pdf", ""),
                "status": "FAILED", "deletions": 0, "error": str(exc), "findings": [],
            }
        item_failed = row.get("status") not in {"READY", "CLEANED"}
        if item_failed:
            stage_failures += 1
        results.append(row)
        _publish_progress(
            progress,
            "代码处理完成" if apply else "代码预检完成",
            current,
            item_name,
            status="FAILED" if item_failed else "RUNNING",
            failures=failures + stage_failures,
            error=row.get("error") or (row.get("pdf_detail") if item_failed else None),
            active=False,
        )
    return {
        "stage": "code_comments",
        "status": "PARTIAL" if stage_failures else ("APPLIED" if apply else "READY_FOR_REVIEW"),
        "documents": results,
        "deletions": sum(int(row.get("deletions", 0)) for row in results),
        "failures": stage_failures,
    }, stage_failures


def _run_txt_placeholder(
    root: Path,
    *,
    progress: Any,
    offset: int,
    failures: int,
) -> dict[str, Any]:
    files = sorted((path for path in root.rglob("*.txt") if path.is_file()), key=lambda path: str(path).casefold())
    items = [{"path": str(path), "status": "UNCHANGED_PLACEHOLDER"} for path in files]
    work_items = files or [None]
    for index, path in enumerate(work_items, 1):
        _publish_progress(
            progress,
            "TXT阶段占位",
            offset + index,
            path.name if path else "未找到TXT",
            failures=failures,
            active=False,
        )
    return {
        "stage": "txt_placeholder",
        "status": "PLACEHOLDER_NOT_IMPLEMENTED",
        "message": "TXT 阶段当前仅登记文件，不读取也不修改内容。",
        "files": items,
        "file_count": len(items),
        "failures": 0,
    }


def _manual_failure_count(report: dict[str, Any], apply: bool) -> int:
    if not apply:
        return 0
    failed_statuses = {"BATCH_FUSED", "BLOCKED_IRRELEVANT", "DATE_REWRITER_UNAVAILABLE"}
    return sum(project.get("status") in failed_statuses for project in report.get("projects", []))


def _build_manual_cover_plan(root: Path, path: Path) -> dict[str, Any]:
    packages: list[dict[str, Any]] = []
    audit_errors: list[dict[str, str]] = []
    for docx in sorted(root.rglob("*说明.docx"), key=lambda item: str(item).casefold()):
        title = docx.stem.removesuffix("说明")
        try:
            audit = process_batch.inspect_docx(docx, title, False)
            actions: list[str] = []
            if not audit.get("cover_found"):
                actions.append("ADD_COVER")
            if audit.get("duplicate_cover_blocks") or audit.get("residual_manual_labels"):
                actions.append("REMOVE_DUPLICATE_COVER_TEXT")
            if audit.get("cover_found") and (
                not audit.get("cover_style_explicit_ok") or not audit.get("cover_layout_explicit_ok")
            ):
                actions.append("NORMALIZE_COVER_STYLE")
            if audit.get("cover_title_has_version"):
                actions.extend(["REMOVE_COVER_VERSION", "NORMALIZE_COVER_STYLE"])
            if not audit.get("header_compliant"):
                actions.append("NORMALIZE_HEADER")
            if audit.get("body_text_style_mismatch") or audit.get("body_text_style_needs_word_check") or audit.get("inline_body_font_mismatches"):
                actions.append("NORMALIZE_BODY_TEXT_STYLE")
            if audit.get("first_post_cover_heading_needs_word_check"):
                actions.append("NORMALIZE_FIRST_POST_COVER_HEADING")
            # Every manual enters the Word pass: the pass always verifies the
            # cover/page boundary (including images/tables) before exporting PDF.
            actions.append("EXPORT_PDF")
            pdf = docx.with_suffix(".pdf")
            packages.append({
                "name": title,
                "folder": str(docx.parent.relative_to(root)),
                "docx": docx.name,
                "pdf": pdf.name,
                "actions": list(dict.fromkeys(actions)),
            })
        except Exception as exc:
            audit_errors.append({"docx": str(docx), "error": str(exc)})
    plan = {
        "version": VERSION,
        "base": str(root),
        "created_at": datetime.now().astimezone().isoformat(),
        "packages": packages,
        "audit_errors": audit_errors,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    return plan


def _cover_result_failures(report: dict[str, Any]) -> int:
    return sum(
        row.get("status") not in {"OK"} or bool(row.get("failures"))
        for row in report.get("packages", [])
    )


def _year_image_comparisons(path: Path, report: dict[str, Any]) -> str:
    """Build offline image evidence from the original and processed DOCX packages."""
    if not report.get("applied"):
        return '<p>预览阶段尚未修改图片；执行后显示原稿与输出的图片对比。</p>'
    source_root = Path(report["material_root"]).resolve()
    output_root = Path(report["output"]).resolve()
    assets = path.with_name(path.stem + "_assets")
    cards: list[str] = []
    try:
        from PIL import Image, ImageChops, ImageOps
    except ImportError:
        return '<p>缺少 Pillow，无法核对原稿与输出图片的像素差异。</p>'
    supported = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
    for project in report["stages"]["manual"].get("projects", []):
        result = project.get("date_rewrite") or {}
        if not result.get("changed"):
            continue
        recorded_media = {
            item.get("media_path") for item in result.get("changes", [])
            if item.get("changed_image") and item.get("media_path")
        }
        relative = Path(str(project.get("relative_path", "")))
        original = (source_root / relative).resolve()
        processed = (output_root / relative).resolve()
        if not original.is_relative_to(source_root) or not processed.is_relative_to(output_root):
            continue
        if not original.is_file() or not processed.is_file():
            continue
        try:
            with zipfile.ZipFile(original) as before_doc, zipfile.ZipFile(processed) as after_doc:
                before_names = {
                    name for name in before_doc.namelist()
                    if name.lower().startswith("word/media/") and Path(name).suffix.lower() in supported
                }
                after_names = set(after_doc.namelist())
                changed = []
                for media_name in sorted(before_names & after_names & recorded_media):
                    before_data = before_doc.read(media_name)
                    after_data = after_doc.read(media_name)
                    if before_data == after_data:
                        continue
                    try:
                        before_image = ImageOps.exif_transpose(Image.open(io.BytesIO(before_data))).convert("RGB")
                        after_image = ImageOps.exif_transpose(Image.open(io.BytesIO(after_data))).convert("RGB")
                        # Word's later PDF/cover pass can resize other pictures.
                        # The year rewriter preserves dimensions, so those are not
                        # evidence of its date edits.
                        if before_image.size != after_image.size:
                            continue
                        if ImageChops.difference(before_image, after_image).getbbox() is None:
                            continue
                    except (OSError, ValueError):
                        continue
                    changed.append((media_name, before_data, after_data, before_image, after_image))
                for index, (media_name, before_data, after_data, before_image, after_image) in enumerate(changed, 1):
                    assets.mkdir(parents=True, exist_ok=True)
                    prefix = f"{len(cards) + 1:03d}"
                    files = []
                    for label, picture in (("before", before_image), ("after", after_image)):
                        suffix = ".png"
                        filename = f"{prefix}-{label}{suffix}"
                        target = assets / filename
                        picture.save(target, format="PNG")
                        files.append(f"{assets.name}/{filename}")
                    zoom = ""
                    if before_image is not None and after_image is not None and before_image.size == after_image.size:
                        # JPEG recompression causes low-level noise across the image.
                        # A high threshold isolates the visually changed year glyphs.
                        difference = ImageChops.difference(before_image, after_image).convert("L")
                        strong = difference.point(lambda value: 255 if value >= 45 else 0)
                        bounds = strong.getbbox()
                        if bounds:
                            width, height = before_image.size
                            left, top, right, bottom = bounds
                            if (right - left) * (bottom - top) < width * height * 0.5:
                                margin = max(20, min(width, height) // 30)
                                crop_box = (max(0, left - margin), max(0, top - margin),
                                            min(width, right + margin), min(height, bottom + margin))
                                crop_files = []
                                for label, picture in (("before", before_image), ("after", after_image)):
                                    filename = f"{prefix}-{label}-detail.png"
                                    picture.crop(crop_box).save(assets / filename, format="PNG")
                                    crop_files.append(f"{assets.name}/{filename}")
                                zoom = ("<div class='pair'><figure><img src='{}' alt='修改前局部'><figcaption>修改前局部</figcaption></figure>"
                                        "<figure><img src='{}' alt='修改后局部'><figcaption>修改后局部</figcaption></figure></div>").format(
                                            *(html.escape(value, quote=True) for value in crop_files))
                    title = html.escape(str(relative))
                    name = html.escape(media_name)
                    recorded_fields = [
                        f"{item.get('source_year')} → {item.get('target_year')}（{item.get('ocr_text')}）"
                        for item in result.get("changes", [])
                        if item.get("media_path") == media_name and item.get("changed_image")
                    ]
                    field_detail = (
                        "<p>识别到的旧日期年份：" + html.escape("；".join(recorded_fields)) + "</p>"
                        if recorded_fields else ""
                    )
                    cards.append(
                        f"<article class='comparison'><h3>{title} · {name}</h3>"
                        + field_detail +
                        "<p class='review'>待核验：这张图片的像素发生变化，但仅凭像素差异无法确认修改的是年份。"
                        "请对照原图检查日期语义和处理后字符，尤其留意时间、编号被误识别为年份的情况。</p>"
                        + zoom +
                        "<div class='pair'>" + "".join(
                            f"<figure><a href='{html.escape(src, quote=True)}' target='_blank'>"
                            f"<img src='{html.escape(src, quote=True)}' alt='{caption}'></a>"
                            f"<figcaption>{caption}（点击查看原尺寸）</figcaption></figure>"
                            for src, caption in zip(files, ("原稿", "处理后"))
                        ) + "</div></article>"
                    )
        except (OSError, zipfile.BadZipFile):
            continue
    if not cards:
        return '<p>没有可展示的图片年份修改对比。</p>'
    return (f"<p>以下 {len(cards)} 张图片从原稿与处理后 DOCX 提取，确认存在像素变化。"
            "自动检测结果尚需逐张核验，不能仅据此认定年份改写正确。</p>" + "".join(cards))


def _write_report(path: Path, report: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)

    stage_rows = []
    for key in ("manual", "code", "txt"):
        stage = report["stages"][key]
        stage_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape({"manual": "说明书", "code": "代码", "txt": "TXT"}[key]),
                html.escape(str(stage.get("status", ""))),
                int(stage.get("failures", 0)),
            )
        )
    code_rows = []
    for item in report["stages"]["code"].get("documents", []):
        name = item.get("docx") or item.get("pdf") or "代码材料"
        code_rows.append(
            f"<li><b>{html.escape(str(name))}</b>：{html.escape(str(item.get('status', '')))}；"
            f"预计/实际删除 {int(item.get('deletions', 0))} 处"
            f"{('；' + html.escape(str(item.get('error')))) if item.get('error') else ''}</li>"
        )
    manual_rows = []
    visual_rows = []
    for project in report["stages"]["manual"].get("projects", []):
        samples = project.get("samples", [])
        counts = {
            state: sum(item.get("status") == state for item in samples)
            for state in ("relevant", "irrelevant", "unknown")
        }
        year_rows = project.get("year_changes", [])
        updated_years = sum(bool(item.get("changed")) for item in year_rows)
        review_years = sum(item.get("decision") == "REVIEW" for item in year_rows)
        date_result = project.get("date_rewrite", {})
        body_relevance = project.get("body_relevance", {})
        body_status = "需人工复核" if body_relevance.get("status") == "review_topic_mismatch" else "证据不足"
        body_detail = str(body_relevance.get("reason", "尚未审核"))
        visual_samples = [item for item in samples if item.get("visual_review", {}).get("status") != "not_requested" and item.get("visual_review")]
        visual_counts = {
            state: sum(item["visual_review"].get("status") == state for item in visual_samples)
            for state in ("relevant", "irrelevant", "unknown")
        }
        visual_summary = (
            f"抽看 {int(project.get('visual_call_count', 0))}/{int(project.get('visual_candidate_count', len(visual_samples)))}；"
            f"相关 {visual_counts['relevant']} / 不相关 {visual_counts['irrelevant']} / 无法判断 {visual_counts['unknown']}；"
            f"接口失败 {int(project.get('visual_failure_count', 0))}"
            if visual_samples else "未请求"
        )
        manual_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}：{}</td><td>相关 {} / 不相关 {} / 无法判断 {}</td><td>{}</td>"
            "<td>更新 {} / 人工复核 {}</td><td>{}</td></tr>".format(
                html.escape(str(project.get("relative_path", ""))),
                html.escape(str(project.get("status", ""))),
                body_status, html.escape(body_detail),
                counts["relevant"], counts["irrelevant"], counts["unknown"],
                visual_summary,
                updated_years, review_years,
                "检测到变化（待核验）" if date_result.get("changed") else html.escape(str(date_result.get("status", "预览阶段未修改"))),
            )
        )
        if visual_samples:
            items = "".join(
                "<li>{}：{}；{}</li>".format(
                    html.escape(str(item.get("image", ""))),
                    html.escape(str(item["visual_review"].get("status", "unknown"))),
                    html.escape(str(item["visual_review"].get("reason") or item["visual_review"].get("error") or "无可用证据")),
                )
                for item in visual_samples
            )
            visual_rows.append(
                f"<details><summary>{html.escape(str(project.get('relative_path', '')))}：{visual_summary}</summary><ul>{items}</ul></details>"
            )
    cover_rows = []
    year_comparisons = _year_image_comparisons(path, report)
    cover_result = report["stages"]["manual"].get("cover_result", {})
    for item in cover_result.get("packages", []):
        requested = "、".join(map(str, item.get("actions_requested", item.get("actions", [])))) or "无"
        applied = "、".join(map(str, item.get("actions_applied", []))) or ("执行阶段生成" if not report["applied"] else "无")
        failures_text = "；".join(map(str, item.get("failures", []))) or "无"
        cover_rows.append(
            f"<tr><td>{html.escape(str(item.get('name', '')))}</td><td>{html.escape(requested)}</td>"
            f"<td>{html.escape(applied)}</td><td>{html.escape(failures_text)}</td></tr>"
        )
    document = f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>材料统一处理报告</title>
<style>body{{font:15px/1.6 'Microsoft YaHei',sans-serif;max-width:1100px;margin:30px auto;color:#17202a}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #d9e2ec;padding:9px;text-align:left}}th{{background:#eef4ff}}section{{margin:24px 0}}.note,.review{{background:#fff7e6;padding:12px}}.comparison{{border:1px solid #d9e2ec;padding:14px;margin:18px 0}}.comparison h3{{overflow-wrap:anywhere}}.pair{{display:flex;gap:16px;flex-wrap:wrap}}figure{{margin:0;flex:1 1 300px;min-width:0}}figure img{{max-width:100%;max-height:460px;object-fit:contain;border:1px solid #ddd}}figcaption{{font-weight:bold}}</style></head><body>
<h1>材料统一处理{'执行结果' if report['applied'] else '执行前预览'}</h1>
<p class="note">所选流程按说明书 → 代码 → TXT 顺序处理，未选流程跳过。TXT 当前仅占位登记，不读取、不修改。图片日期仅将早于 2025 的年份更新为 2025；2025 及之后的年份保留。</p>
<p>状态：{html.escape(report['status'])}；失败/待处理项：{report['failures']}；输出目录：{html.escape(str(report.get('output') or '预览未创建'))}</p>
<table><thead><tr><th>阶段</th><th>状态</th><th>失败/待处理</th></tr></thead><tbody>{''.join(stage_rows)}</tbody></table>
<section><h2>说明书相关性、正文年份与图片年份</h2><p>视觉复核仅提供人工审核证据，不参与现有 OCR 拦截及批量熔断。</p><table><thead><tr><th>说明书</th><th>状态</th><th>标题与正文</th><th>抽样图片 OCR</th><th>视觉复核</th><th>正文年份</th><th>图片年份</th></tr></thead><tbody>{''.join(manual_rows)}</tbody></table>{''.join(visual_rows)}</section>
<section><h2>图片变化前后对比（待核验）</h2>{year_comparisons}</section>
<section><h2>封面与 PDF</h2><table><thead><tr><th>项目</th><th>计划动作</th><th>实际动作</th><th>仍需处理</th></tr></thead><tbody>{''.join(cover_rows)}</tbody></table></section>
<section><h2>代码文件明细</h2><ul>{''.join(code_rows)}</ul></section>
</body></html>"""
    html_path = path.with_suffix(".html")
    html_temporary = html_path.with_suffix(html_path.suffix + ".tmp")
    html_temporary.write_text(document, encoding="utf-8")
    os.replace(html_temporary, html_path)


def _unique_output(source: Path, label: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = source.parent / f"{label}_已处理_{stamp}"
    suffix = 2
    while output.exists():
        output = source.parent / f"{label}_已处理_{stamp}_{suffix}"
        suffix += 1
    return output


def run(
    source: str | Path,
    work: str | Path,
    app_root: str | Path,
    apply: bool = False,
    progress: Any = None,
    *,
    manual_options: dict[str, Any] | None = None,
    document_stages: str | list[str] | tuple[str, ...] | None = None,
) -> dict[str, Any]:
    selected_stages = _selected_stages(document_stages)
    source_path = Path(source).resolve()
    work_path = Path(work).resolve()
    app_root_path = Path(app_root).resolve()
    work_path.mkdir(parents=True, exist_ok=True)
    material_root, label = _material_root(source_path, work_path, app_root_path)

    manual_count = len(list(material_root.rglob("*说明.docx"))) if "manual" in selected_stages else 0
    code_count = len(_code_inventory(material_root, inspect=False)) if "code" in selected_stages else 0
    txt_count = sum(path.is_file() for path in material_root.rglob("*.txt")) if "txt_placeholder" in selected_stages else 0
    total = manual_count + code_count + (max(1, txt_count) if "txt_placeholder" in selected_stages else 0)
    if progress is not None and hasattr(progress, "set_total"):
        progress.set_total(total)
    _publish_progress(progress, "材料清单已就绪", 0, failures=0, active=False)

    output: Path | None = None
    processing_root = material_root
    if apply:
        output = _unique_output(source_path, label)
        shutil.copytree(material_root, output)
        processing_root = output

    manual_stage: dict[str, Any] = {"stage": "manual", "status": "SKIPPED", "failures": 0}
    manual_failures = 0
    if "manual" in selected_stages:
        options = dict(manual_options or {})
        for reserved in ("input_root", "output_root", "report_path", "progress", "apply"):
            options.pop(reserved, None)
        manual_report_path = work_path / ("combined_manual_result.json" if apply else "combined_manual_preview.json")
        manual_report = manual_content_stage.run_manual_content_stage(
            processing_root, processing_root, manual_report_path,
            progress=_StageProgress(progress, 0), apply=apply, **options,
        )
        cover_plan_path = work_path / ("combined_cover_plan_result.json" if apply else "combined_cover_plan_preview.json")
        cover_plan = _build_manual_cover_plan(processing_root, cover_plan_path)
        cover_report: dict[str, Any] = {
            "status": "READY_FOR_REVIEW", "packages": cover_plan.get("packages", []),
            "audit_errors": cover_plan.get("audit_errors", []),
        }
        cover_failures = len(cover_plan.get("audit_errors", []))
        if apply and cover_plan.get("packages"):
            _publish_progress(progress, "说明书Word封面与PDF处理", 0, failures=cover_failures, active=True)
            cover_report = process_batch.invoke_word_pipeline(
                processing_root, cover_plan_path, work_path / "combined_cover_result.json",
                SCRIPT_DIR / "word_manual_pipeline_stable.ps1", progress=progress, module="说明书处理",
            )
            cover_failures += _cover_result_failures(cover_report)
            _publish_progress(progress, "说明书Word封面与PDF处理完成", manual_count, failures=cover_failures, active=False)
        manual_failures = _manual_failure_count(manual_report, apply) + cover_failures
        manual_stage = {
            **manual_report,
            "status": "PARTIAL" if manual_failures else ("APPLIED" if apply else "READY_FOR_REVIEW"),
            "failures": manual_failures, "cover_plan": cover_plan, "cover_result": cover_report,
        }

    code_stage: dict[str, Any] = {"stage": "code_comments", "status": "SKIPPED", "failures": 0}
    code_failures = 0
    if "code" in selected_stages:
        code_stage, code_failures = _run_code_stage(
            processing_root, app_root_path, work_path, apply=apply, progress=progress,
            offset=manual_count, failures=manual_failures,
        )
    txt_stage: dict[str, Any] = {"stage": "txt_placeholder", "status": "SKIPPED", "failures": 0}
    if "txt_placeholder" in selected_stages:
        txt_stage = _run_txt_placeholder(
            processing_root, progress=progress, offset=manual_count + code_count,
            failures=manual_failures + code_failures,
        )
    failures = manual_failures + code_failures
    status = PREVIEW_STATUS if not apply else ("DOCUMENT_PROCESSING_PARTIAL" if failures else "DOCUMENT_PROCESSING_OK")
    report = {
        "version": VERSION,
        "created_at": datetime.now().astimezone().isoformat(),
        "source": str(source_path),
        "material_root": str(material_root),
        "applied": apply,
        "status": status,
        "output": str(output) if output else None,
        "failures": failures,
        "stage_order": ["manual", "code", "txt"],
        "selected_stages": list(selected_stages),
        "stages": {"manual": manual_stage, "code": code_stage, "txt": txt_stage},
    }
    report_path = work_path / ("combined_document_result.json" if apply else "combined_document_preview.json")
    _write_report(report_path, report)
    _publish_progress(
        progress,
        "材料统一处理完成" if apply else "材料统一预检完成",
        total,
        status="FAILED" if failures else "DONE",
        failures=failures,
        active=False,
    )
    return {
        "status": status,
        "report": str(report_path.with_suffix(".html")),
        "report_json": str(report_path),
        "output": str(output) if output else None,
        "outputs": [str(output)] if output else [],
        "failures": failures,
        "work_dir": str(work_path),
    }


def preview(
    source: str | Path,
    work: str | Path,
    app_root: str | Path,
    progress: Any = None,
    *,
    manual_options: dict[str, Any] | None = None,
    document_stages: str | list[str] | tuple[str, ...] | None = None,
) -> dict[str, Any]:
    return run(source, work, app_root, apply=False, progress=progress,
               manual_options=manual_options, document_stages=document_stages)


def execute(
    source: str | Path,
    work: str | Path,
    app_root: str | Path,
    progress: Any = None,
    *,
    manual_options: dict[str, Any] | None = None,
    document_stages: str | list[str] | tuple[str, ...] | None = None,
) -> dict[str, Any]:
    return run(source, work, app_root, apply=True, progress=progress,
               manual_options=manual_options, document_stages=document_stages)


__all__ = ["PREVIEW_STATUS", "execute", "preview", "run"]

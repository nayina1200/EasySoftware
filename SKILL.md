---
name: process-software-copyright-materials
description: Fully programmatic audit, DOCX-first repair, two-stage visual review, and ZIP/RAR delivery workflow for batches of Chinese software-copyright materials.
---

# Process software-copyright materials

Use the single program entrypoint. Before modifying this program, read and follow `开发维护规则.md`.

## Entry command

Resolve `SKILL_DIR` as the directory containing this `SKILL.md`, then run:

`python "SKILL_DIR/scripts/unified_workflow.py" "BATCH_ROOT"`

The program recursively expands ZIP/RAR/7z inputs, inventories real project folders, audits every project, applies deterministic repairs only in `BATCH_ROOT/修改版`, verifies the outputs, and creates two local HTML review pages. Input archives and `解压版` remain unchanged.

For the isolated code-comment phase, use `--clean-code-comments` to create the mandatory preview and `--apply-comment-cleanup` after approval. This phase only copies paired `*代码.docx`/PDF files into a timestamped output tree, removes confirmed Han-containing comments from the DOCX source, re-exports its PDF, writes HTML/JSON audit reports, and never runs TXT/manual repair or packaging.

## Two human-review gates

The initial command returns `NEEDS_IMAGE_REVIEW`. Open `review_link`; only select image-failed projects for return. Unselected projects continue with:

`python "SKILL_DIR/scripts/unified_workflow.py" "BATCH_ROOT" --reject-images "2,5"`

Use `--reject-images none` when all projects pass the image review. The next result is `NEEDS_UNIFIED_REVIEW`, containing the final result table, TXT changes, repaired manual-cover evidence, and code first/last-page evidence. Select final passing projects with:

`python "SKILL_DIR/scripts/unified_workflow.py" "BATCH_ROOT" --approve "1,3-5"`

Use `--approve all` when every remaining project passes. Packaging contains only final passing project folders. If the input includes both ZIP and RAR and the program returns `NEEDS_ARCHIVE_FORMAT`, add `--format zip` or `--format rar` to the same approval command.

For `MISSING_DEPENDENCIES`, `FAILED`, or any other status, report the exact returned error and `work_dir`. Do not switch to manual document processing.

## Complete programmatic scope

The unified program must preserve every existing function:

- recursively extract nested archives, cache by SHA-256, and preserve source data;
- inventory TXT, code PDF/DOCX, and description PDF/DOCX with exact paths and hashes;
- inspect only code-PDF page 1 and last page, check page count, code-like formatting, title, `V1.0`, page numbers, punctuation/brackets, and generate review boards;
- automatically normalize confirmed code header/page-number defects in the paired code DOCX, preserve code body text, re-export PDF from that DOCX only, and QA first/last pages; never patch PDF directly;
- validate TXT encoding, fields, exact TXT-derived software name, obvious category/content mismatch, and `主要功能`/`软件主要功能` rules: digits or ASCII/full-width parentheses are errors; empty text, missing Chinese, malformed punctuation, or no recognizable functional action are errors;
- compare complete description DOCX text with the TXT-derived software name and flag obvious cross-domain mismatch;
- extract substantial interface screenshots in document order, preserve source resolution, mark programmatic visual outliers, and include login/interface judgment in the unified review;
- normalize description cover, duplicate cover blocks, typography, layout, every-page running header, `V1.0`, and dynamic PAGE field;
- compare all ordinary body paragraphs and image captions with the document's dominant body font, size, bold, and italic format; normalize only mismatched ordinary body paragraphs while excluding headings, numbered items, tables, and cover content;
- remove paired ASCII (`"名称"`) and Chinese (`“名称”`) double quotes wrapping the exact TXT-derived project name throughout Word stories;
- save/export only the modified copy, retain completed repairs for review when a later check fails, and preserve unrelated files;
- verify DOCX validity, cover, title line count, typography, duplicate covers, header/PAGE fields, Word/PDF page counts, title/version/page numbers, and intended body-text preservation;
- release a repair rule only after its verification loop passes: relevant automated tests, a real DOCX-to-PDF temporary-copy run, final QA, and the expected review-page result;
- show processed description-PDF pages 1–2 in the final review, without repeating first-review interface findings;
- after final approval, remove delivery-temporary files, preserve incoming archive format, package only approved projects, integrity-test the archive, and report path, counts, size, and SHA-256.

## Performance and safety

- Trust SHA-256/version caches and reuse unchanged extraction and visual artifacts.
- Use one Word application and one open/save/export cycle per actionable document.
- Render no middle code-PDF pages and do not convert or render the complete explanation PDF for textual review.
- Large embedded screenshots must be processed without Pillow's decompression-bomb pixel limit, while overview splitting prevents oversized review canvases.
- Preserve original archives and `解压版`; only `修改版`, the stable temporary work directory, and final delivery archives are mutable.
- Use `--force` only after source/rule changes or a demonstrated stale-cache problem.

## Lower-level implementation

`scripts/unified_workflow.py` is the only public entrypoint. It invokes:

- `scripts/process_batch.py` for extraction, audit, repair, verification, and visual evidence;
- `scripts/word_manual_pipeline_stable.ps1` for description DOCX repair/export;
- `scripts/repair_code_headers.ps1` for code DOCX header repair/export;
- `scripts/package_approved_modified.py` for approved-folder ZIP/RAR creation and integrity testing.

Use specialized scripts only to diagnose an explicit failure returned by the unified entrypoint.

# Acceptance rules

## Table of contents

1. Batch and naming
2. Code PDF
3. Description DOCX/PDF
4. Interface screenshot and login review
5. Content correspondence
6. TXT
7. Mutation safety
8. Final verification and reporting

## Batch and naming

- At the start, scan the requested batch root for ZIP/RAR/7z inputs. Recursively expand archives found inside outer archives until the actual TXT/PDF/DOCX project folders are visible; never count one outer archive as one project when it contains multiple inner project archives.
- Reuse unchanged archives by SHA-256. If the archive changed or an untracked same-name folder already exists, preserve the old folder as a timestamped backup before switching the newly extracted folder into place.
- Use each same-name folder as a batch root. If it contains already-extracted TXT/PDF/DOCX materials rather than inner archives, copy that source tree once into its read-only `解压版`, then continue normally.
- Derive the authoritative software name from the TXT filename without `.txt`.
- Expect one TXT, one code PDF, one description DOCX, and one paired description PDF per program unless the batch explicitly differs.
- Keep archives and the extracted audit tree unchanged. Work only in the modified copy.
- Record exact relative paths and SHA-256 values before repair.
- Immediately after extraction, inspect every code PDF's page 1 and last page, retain confirmed code errors with exact locations, and generate review evidence.
- Complete deterministic audit, repair, export, and verification for every project in the modified copy, then place code edges, interface/login evidence, TXT/content errors, and processed description-PDF pages 1–2 into one consolidated HTML review. Persist the user's single pass/fail decision; only approved projects enter delivery packaging.

## Code PDF

- Inspect only page 1 and the last page visually; do not read or render middle pages.
- Record total page count. Approximately 100 pages is expected; use 70–130 only as an initial candidate range unless the user gives a tighter range.
- Confirm code-like structure, indentation, common punctuation, brackets, quotation marks, and consistent formatting.
- Confirm a visible running header and page number on both inspected pages.
- Check software-name and `V1.0` consistency where extractable.
- Treat automated punctuation and bracket counts as candidate evidence, not a standalone error, because PDF extraction can reorder code.
- Ignore business correctness and middle-page code content.
- Automatically persist and repair code headers/page numbers when extracted first/last-page evidence confirms a missing title, `V1.0`, or expected page number. Retain visual code-format judgment in the unified review.
- Require the paired code DOCX for automatic repair; if it is absent, report the code PDF and the missing source document at `代码页眉修复`.
- Normalize every code-DOCX section header to left-aligned `软件名称 V1.0` and a right-aligned dynamic `PAGE` field, using 9 pt SimSun/宋体. Save the changed code DOCX and export over the paired code PDF in the modified copy.
- Back up the code DOCX/PDF pair and roll back that pair on open, save, export, or QA failure. Preserve the complete code-DOCX main-document text byte-logically by normalized text digest.
- After export, parse and render only the repaired code PDF's first and last pages. Require title, `V1.0`, expected page numbers, and Word/PDF page-count agreement before delivery.

## Description DOCX/PDF

- Require one valid cover on page 1 containing the exact software name and one label from: 使用说明书、使用手册、用户手册、用户使用手册、用户操作手册、用户操作说明、操作说明书、操作手册、软件设计说明书、软件设计说明.
- Compare cover-title candidates after removing spaces, tabs, and line breaks, so a wrapped title remains equivalent to the authoritative software name.
- The visible cover title block is distinct from the running header.
- Center the software name and manual label horizontally. Leave generous blank space before the software name so it sits in the middle-upper region rather than at the page top. Place the manual label around the page's vertical middle/lower-middle.
- Use exactly two blank paragraphs between newly created visible lines.
- Add a visible paragraph gap before the manual label even when normalizing an existing cover. Target the software-name baseline at 30%–56% of page height, the manual-label baseline at 48%–72%, and at least 48 pt of vertical separation. Treat unavailable Word coordinates as a visual-QA item rather than inventing a failure.
- Use SimSun/宋体, bold, and the same integer point size for both visible lines.
- Start at 26 pt and reduce both lines together until the complete software name occupies exactly one line. Do not reduce below 14 pt without reporting a layout exception.
- Remove inherited numbering and paragraph indents from cover lines.
- The cover follows the supplied sample and must show the same running header and page number as body pages.
- Require the software name plus `V1.0` in every page's running header and require a dynamic `PAGE` field. Do not retain a separate blank first-page header.
- Do not accept a static typed page number as the dynamic PAGE field.
- Use the TXT filename stem as the authoritative project name. Remove only paired ASCII or Chinese double quotes that wrap that exact name in any Word story; preserve all unrelated quotation marks.
- Identify normal body paragraphs using Word's effective outline level and typography. Compare the first body paragraph with the modal font and size of later body paragraphs, and copy only font name and size when they differ.
- A duplicate old title block may share its page with a TOC or body. Delete only the exact software-name/manual-label paragraphs and adjacent cover-only blanks. Preserve the page, page break, TOC, tables, images, and body content.
- If an old duplicate cover occupies an otherwise empty page, remove that cover-only page boundary. If the page contains any body text, table, or image, retain the page and remove only the duplicate title/label text.
- Replace the paired description PDF only after the modified DOCX saves and exports successfully.
- Check the rendered PDF as well as DOCX structure; Word-effective formatting and the exported appearance are the final typography evidence.

## Interface screenshot and login review

- Extract substantial screenshots from each description DOCX in body relationship order. Exclude header/footer images, logos, icons, decorations, and images smaller than the substantial-image thresholds.
- Create exactly one lossless `interfaces_<software-name>.png` per program. Keep every embedded screenshot at its original pixel dimensions without downsampling or upsampling. Label every card with screenshot order, nearby paragraph number, and source dimensions.
- Use programmatic checks first: compare aspect ratios and palette signatures against the project median; draw a red border around outliers so visual review can focus on them.
- Combine projects in source order, five projects per lossless PNG review group. If a five-project canvas exceeds the image-library dimension or pixel safety limit, split that same group into numbered parts without resampling project sheets. Place every group and per-project interface sheet in the unified HTML. The user personally judges whether screenshots belong together and whether login pages are acceptable as part of the single project pass/fail decision; the agent performs no separate login analysis or report.
- Check visible text for replacement boxes, random symbols, mojibake, broken Chinese glyphs, or mixed encodings. A clearly garbled screenshot is a confirmed error.
- Treat responsive layouts, modal dialogs, different modules, selected states, and coherent dark/light themes as normal variation when the underlying design system remains consistent.
- Treat metric outliers only as candidates. Report an interface-image error only when the user rejects or identifies it from the initial combined sheet.
- For every confirmed inconsistency, report the DOCX filename, screenshot order, paragraph position, conflicting visual traits, and the full per-program interface sheet. Show or link that sheet in the final response.
- With only one substantial screenshot, skip cross-image consistency comparison but still check it for visible garbled text and obvious topic/product mismatch.

## Content correspondence

- Confirm that the manual topic corresponds to the software name.
- Use DOCX text, chapter titles, image-adjacent text, and cross-program title leakage for programmatic nomination.
- Scan the complete DOCX text programmatically instead of converting the whole PDF to Markdown or rendering every PDF page. Use strong repeated cross-domain terms to confirm only obvious mismatches, such as an earthquake product containing a baby-education manual. Do not report ordinary generic wording as a mismatch.

## TXT

- Detect BOM and encoding once. Prefer UTF-8 or UTF-8 BOM unless the batch convention specifies otherwise.
- Detect replacement characters, control characters, and obvious mojibake.
- Check basic key/value field structure, empty required-looking values, duplicate fields, software name, version form, and obvious wrong categories/options.
- Cross-check the software name against the TXT filename stem and description title.
- For `主要功能` or `软件主要功能`, report any digit or ASCII/full-width parenthesis as an error. Also report deterministic fluency/compliance defects: empty content, missing Chinese text, malformed punctuation, or no recognizable functional action.
- Scan rather than deeply reading narrative text.
- Treat GB18030 as a candidate unless UTF-8 is explicitly required; treat undecodable or visibly corrupted content as a clear error.

## Mutation safety

- Preserve the source archive format for approved delivery: ZIP input packages return a same-name ZIP; RAR input packages return a same-name RAR.
- Create the modified copy from all inventoried project folders before repair; preserve the source tree and package only the projects approved in the consolidated review.
- Back up each actionable DOCX/PDF pair before opening it for changes.
- Use one Word application per batch and one open/save/export cycle per actionable document.
- Do not open normal documents in Word when static and PDF checks already prove acceptance.
- Update header/footer page fields only; do not refresh unrelated TOCs or body fields.
- Roll back the individual pair on a Word, save, export, or QA failure.
- Preserve non-target files byte-for-byte.
- Never use hard links between extracted and modified copies.

## Final verification and reporting

- Before delivery packaging, generate pages 1 and 2 from every processed description PDF and include them in the consolidated HTML alongside every other human-review item. Package only after the user submits one explicit list of approved project sequence numbers.

- Verify DOCX ZIP validity, cover block, effective Word typography, title line count, duplicate-block count, running header, and PAGE field.
- Verify the paired PDF parses, has at least one page, and shows the exact title on page 1.
- Compare body content before and after; allow only intended cover insertion or exact duplicate-title-block removal.
- Record per-stage and per-document timings for regression.
- Cache only by SHA-256 plus pipeline/rule version. Reuse cached visual decisions only when underlying hashes are unchanged.
- Lead the final response with destination and totals.
- List modifications applied.
- List only remaining confirmed content, TXT, image, code-format, or QA errors with exact program, file, page/paragraph/field/image location, and offending value.
- Omit all normal-by-file entries.
- Treat the consolidated HTML—containing processed-description-PDF pages 1–2, code edges, interface/login evidence, and programmatic errors—as the sole pre-packaging approval gate.
- After layout confirmation, delete confirmed delivery-temporary files (`*.tmp`, `*.temp`, and Word `~$*` files) from the modified tree, then create and integrity-test one archive inside `修改版`, named after the batch root. Include only the user-approved processed program directories and exclude backup/hidden directories. Preserve the incoming format (ZIP to ZIP, RAR to RAR); remove a same-name archive of the other format only after the selected archive passes, then report its path, file count, removed temporary files, and SHA-256.

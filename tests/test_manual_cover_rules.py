from __future__ import annotations

import importlib.util
import base64
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
# 被测模块按文件路径加载，需要把 scripts 加入 sys.path，
# 否则被测模块内部的同级 import（如 archive_types）无法解析。
sys.path.insert(0, str(ROOT / "scripts"))
SPEC = importlib.util.spec_from_file_location("process_batch_cover_tests", ROOT / "scripts" / "process_batch.py")
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


def make_docx(path: Path, lines: list[str]) -> None:
    document = etree.Element(f"{{{MODULE.W}}}document", nsmap={"w": MODULE.W})
    body = etree.SubElement(document, f"{{{MODULE.W}}}body")
    for line in lines:
        paragraph = etree.SubElement(body, f"{{{MODULE.W}}}p")
        if line:
            run = etree.SubElement(paragraph, f"{{{MODULE.W}}}r")
            text = etree.SubElement(run, f"{{{MODULE.W}}}t")
            text.text = line
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>",
        )
        archive.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8"))


class ManualCoverAuditTests(unittest.TestCase):
    def inspect(self, lines: list[str], title: str = "示例软件") -> dict:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "manual.docx"
            make_docx(path, lines)
            return MODULE.inspect_docx(path, title, check_login=False)

    def test_cover_label_may_follow_empty_paragraphs(self):
        audit = self.inspect(["示例软件", "", "", "使用说明书", "第一章 功能介绍"])

        self.assertTrue(audit["cover_found"])
        self.assertEqual(audit["cover_label"], "使用说明书")

    def test_cover_scan_stops_at_first_visible_non_label(self):
        audit = self.inspect(["示例软件", "第一章 功能介绍", "使用说明书"])

        self.assertFalse(audit["cover_found"])
        self.assertEqual(audit["cover_blocks"], [])
        self.assertEqual([row["text"] for row in audit["residual_manual_labels"]], ["使用说明书"])

    def test_incomplete_manual_label_is_not_completed_or_accepted(self):
        audit = self.inspect(["示例软件", "用户操作", "第一章 功能介绍"])

        self.assertFalse(MODULE.is_manual_label_candidate("用户操作"))
        self.assertFalse(audit["cover_found"])
        self.assertIn("用户操作", audit["body_text"])


class _Callable:
    def __init__(self, result=None):
        self.result = result

    def __call__(self, *_args, **_kwargs):
        return self.result


class _TimedOutKernel32:
    def __init__(self):
        self.CreateMutexW = _Callable(1)
        self.WaitForSingleObject = _Callable(0x102)
        self.ReleaseMutex = _Callable()
        self.CloseHandle = _Callable()


class WordQueueTests(unittest.TestCase):
    def test_word_queue_times_out_instead_of_waiting_forever(self):
        reporter = mock.Mock()
        kernel32 = _TimedOutKernel32()
        fake_ctypes = mock.Mock(windll=mock.Mock(kernel32=kernel32))

        with mock.patch.object(MODULE.os, "name", "nt"), \
                mock.patch.object(MODULE.ctypes, "windll", fake_ctypes.windll, create=True), \
                mock.patch.object(MODULE.time, "monotonic", side_effect=[0.0, 1.0]):
            with self.assertRaisesRegex(TimeoutError, "等待 Word 处理队列超时"):
                with MODULE.word_automation_lock(reporter, "说明书处理", timeout_seconds=0):
                    self.fail("mutex must not be acquired")

        reporter.update.assert_called_once_with("等待其他任务完成 Word 导出", 0, "说明书处理")


class PowerShellCoverSafetyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.script = (ROOT / "scripts" / "word_manual_pipeline_stable.ps1").read_text(encoding="utf-8-sig")

    def test_residual_label_with_content_or_objects_is_report_only(self):
        start = self.script.index("if ($null -ne $residualLabel)")
        end = self.script.index("if ($null -ne $isolated)", start)
        branch = self.script[start:end]

        self.assertIn("$otherText.Count -gt 0 -or $hasObjects", branch)
        self.assertIn("{ break }", branch)
        self.assertLess(branch.index("{ break }"), branch.index(".Delete()"))

    def test_cover_boundary_includes_images_and_tables(self):
        start = self.script.index("function Ensure-CoverPageBoundary")
        end = self.script.index("function Remove-BodyLeadingBlankParagraphs", start)
        function = self.script[start:end]

        self.assertIn("$Document.InlineShapes", function)
        self.assertIn("$Document.Tables", function)
        self.assertIn("$extra.Count -eq 0 -and -not $hasObjects", function)

    def test_cover_keeps_existing_page_dimensions(self):
        start = self.script.index("function Add-Cover")
        end = self.script.index("function Remove-DuplicateCoverBlocks", start)
        function = self.script[start:end]

        self.assertNotIn("PageWidth", function)
        self.assertNotIn("PageHeight", function)
        self.assertNotIn("wdPaperA4", function)

    def test_two_line_title_is_allowed_but_short_last_line_is_flagged(self):
        self.assertIn("if ($lines -le 2 -and -not $orphanedLastLine)", self.script)
        self.assertIn('if ($coverAudit.Lines -gt 2)', self.script)
        self.assertIn('if ($coverAudit.OrphanedLastLine)', self.script)

    def test_title_only_cover_is_accepted_without_inventing_subtitle(self):
        start = self.script.index("function Set-CoverTypography")
        end = self.script.index("function Remove-CoverVersion", start)
        function = self.script[start:end]

        self.assertIn("if ($null -eq $titleItem)", function)
        self.assertNotIn("$null -eq $titleItem -or $null -eq $labelItem", function)
        self.assertIn("LabelPresent", function)
        self.assertIn("$between | Sort-Object Index -Descending", function)
        self.assertIn("$_.Text -ne \"\"", function)

    def test_duplicate_qa_requires_a_standalone_cover_page(self):
        start = self.script.index("function Test-DuplicateCoverBlock")
        end = self.script.index("function Get-CoverLastLineUnitCount", start)
        function = self.script[start:end]

        self.assertIn("$otherText.Count -eq 0", function)
        self.assertIn("Test-PageHasObjects", function)
        self.assertIn("Find-IsolatedCombinedCoverFragment", function)

    def test_redundant_post_cover_breaks_are_removed_after_first_boundary(self):
        start = self.script.index("function Remove-RedundantPostCoverFrontMatter")
        end = self.script.index("function Story-HasContent", start)
        function = self.script[start:end]

        self.assertIn("$protectedBoundaryFound", function)
        self.assertIn("$Document.Sections.Item($sectionIndex).Range.End", function)
        self.assertIn("Test-CombinedCoverFragment", function)
        self.assertIn("$realBody.Page -lt 3", function)
        self.assertIn("Sort-Object Index -Descending", function)

    def test_real_output_post_cover_fragments_match_only_front_matter(self):
        function_names = (
            "Clean-Text", "Compact-Text", "Strip-OrdinalPrefix",
            "Test-CoverTitle", "Test-ManualLabel", "Test-PostCoverFragment",
        )
        definitions = []
        for name in function_names:
            start = self.script.index(f"function {name}(")
            end = self.script.find("\nfunction ", start + 1)
            definitions.append(self.script[start:end if end != -1 else None])
        cases = [
            ("弘元绿能金刚线蓝宝石切片机控制系统软件", "弘元绿能金刚线蓝宝石切片机控制系统软件", True),
            ("用户手册", "医院人力资源综合管理系统", True),
            ("技术说明手册", "建筑项目数据综合分析平台", True),
            ("文档编号:ARCH-ENERGY-2024-UM01 系统状态:正在运行 机密程度:内部使用", "建筑项目数据综合分析平台", True),
            ("目录", "建筑项目数据综合分析平台", False),
            ("1. 系统概述", "建筑项目数据综合分析平台", False),
            ("如上图所示可以看到切割过程中，参数在发生变化。", "弘元绿能金刚线蓝宝石切片机控制系统软件", False),
            ("建筑项目数据综合分析平台是一款面向现代化大型建筑工程的数字化能效管理平台", "建筑项目数据综合分析平台", False),
        ]
        commands = [*definitions]
        for text, title, expected in cases:
            commands.append(f"if ((Test-PostCoverFragment '{text}' '{title}') -ne ${str(expected).lower()}) {{ throw 'fragment rule mismatch' }}")
        encoded = base64.b64encode("\n".join(commands).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_content_page_cleanup_deletes_only_leading_fragment_paragraph(self):
        start = self.script.index("function Remove-DuplicateCoverBlocks")
        end = self.script.index("function Test-DuplicateCoverBlock", start)
        function = self.script[start:end]
        self.assertIn("Test-PostCoverFragment", function)
        self.assertIn("$leadingFragment.Page -eq 2", function)
        self.assertIn("$Document.Range($leadingFragment.Start, $leadingFragment.End).Delete()", function)
        self.assertIn("$wdWithInTable", function)
        self.assertIn("Test-RangeHasProtectedContent $Document $fragmentRange $true", function)
        self.assertIn('or (Test-DuplicateCoverBlock $doc $title)', self.script)

    def test_blank_range_protection_rules_execute(self):
        start = self.script.index("function Test-RangeHasProtectedContent(")
        end = self.script.index("\nfunction ", start + 1)
        function = self.script[start:end]
        checks = r'''
function New-Range($text = '', $pageBreak = $false) {
    return [pscustomobject]@{ Start = 10; End = 20; Text = $text; ParagraphFormat = [pscustomobject]@{ PageBreakBefore = $pageBreak } }
}
function New-Document {
    return [pscustomobject]@{ Sections = @(); InlineShapes = @(); Shapes = @(); Tables = @() }
}
$doc = New-Document
if (Test-RangeHasProtectedContent $doc (New-Range)) { throw 'plain blank was protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)))) { throw 'page break was not protected' }
if (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)) $false $true) { throw 'cover-only page break was protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range ([string][char]14)) $false $true $true)) { throw 'column break was not protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range '' $true) $true)) { throw 'PageBreakBefore was not protected' }
$doc.Sections = @([pscustomobject]@{ Range = [pscustomobject]@{ End = 19 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'section end was not protected' }
if (-not (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)) $false $true)) { throw 'section end was lost when allowing page break' }
if (Test-RangeHasProtectedContent $doc (New-Range ([string][char]12)) $false $true $true) { throw 'standalone cover section break was protected' }
$doc.Sections = @()
$doc.InlineShapes = @([pscustomobject]@{ Range = [pscustomobject]@{ Start = 15; End = 16 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'inline image was not protected' }
$doc.InlineShapes = @()
$doc.Shapes = @([pscustomobject]@{ Anchor = [pscustomobject]@{ Start = 15 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'floating image was not protected' }
$doc.Shapes = @()
$doc.Tables = @([pscustomobject]@{ Range = [pscustomobject]@{ Start = 15; End = 18 } })
if (-not (Test-RangeHasProtectedContent $doc (New-Range))) { throw 'table was not protected' }
'''
        encoded = base64.b64encode((function + "\n" + checks).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_floating_shape_counts_as_page_object(self):
        start = self.script.index("function Test-PageHasObjects(")
        end = self.script.index("\nfunction ", start + 1)
        function = self.script[start:end]
        checks = r'''
$wdActiveEndPageNumber = 3
$anchor = [pscustomobject]@{}
$anchor | Add-Member -MemberType ScriptMethod -Name Information -Value { param($kind) return 2 }
$doc = [pscustomobject]@{ InlineShapes = @(); Tables = @(); Shapes = @([pscustomobject]@{ Anchor = $anchor }) }
if (-not (Test-PageHasObjects $doc 2)) { throw 'floating shape on page 2 was missed' }
if (Test-PageHasObjects $doc 1) { throw 'floating shape was assigned to wrong page' }
'''
        encoded = base64.b64encode((function + "\n" + checks).encode("utf-16le")).decode("ascii")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_cover_style_uses_com_double_values(self):
        start = self.script.index("function Set-CoverTypography")
        end = self.script.index("function Remove-CoverVersion", start)
        function = self.script[start:end]
        self.assertIn("$titleRange.Font.Size = [double]$size", function)
        self.assertIn("$labelRange.Font.Size = [double]$size", function)
        self.assertIn("$labelRange.ParagraphFormat.SpaceBefore = [double]48", function)


if __name__ == "__main__":
    unittest.main()

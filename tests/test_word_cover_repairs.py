from __future__ import annotations

import re
import subprocess
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "word_manual_pipeline_stable.ps1"


def extract_function(name: str) -> str:
    source = SCRIPT.read_text(encoding="utf-8-sig")
    match = re.search(rf"(?ms)^function {re.escape(name)}\b.*?(?=^function |\Z)", source)
    if match is None:
        raise AssertionError(f"PowerShell function not found: {name}")
    return match.group(0)


def run_powershell(source: str) -> None:
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "isolated_mock.ps1"
        # Windows PowerShell 5 needs a BOM to read Chinese test literals as UTF-8.
        path.write_text(source, encoding="utf-8-sig")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )
    if result.returncode != 0:
        raise AssertionError(result.stderr or result.stdout)


class WordCoverRepairMockTests(unittest.TestCase):
    def test_missing_bookmark_fields_preserve_cached_results_and_fail_closed(self):
        functions = "\n".join(
            extract_function(name)
            for name in (
                "Get-BookmarkFieldDescriptor",
                "Resolve-MissingBookmarkField",
                "Repair-MissingBookmarkFields",
            )
        )
        harness = r'''
$script:BookmarksByName = @{ '_ExistingRef' = $true }
$script:Fields = @()
$script:UnlinkedCodes = @()
function New-MockField([int]$type, [string]$code, [string]$result, [bool]$unlinkThrows = $false) {
    $field = [pscustomobject]@{
        Type = $type
        Code = [pscustomobject]@{ Text = $code }
        Result = [pscustomobject]@{ Text = $result }
        UnlinkThrows = $unlinkThrows
    }
    $field | Add-Member ScriptMethod Unlink {
        if ($this.UnlinkThrows) { throw 'mock unlink failure' }
        $script:UnlinkedCodes += [string]$this.Code.Text
    }
    return $field
}
function New-MockDocument($fields) {
    $script:Fields = @($fields)
    $collection = [pscustomobject]@{ Count = $script:Fields.Count }
    $collection | Add-Member ScriptMethod Item { param([int]$index) return $script:Fields[$index - 1] }
    $bookmarks = [pscustomobject]@{}
    $bookmarks | Add-Member ScriptMethod Exists { param([string]$name) return $script:BookmarksByName.ContainsKey($name) }
    return [pscustomobject]@{ Fields = $collection; Bookmarks = $bookmarks }
}
$script:UnlinkedCodes = @()
$missing = New-MockField 37 ' PAGEREF _Toc32659 \h' '3'
$existing = New-MockField 3 ' REF _ExistingRef \h' '8'
$doc = New-MockDocument @($missing, $existing)
$changes = @(Repair-MissingBookmarkFields $doc)
if ($changes.Count -ne 1) { throw "Expected one broken reference to be unlinked; got $($changes.Count)" }
if ($changes[0].type -ne 'PAGEREF' -or $changes[0].target -ne '_Toc32659' -or $changes[0].cached_result -ne '3') {
    throw 'Unlinked-field evidence did not retain its type, target, and cached display value'
}
if ($script:UnlinkedCodes.Count -ne 1 -or $script:UnlinkedCodes[0] -notmatch 'PAGEREF _Toc32659') { throw 'The missing-target field was not unlinked' }
if ($existing.Result.Text -ne '8' -or $changes[0].target -eq '_ExistingRef') { throw 'A field with an existing target was modified' }

$empty = New-MockField 37 ' PAGEREF _MissingEmpty \h' ''
$doc = New-MockDocument @($empty)
$failedForEmptyCache = $false
try { [void](Repair-MissingBookmarkFields $doc) } catch { $failedForEmptyCache = $_.Exception.Message -match '缓存结果为空' }
if (-not $failedForEmptyCache -or $script:UnlinkedCodes.Count -ne 1) { throw 'An empty cached result was not rejected before unlinking' }

$unlinkFailure = New-MockField 3 ' REF _MissingUnlink \h' '5' $true
$doc = New-MockDocument @($unlinkFailure)
$failedForUnlinkError = $false
try { [void](Repair-MissingBookmarkFields $doc) } catch { $failedForUnlinkError = $_.Exception.Message -match '无法保留 REF 字段' }
if (-not $failedForUnlinkError) { throw 'A Word field unlink error was not propagated as a failure' }

$malformed = New-MockField 37 ' PAGEREF \h' '2'
$doc = New-MockDocument @($malformed)
$failedForParseError = $false
try { [void](Repair-MissingBookmarkFields $doc) } catch { $failedForParseError = $_.Exception.Message -match '无法解析 PAGEREF' }
if (-not $failedForParseError) { throw 'A malformed field target was not propagated as a failure' }
'''
        run_powershell(functions + "\n" + harness)

    def test_mixed_page_removes_only_exact_standalone_duplicate_title(self):
        remove_function = extract_function("Remove-DuplicateCoverBlocks")
        harness = r'''
$wdWithInTable = 12
$wdActiveEndPageNumber = 3
$script:MockItems = @(
    [pscustomobject]@{ Page = 3; Index = 1; Text = '示例软件'; Start = 90; End = 94; InTable = $true },
    [pscustomobject]@{ Page = 3; Index = 2; Text = '示例软件'; Start = 100; End = 104; InTable = $false },
    [pscustomobject]@{ Page = 3; Index = 3; Text = '示例软件用户指南'; Start = 110; End = 119; InTable = $false },
    [pscustomobject]@{ Page = 3; Index = 4; Text = '示例软件是一款用于管理设备的工具。'; Start = 120; End = 139; InTable = $false }
)
$script:DeletedStarts = @()
$script:DeletedEnds = @()
function Get-FrontParagraphs($Document, [int]$MaxPage = 5) {
    return @($script:MockItems | Where-Object { $_.Page -le $MaxPage })
}
function Compact-Text([string]$Text) { return ($Text -replace '\s+', '') }
function Test-PostCoverFragment([string]$Text, [string]$Title) { return $Text -ceq $Title }
function Find-CoverBlock { return $null }
function Find-IsolatedTitleCover { return $null }
function Find-IsolatedManualLabelCover { return $null }
function Find-IsolatedCombinedCoverFragment { return $null }
function Find-ResidualManualLabel { return $null }
function Test-PageHasObjects($Document, [int]$Page) { return $true }
function Test-RangeHasProtectedContent($Document, $Range, [bool]$CheckParagraphBreakBefore = $false, [bool]$AllowPageBreak = $false, [bool]$AllowSectionBreak = $false) { return ([int]$Range.Start -eq 100 -and [int]$Range.End -eq 104) }
$image = [pscustomobject]@{ Kind = 'image' }
$table = [pscustomobject]@{ Kind = 'table' }
$doc = [pscustomobject]@{ InlineShapes = @($image); Tables = @($table); Shapes = @() }
$doc | Add-Member ScriptMethod Range {
    param($start, $end)
    $range = [pscustomobject]@{ Start = $start; End = $end }
    $range | Add-Member ScriptMethod Information { param($kind) return 0 }
    $range | Add-Member ScriptMethod Delete {
        $script:DeletedStarts += [int]$this.Start
        $script:DeletedEnds += [int]$this.End
        $script:MockItems = @($script:MockItems | Where-Object { $_.Start -ne [int]$this.Start })
    }
    return $range
}
$removed = Remove-DuplicateCoverBlocks $doc '示例软件'
if ($script:DeletedStarts.Count -ne 1 -or $script:DeletedStarts[0] -ne 100) {
    throw "Expected deletion of the exact title paragraph at 100 only; removed starts: $($script:DeletedStarts -join ',')"
}
if ($script:DeletedEnds[0] -ne 103) { throw 'The title cleanup removed its paragraph mark instead of preserving the section boundary' }
if (@($script:MockItems | Where-Object { $_.Text -eq '示例软件是一款用于管理设备的工具。' }).Count -ne 1) {
    throw 'Body paragraph was not preserved'
}
if (@($script:MockItems | Where-Object { $_.Text -eq '示例软件用户指南' }).Count -ne 1) {
    throw 'Near-match title was not preserved'
}
if (@($script:MockItems | Where-Object { $_.Text -eq '示例软件' -and $_.InTable }).Count -ne 1) {
    throw 'Table-cell title was not preserved'
}
if ($doc.InlineShapes.Count -ne 1 -or $doc.Tables.Count -ne 1) {
    throw 'Image or table on the mixed page was not preserved'
}
$script:DeletedStarts = @()
$script:MockItems = @(
    [pscustomobject]@{ Page = 2; Index = 1; Text = '示例软件'; Start = 200; End = 204; InTable = $true; Style = 'Normal' },
    [pscustomobject]@{ Page = 2; Index = 2; Text = '目录正文'; Start = 205; End = 209; InTable = $false; Style = 'Normal' }
)
if ((Remove-DuplicateCoverBlocks $doc '示例软件') -ne 0 -or $script:DeletedStarts.Count -gt 0) { throw 'A table-cell title on page 2 was removed' }
$script:MockItems[0].InTable = $false
$script:MockItems[0].Style = 'TOC1'
if ((Remove-DuplicateCoverBlocks $doc '示例软件') -ne 0 -or $script:DeletedStarts.Count -gt 0) { throw 'A table-of-contents title on page 2 was removed' }
'''
        run_powershell(remove_function + "\n" + harness)

    def test_final_duplicate_audit_catches_removal_skips_but_excludes_tables_and_toc(self):
        audit_function = extract_function("Test-DuplicateCoverBlock")
        harness = r'''
function Get-FrontParagraphs($Document, [int]$MaxPage = 5) { return @($script:MockItems) }
function Compact-Text([string]$Text) { return ($Text -replace '\s+', '') }
function Test-PostCoverFragment([string]$Text, [string]$Title) { return $false }
function Find-CoverBlock { return $null }
function Find-IsolatedTitleCover { return $null }
function Find-IsolatedManualLabelCover { return $null }
function Find-IsolatedCombinedCoverFragment { return $null }
function Find-ResidualManualLabel { return $null }
function Test-PageHasObjects($Document, [int]$Page) { return $false }
$doc = [pscustomobject]@{ InlineShapes = @(); Tables = @(); Shapes = @() }
$script:MockItems = @([pscustomobject]@{ Page = 3; Index = 1; Text = '示例 软件'; Start = 100; End = 105; InTable = $false; Style = 'Normal' })
if (-not (Test-DuplicateCoverBlock $doc '示例软件')) { throw 'An exact standalone title skipped by cleanup was not reported' }
$script:MockItems[0].InTable = $true
if (Test-DuplicateCoverBlock $doc '示例软件') { throw 'A table-cell title was reported as a removable duplicate' }
$script:MockItems[0].InTable = $false
$script:MockItems[0].Style = 'TOC1'
if (Test-DuplicateCoverBlock $doc '示例软件') { throw 'A table-of-contents item was reported as a duplicate title' }
$script:MockItems[0].Style = 'Normal'
$script:MockItems[0].Text = '示例软件 V1.0'
if (Test-DuplicateCoverBlock $doc '示例软件') { throw 'A versioned near-match was reported as an exact duplicate title' }
$script:MockItems = @([pscustomobject]@{ Page = 5; Index = 1; Text = '示例软件'; Start = 500; End = 504; InTable = $false; Style = 'Normal' })
if (-not (Test-DuplicateCoverBlock $doc '示例软件')) { throw 'An exact title on page 5 was outside the duplicate audit' }
$script:MockItems[0].Page = 6
if (Test-DuplicateCoverBlock $doc '示例软件') { throw 'An exact title after page 5 was included in the duplicate audit' }
'''
        run_powershell(audit_function + "\n" + harness)
        script = SCRIPT.read_text(encoding="utf-8-sig")
        self.assertIn("$unresolvedBlankPages.Count -gt 0", script)
        self.assertIn('全书仍存在无法安全修复的空白页:', script)

    def test_unclassified_blank_page_break_stays_unresolved(self):
        repair_function = extract_function("Repair-WordBlankPages")
        harness = r'''
$wdActiveEndPageNumber = 3
$wdWithInTable = 12
$script:AuditCalls = 0
$script:DeletedStarts = @()
function Get-WordPageAudit($Document) {
    $script:AuditCalls++
    return @(
        [pscustomobject]@{ page = 1; has_body_content = $true; visible_text = @('Page 1 body'); page_breaks = 0; section_breaks = 0; page_break_before = 0; section_starts = @(); tables = 0; inline_images = 0; floating_shapes = 0 },
        [pscustomobject]@{ page = 2; has_body_content = $false; visible_text = @(); page_breaks = 1; section_breaks = 0; page_break_before = 0; section_starts = @(); tables = 0; inline_images = 0; floating_shapes = 0 }
    )
}
$paragraphRange = [pscustomobject]@{ Start = 200; End = 202; Text = "`f`r"; Page = 2; ParagraphFormat = [pscustomobject]@{ PageBreakBefore = $false } }
$paragraphRange | Add-Member ScriptMethod Information { param($kind) if ($kind -eq 12) { return $false }; return $this.Page }
$paragraphRange | Add-Member ScriptMethod Delete { $script:DeletedStarts += [int]$this.Start }
$paragraph = [pscustomobject]@{ Range = $paragraphRange }
$paragraphs = [pscustomobject]@{ Count = 1 }
$paragraphs | Add-Member ScriptMethod Item { param($index) return $paragraph }
$doc = [pscustomobject]@{ Sections = @(); Paragraphs = $paragraphs; Content = [pscustomobject]@{ End = 300 }; InlineShapes = @(); Shapes = @(); Tables = @() }
$doc | Add-Member ScriptMethod Range { throw 'An unclassified blank page must not be edited' }
$doc | Add-Member ScriptMethod Repaginate { }
function Test-RangeHasProtectedContent($Document, $Range, [bool]$CheckParagraphBreakBefore = $false, [bool]$AllowPageBreak = $false, [bool]$AllowSectionBreak = $false) { return $false }
function Repair-TerminalBlankParagraph($Document, $Audit) { return [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = $null } }
$result = Repair-WordBlankPages $doc
if ($result.Changed -or $script:DeletedStarts.Count -ne 0) { throw 'An unclassified blank-page break was deleted' }
if ($result.UnresolvedPages.Count -ne 1 -or $result.UnresolvedPages[0].page -ne 2) { throw 'An unclassified blank page was not reported as unresolved' }
'''
        run_powershell(repair_function + "\r\n" + harness)

    def test_duplicate_page_break_at_next_page_section_start_is_removed(self):
        repair_function = extract_function("Repair-WordBlankPages")
        harness = r'''
$wdActiveEndPageNumber = 3
$wdWithInTable = 12
$script:DeletedStarts = @()
$script:AuditCalls = 0
$script:ParagraphRange = [pscustomobject]@{ Start = 17; End = 19; Text = "`f`r"; Page = 2; ParagraphFormat = [pscustomobject]@{ PageBreakBefore = $false } }
$script:ParagraphRange | Add-Member ScriptMethod Information { param($kind) if ($kind -eq 12) { return $false }; return $this.Page }
$paragraph = [pscustomobject]@{ Range = $script:ParagraphRange }
$paragraphs = [pscustomobject]@{ Count = 1 }
$paragraphs | Add-Member ScriptMethod Item { param($index) return $paragraph }
$section = [pscustomobject]@{ Range = [pscustomobject]@{ Start = 17; End = 100 }; PageSetup = [pscustomobject]@{ SectionStart = 2 } }
$doc = [pscustomobject]@{ Sections = @($section); Paragraphs = $paragraphs; Content = [pscustomobject]@{ End = 100 }; InlineShapes = @(); Shapes = @(); Tables = @() }
$doc | Add-Member ScriptMethod Range {
    param($start, $end)
    $range = [pscustomobject]@{ Start = [int]$start; End = [int]$end; Text = if ([int]$end -eq ([int]$start + 1)) { [string][char]12 } else { '' }; Page = 2 }
    $range | Add-Member ScriptMethod Information { param($kind) if ($kind -eq 12) { return $false }; return $this.Page }
    $range | Add-Member ScriptMethod Delete {
        $script:DeletedStarts += [int]$this.Start
        $script:ParagraphRange.Text = [string][char]13
        $script:ParagraphRange.End = 18
    }
    return $range
}
$doc | Add-Member ScriptMethod Repaginate { }
function Get-WordPageAudit($Document) {
    $script:AuditCalls++
    $repaired = $script:DeletedStarts.Count -gt 0
    return @(
        [pscustomobject]@{ page = 1; has_body_content = $true; visible_text = @('Page 1 body'); page_breaks = 0; section_breaks = 0; page_break_before = 0; section_starts = @(); tables = 0; inline_images = 0; floating_shapes = 0 },
        [pscustomobject]@{ page = 2; has_body_content = $repaired; visible_text = $(if ($repaired) { @('Section body') } else { @() }); page_breaks = $(if ($repaired) { 0 } else { 1 }); section_breaks = 1; page_break_before = 0; section_starts = @('new_page'); tables = 0; inline_images = 0; floating_shapes = 0 }
    )
}
function Test-RangeHasProtectedContent($Document, $Range, [bool]$CheckParagraphBreakBefore = $false, [bool]$AllowPageBreak = $false, [bool]$AllowSectionBreak = $false) { return $false }
function Repair-TerminalBlankParagraph($Document, $Audit) { return [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = $null } }
$result = Repair-WordBlankPages $doc
if (-not $result.Changed -or $script:DeletedStarts.Count -ne 1 -or $script:DeletedStarts[0] -ne 17) { throw "The duplicate page break at the section start was not removed exactly once: changed=$($result.Changed), deleted=$($script:DeletedStarts -join ','), unresolved=$($result.UnresolvedPages | ConvertTo-Json -Compress), audits=$($script:AuditCalls)" }
if ([string]$script:ParagraphRange.Text -cne [string][char]13 -or $section.Range.Start -ne 17 -or $section.PageSetup.SectionStart -ne 2) { throw 'The section paragraph mark or next-page section boundary was not preserved' }
if ($result.Changes[0].cause -ne '下一页分节起始页重复显式分页符' -or $result.UnresolvedPages.Count -ne 0) { throw 'The targeted boundary repair was not reported as resolved' }
if ($script:AuditCalls -ne 2) { throw 'The repair did not refresh the page audit after deleting the form-feed' }
'''
        run_powershell(repair_function + "\r\n" + harness)

    def test_word_page_audit_returns_ascending_pages_and_marks_terminal_blank(self):
        functions = extract_function("Clean-Text") + "\n" + extract_function("Get-WordPageAudit")
        harness = r'''
$wdStatisticPages = 2
$wdActiveEndPageNumber = 3
$wdWithInTable = 12
$script:AuditParagraphs = @()
for ($page = 1; $page -le 22; $page++) {
    $text = if ($page -eq 22) { "`r" } else { "Page $page body`r" }
    $range = [pscustomobject]@{ Page = $page; Text = $text }
    $range | Add-Member ScriptMethod Information {
        param($kind)
        if ($kind -eq 12) { return $false }
        return [int]$this.Page
    }
    $script:AuditParagraphs += [pscustomobject]@{ Range = $range }
}
$paragraphs = [pscustomobject]@{ Count = $script:AuditParagraphs.Count }
$paragraphs | Add-Member ScriptMethod Item {
    param([int]$index)
    return $script:AuditParagraphs[$index - 1]
}
$doc = [pscustomobject]@{
    Paragraphs = $paragraphs
    InlineShapes = @()
    Shapes = @()
    Sections = @()
}
$doc | Add-Member ScriptMethod Repaginate { }
$doc | Add-Member ScriptMethod ComputeStatistics { param($kind) return 22 }
$audit = @(Get-WordPageAudit $doc)
if ($audit.Count -ne 22) { throw "Expected 22 page audit entries; got $($audit.Count)" }
if ([int]$audit[0].page -ne 1 -or [int]$audit[-1].page -ne 22) { throw 'Word page audit is not in ascending page order' }
if ($audit[-1].has_body_content -or @($audit[-1].visible_text).Count -ne 0) { throw 'Page 22 should be recognized as blank' }
if (-not $audit[0].has_body_content -or @($audit[0].visible_text).Count -ne 1 -or $audit[0].visible_text[0] -ne 'Page 1 body') {
    throw 'Page 1 body text was not mapped to the correct audit row'
}
'''
        run_powershell(functions + "\n" + harness)

    def test_terminal_empty_paragraph_compaction_preserves_text_and_blocks_section_boundary(self):
        functions = "\n".join(
            extract_function(name)
            for name in ("Clean-Text", "Repair-TerminalBlankParagraph", "Repair-WordBlankPages")
        )
        harness = r'''
$wdWithInTable = 12
$wdActiveEndPageNumber = 3
$wdActiveEndSectionNumber = 2
$wdStatisticPages = 2
$script:PageCount = 22
$script:TerminalCompacted = $false
$script:Sections = @()
$script:Paragraphs = @()
$script:CompressOnRepaginate = $true
$script:CorruptRestore = $false
$script:RepaginateCount = 0
function Test-RangeHasProtectedContent($Document, $Range, [bool]$CheckParagraphBreakBefore = $false, [bool]$AllowPageBreak = $false, [bool]$AllowSectionBreak = $false) {
    return [bool]$Range.Protected
}
function New-MockParagraph([string]$text, [int]$start, [int]$end, [int]$page, [int]$section = 1, [bool]$protected = $false, [bool]$keep = $false) {
    $range = [pscustomobject]@{
        Text = $text; Start = $start; End = $end; Page = $page; Section = $section; Protected = $protected
        Font = [pscustomobject]@{ Size = [double]12 }
        ParagraphFormat = [pscustomobject]@{
            PageBreakBefore = $false; KeepWithNext = $keep; KeepTogether = $false
            LineSpacingRule = 0; LineSpacing = [double]12; SpaceBefore = [double]0; SpaceAfter = [double]0
        }
    }
    $range | Add-Member ScriptMethod Information {
        param($kind)
        if ($kind -eq 12) { return 0 }
        if ($kind -eq 2) { return [int]$this.Section }
        return [int]$this.Page
    }
    return [pscustomobject]@{ Range = $range }
}
function New-MockDocument([string]$terminalText = "`r", [bool]$sectionBoundary = $false, [bool]$keep = $false) {
    $script:PageCount = 22
    $script:TerminalCompacted = $false
    $script:RepaginateCount = 0
    $previous = New-MockParagraph '11.10 正文内容`r' 100 109 21 1
    $terminalSection = if ($sectionBoundary) { 2 } else { 1 }
    $terminal = New-MockParagraph $terminalText 109 110 22 $terminalSection $false $keep
    $script:Paragraphs = @($previous, $terminal)
    $paragraphs = [pscustomobject]@{ Count = $script:Paragraphs.Count }
    $paragraphs | Add-Member ScriptMethod Item { param([int]$index) return $script:Paragraphs[$index - 1] }
    if ($sectionBoundary) {
        $script:Sections = @(
            [pscustomobject]@{ Range = [pscustomobject]@{ Start = 0; End = 110 } },
            [pscustomobject]@{ Range = [pscustomobject]@{ Start = 110; End = 110 } }
        )
    } else {
        $script:Sections = @([pscustomobject]@{ Range = [pscustomobject]@{ Start = 0; End = 110 } })
    }
    $sections = [pscustomobject]@{
        Count = $script:Sections.Count; Range = $script:Sections[0].Range
        PageSetup = [pscustomobject]@{ SectionStart = 2 }
    }
    $sections | Add-Member ScriptMethod Item { param([int]$index) return $script:Sections[$index - 1] }
    $doc = [pscustomobject]@{ Paragraphs = $paragraphs; Sections = $sections }
    $doc | Add-Member ScriptMethod ComputeStatistics { param($kind) return [int]$script:PageCount }
    $doc | Add-Member ScriptMethod Repaginate {
        $script:RepaginateCount++
        if ($script:CompressOnRepaginate) {
            $script:PageCount = 21
            $script:TerminalCompacted = $true
            $script:Paragraphs[-1].Range.Page = 21
        } elseif ($script:CorruptRestore -and $script:RepaginateCount -ge 2) {
            $script:PageCount = 23
        }
    }
    $doc | Add-Member ScriptMethod Range {
        param($start, $end)
        $range = [pscustomobject]@{ Start = [int]$start; End = [int]$end; Page = 1 }
        $range | Add-Member ScriptMethod Information { param($kind) return [int]$this.Page }
        return $range
    }
    return $doc
}
function Get-WordPageAudit($Document) {
    $rows = @()
    for ($page = 1; $page -le $script:PageCount; $page++) {
        $rows += [pscustomobject]@{ page = $page; has_body_content = ($page -ne $script:PageCount -or $script:TerminalCompacted) }
    }
    return $rows
}
$audit = @(
    [pscustomobject]@{ page = 21; has_body_content = $true },
    [pscustomobject]@{ page = 22; has_body_content = $false }
)
$doc = New-MockDocument
$beforeText = [string]$script:Paragraphs[-1].Range.Text
$fixed = Repair-TerminalBlankParagraph $doc $audit
if (-not $fixed.Changed) { throw "Expected the eligible terminal blank paragraph to move: $($fixed.Diagnostic.reason); pages=$($fixed.Diagnostic.pages_after_attempt); terminal=$($fixed.Diagnostic.terminal_page_after_attempt); textmatch=$($fixed.Diagnostic.terminal_text_matches_after_attempt); codes=$($fixed.Diagnostic.terminal_text_codes_after_attempt -join ',')" }
if ($script:PageCount -ne 21 -or $script:Paragraphs[-1].Range.Page -ne 21) { throw 'Word pagination did not reduce by exactly one page' }
if ([string]$script:Paragraphs[-1].Range.Text -cne $beforeText) { throw 'The terminal paragraph text changed' }
if ($script:Paragraphs[-1].Range.Font.Size -ne 1 -or $script:Paragraphs[-1].Range.ParagraphFormat.LineSpacingRule -ne 4 -or
    $script:Paragraphs[-1].Range.ParagraphFormat.LineSpacing -ne 1 -or $script:Paragraphs[-1].Range.ParagraphFormat.SpaceBefore -ne 0 -or
    $script:Paragraphs[-1].Range.ParagraphFormat.SpaceAfter -ne 0) { throw 'The intended compact format was not applied' }
if ($fixed.Change.cause -ne '末尾空白页终段落格式收缩' -or $fixed.Diagnostic.outcome -ne 'REPAIRED') { throw 'The successful repair was not recorded in diagnostics' }

$reverseAudit = @($audit | Sort-Object -Property page -Descending)
$doc = New-MockDocument
$reverseFixed = Repair-TerminalBlankParagraph $doc $reverseAudit
if (-not $reverseFixed.Changed -or $reverseFixed.Change.page -ne 22) { throw 'A descending page audit failed to repair the maximum page number' }

$doc = New-MockDocument
$integrated = Repair-WordBlankPages $doc
if (-not $integrated.Changed -or $integrated.Changes[-1].cause -ne '末尾空白页终段落格式收缩') { throw 'The blank-page repair report did not record terminal paragraph compaction' }
if ($integrated.UnresolvedPages.Count -ne 0) { throw 'The repaired terminal blank page remained unresolved' }

$script:CompressOnRepaginate = $false
$doc = New-MockDocument
$notMoved = Repair-WordBlankPages $doc
if ($notMoved.Changed -or $notMoved.Diagnostics[-1].outcome -ne 'RESTORED') { throw 'A failed repagination was not marked restored' }
if ($script:PageCount -ne 22 -or $script:Paragraphs[-1].Range.Font.Size -ne 12) { throw 'Original page layout and font size were not restored' }
if ($script:Paragraphs[-1].Range.ParagraphFormat.LineSpacingRule -ne 0 -or $script:Paragraphs[-1].Range.ParagraphFormat.LineSpacing -ne 12) { throw 'Original line spacing was not restored' }
if ($notMoved.UnresolvedPages.Count -ne 1 -or $notMoved.UnresolvedPages[0].page -ne 22) { throw 'A failed terminal-page repair was not left unresolved' }
if ($notMoved.Diagnostics[-1].restore_verified -ne $true -or $notMoved.Diagnostics[-1].restore_failures.Count -ne 0) { throw 'Successful format restoration was not verified' }
if ($script:Paragraphs[-1].Range.Page -ne 22 -or [string]$script:Paragraphs[-1].Range.Text -cne "`r") { throw 'Original terminal page or text was not restored' }
if ($script:Paragraphs[-1].Range.ParagraphFormat.SpaceBefore -ne 0 -or $script:Paragraphs[-1].Range.ParagraphFormat.SpaceAfter -ne 0) { throw 'Original paragraph spacing was not restored' }

$script:CorruptRestore = $true
$doc = New-MockDocument
$incompleteRestoreFailed = $false
try { [void](Repair-TerminalBlankParagraph $doc $audit) } catch { $incompleteRestoreFailed = $_.Exception.Message -match '原状态恢复核验未通过' }
if (-not $incompleteRestoreFailed) { throw 'An incomplete state restoration did not fail closed' }

$doc = New-MockDocument '可见正文`r'
$beforeText = [string]$script:Paragraphs[-1].Range.Text
$blockedText = Repair-TerminalBlankParagraph $doc $audit
if ($blockedText.Changed -or [string]$script:Paragraphs[-1].Range.Text -cne $beforeText) { throw 'A terminal paragraph with visible text was changed' }
if ($blockedText.Diagnostic.reason -ne '末尾段落文本不是单独回车符') { throw 'Visible terminal text was not identified as a safety block' }

$doc = New-MockDocument "`r" $true
$blockedSection = Repair-TerminalBlankParagraph $doc $audit
if ($blockedSection.Changed -or $script:PageCount -ne 22) { throw 'A terminal section boundary was compacted' }
if ($blockedSection.Diagnostic.reason -notmatch '分节') { throw 'The terminal section boundary was not reported' }
'''
        run_powershell(functions + "\n" + harness)


if __name__ == "__main__":
    unittest.main()

import unittest

from test_word_cover_repairs import extract_function, run_powershell


class CoverPerformanceTests(unittest.TestCase):
    def test_leading_table_cover_is_preserved_without_cover_insertion_or_breaks(self):
        functions = "\n".join(
            extract_function(name)
            for name in (
                "Find-CoverBlock",
                "Find-LeadingTableCover",
                "Test-FirstPageCover",
                "Ensure-FirstPageCover",
                "Add-Cover",
                "Ensure-CoverPageBoundary",
                "Audit-EffectiveCover",
            )
        )
        run_powershell(functions + r'''
function Test-CoverTitle($Text, $Title) { return $Text -eq $Title }
function Test-ManualLabel($Text) { return $Text -eq '使用说明书' }
function Test-LegacyManualLabel($Text) { return $Text -in @('使用说明书', '操作说明书') }
function Get-FrontParagraphs($Document, $MaxPage) { return @($script:Items | Where-Object { $_.Page -le $MaxPage }) }
$tableRange = [pscustomobject]@{ Start = 0; End = 40 }
$table = [pscustomobject]@{ Range = $tableRange }
$tables = [pscustomobject]@{ First = $table }
$tables | Add-Member ScriptMethod Item { param($index) if ($index -ne 1) { throw 'Expected leading table only' }; return $this.First }
$doc = [pscustomobject]@{ Tables = $tables; RangeCalls = 0 }
$doc | Add-Member ScriptMethod Range { param($start, $end) $this.RangeCalls++; throw 'A table cover must not receive range edits' }
function New-Item($Text, $Index, $Start, $End, $InTable = $true, $Page = 1, $Style = 'Normal') {
    return [pscustomobject]@{ Text = $Text; Index = $Index; Page = $Page; InTable = $InTable; Style = $Style; Start = $Start; End = $End }
}
$script:Items = @(
    (New-Item '示例软件' 1 0 10),
    (New-Item '操作说明书' 2 10 20),
    (New-Item '表格内原始业务字段：A-001，必须保留。' 3 20 35),
    (New-Item '示例软件' 4 45 55 $false 2)
)
$block = Find-LeadingTableCover $doc $script:Items '示例软件'
if ($null -eq $block -or $block.TitleIndex -ne 0 -or $block.LabelIndex -ne 1) { throw 'Leading table cover not identified' }
if (-not (Test-FirstPageCover $doc '示例软件')) { throw 'ADD_COVER preflight rejected an existing table cover' }
if (-not (Test-FirstPageCover $doc '示例软件')) { Add-Cover $doc '示例软件' }
$ensured = Ensure-FirstPageCover $doc '示例软件'
if ($ensured.Changed -or $ensured.Result -ne 'EXISTING_COVER') { throw 'Ensure-Cover inserted or rewrote a table cover' }
if (Ensure-CoverPageBoundary $doc '示例软件') { throw 'Cover boundary inserted a page break inside a protected table' }
$audit = Audit-EffectiveCover $doc '示例软件'
if (-not $audit.Found -or -not $audit.ProtectedTable -or -not $audit.LabelPresent -or $audit.Label -ne '操作说明书') { throw 'Protected table cover was not recognized by QA' }
if ($doc.RangeCalls -ne 0) { throw 'A protected table range was edited' }

# A table that ends before its manual label is not a cover block. Same-named
# body text on page two cannot be used to authorize an edit to the leading table.
$tableRange.End = 10
if ($null -ne (Find-LeadingTableCover $doc $script:Items '示例软件')) { throw 'Label outside the leading table was accepted' }
''')

    def test_cover_insertion_fails_safely_for_an_unprotected_leading_table(self):
        functions = "\n".join(
            extract_function(name)
            for name in (
                "Find-CoverBlock",
                "Find-LeadingTableCover",
                "Test-FirstPageCover",
                "Ensure-FirstPageCover",
                "Add-Cover",
                "Ensure-CoverPageBoundary",
            )
        )
        run_powershell(functions + r'''
function Test-CoverTitle($Text, $Title) { return $Text -eq $Title }
function Test-ManualLabel($Text) { return $Text -eq '使用说明书' }
function Test-LegacyManualLabel($Text) { return $Text -in @('使用说明书', '操作说明书') }
function Get-FrontParagraphs($Document, $MaxPage) { return @($script:Items | Where-Object { $_.Page -le $MaxPage }) }
$wdActiveEndPageNumber = 3
$wdPageBreak = 7
$script:RangeWrites = 0
$tableRange = [pscustomobject]@{ Start = 0; End = 30 }
$tableRange | Add-Member ScriptMethod Information { param($kind) return 1 }
$table = [pscustomobject]@{ Range = $tableRange }
$tables = New-Object System.Collections.ArrayList
[void]$tables.Add($table)
$tables | Add-Member ScriptMethod Item { param($index) return $this[$index - 1] }
$doc = [pscustomobject]@{ Tables = $tables; InlineShapes = @(); RangeCalls = 0 }
$doc | Add-Member ScriptMethod Range {
    param($start, $end)
    $this.RangeCalls++
    $range = [pscustomobject]@{ }
    $range | Add-Member ScriptMethod InsertBefore { $script:RangeWrites++ }
    $range | Add-Member ScriptMethod InsertBreak { param($kind) $script:RangeWrites++ }
    $range | Add-Member ScriptMethod Delete { $script:RangeWrites++ }
    return $range
}
$script:Items = @(
    [pscustomobject]@{ Text = '业务名称'; Index = 1; Page = 1; InTable = $true; Style = 'Normal'; Start = 0; End = 10 },
    [pscustomobject]@{ Text = '示例软件'; Index = 2; Page = 1; InTable = $true; Style = 'Normal'; Start = 10; End = 20 },
    [pscustomobject]@{ Text = '操作说明书'; Index = 3; Page = 1; InTable = $true; Style = 'Normal'; Start = 20; End = 30 }
)
if (Test-FirstPageCover $doc '示例软件') { throw 'Ordinary leading table was recognized as a cover' }
$addError = ''
try { Add-Cover $doc '示例软件' } catch { $addError = $_.Exception.Message }
if ($addError -notmatch '普通表格开头.*人工处理') { throw "ADD_COVER did not report manual handling: $addError" }
$ensureError = ''
try { [void](Ensure-FirstPageCover $doc '示例软件') } catch { $ensureError = $_.Exception.Message }
if ($ensureError -notmatch '普通表格开头.*人工处理') { throw "Ensure-FirstPageCover did not report manual handling: $ensureError" }
if (Ensure-CoverPageBoundary $doc '示例软件') { throw 'A leading ordinary table received a boundary page break' }
if ($doc.RangeCalls -ne 0 -or $script:RangeWrites -ne 0) { throw 'A leading ordinary table received a Range, Delete, or Insert write' }
''')

    def test_boundary_fallback_ignores_a_table_cell_label(self):
        functions = "\n".join(
            extract_function(name)
            for name in ("Find-CoverBlock", "Find-LeadingTableCover", "Ensure-CoverPageBoundary")
        )
        run_powershell(functions + r'''
$wdActiveEndPageNumber = 3
$wdPageBreak = 7
$script:RangeWrites = 0
function Test-CoverTitle($Text, $Title) { return $Text -eq $Title }
function Test-ManualLabel($Text) { return $Text -eq '使用说明书' }
function Test-LegacyManualLabel($Text) { return $Text -in @('使用说明书', '操作说明书') }
function Get-FrontParagraphs($Document, $MaxPage) { return @($script:Items | Where-Object { $_.Page -le $MaxPage }) }
$tableRange = [pscustomobject]@{ Start = 10; End = 20 }
$tableRange | Add-Member ScriptMethod Information { param($kind) return 1 }
$table = [pscustomobject]@{ Range = $tableRange }
$tables = New-Object System.Collections.ArrayList
[void]$tables.Add($table)
$tables | Add-Member ScriptMethod Item { param($index) return $this[$index - 1] }
$doc = [pscustomobject]@{ Tables = $tables; InlineShapes = @(); RangeCalls = 0 }
$doc | Add-Member ScriptMethod Range {
    param($start, $end)
    $this.RangeCalls++
    $range = [pscustomobject]@{}
    $range | Add-Member ScriptMethod InsertBreak { param($kind) $script:RangeWrites++ }
    $range | Add-Member ScriptMethod Delete { $script:RangeWrites++ }
    return $range
}
$script:Items = @(
    [pscustomobject]@{ Text = '示例软件'; Index = 1; Page = 1; InTable = $false; Style = 'Normal'; Start = 0; End = 10 },
    [pscustomobject]@{ Text = '使用说明书'; Index = 2; Page = 1; InTable = $true; Style = 'Normal'; Start = 10; End = 20 }
)
if (Ensure-CoverPageBoundary $doc '示例软件') { throw 'A label inside a table cell was accepted by the boundary fallback' }
if ($doc.RangeCalls -ne 0 -or $script:RangeWrites -ne 0) { throw 'The boundary fallback wrote into a table cell' }
''')

    def test_version_removal_does_not_write_an_identified_table_cover(self):
        functions = "\n".join(
            extract_function(name)
            for name in ("Find-LeadingTableCover", "Remove-CoverVersion")
        )
        run_powershell(functions + r'''
function Test-CoverTitle($Text, $Title) { return $Text -match '^示例软件(?:\s*V\d+(?:\.\d+)*)?$' }
function Test-LegacyManualLabel($Text) { return $Text -in @('使用说明书', '操作说明书') }
function Get-FrontParagraphs($Document, $MaxPage) { return @($script:Items | Where-Object { $_.Page -le $MaxPage }) }
$tableRange = [pscustomobject]@{ Start = 0; End = 30 }
$table = [pscustomobject]@{ Range = $tableRange }
$tables = [pscustomobject]@{ First = $table }
$tables | Add-Member ScriptMethod Item { param($index) return $this.First }
$doc = [pscustomobject]@{ Tables = $tables; RangeCalls = 0 }
$doc | Add-Member ScriptMethod Range { param($start, $end) $this.RangeCalls++; throw 'A protected cover cell must not be written' }
$script:Items = @(
    [pscustomobject]@{ Text = '示例软件 V1.2'; Index = 1; Page = 1; InTable = $true; Style = 'Normal'; Start = 0; End = 10 },
    [pscustomobject]@{ Text = '操作说明书'; Index = 2; Page = 1; InTable = $true; Style = 'Normal'; Start = 10; End = 20 },
    [pscustomobject]@{ Text = '表格原始字段'; Index = 3; Page = 1; InTable = $true; Style = 'Normal'; Start = 20; End = 30 }
)
$result = Remove-CoverVersion $doc '示例软件'
if (-not $result.Found -or -not $result.ProtectedTable -or $result.Changed) { throw 'Version removal did not preserve the identified table cover' }
if ($doc.RangeCalls -ne 0 -or $script:Items[0].Text -ne '示例软件 V1.2') { throw 'REMOVE_COVER_VERSION edited a protected table cell' }
''')

    def test_section_repair_invalidates_audit(self):
        run_powershell(extract_function("Repair-WordBlankPages") + r'''
$wdActiveEndPageNumber = 3
$script:Audits = 0
$section = [pscustomobject]@{ Range = [pscustomobject]@{ Start = 10 }; PageSetup = [pscustomobject]@{ SectionStart = 4 } }
$doc = [pscustomobject]@{ Sections = @($section); Content = [pscustomobject]@{ End = 50 } }
$doc | Add-Member ScriptMethod Range {
    param($start, $end)
    $range = [pscustomobject]@{}
    $range | Add-Member ScriptMethod Information { param($kind) return 3 }
    return $range
}
$doc | Add-Member ScriptMethod Repaginate { }
function Get-WordPageAudit($Document) {
    $script:Audits++
    return @([pscustomobject]@{ page = 1; has_body_content = $true },
        [pscustomobject]@{ page = 2; has_body_content = ($Document.Sections[0].PageSetup.SectionStart -eq 2) })
}
function Repair-TerminalBlankParagraph($Document, $Audit) { return [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = $null } }
$result = Repair-WordBlankPages $doc
if (-not $result.Changed -or $result.UnresolvedPages.Count -ne 0) { throw 'Section repair failed' }
if ($script:Audits -ne 2 -or $result.AuditCount -ne 2) { throw 'Section repair audit not invalidated' }
''')

    def test_front_paragraph_count_is_read_once(self):
        run_powershell(extract_function("Clean-Text") + "\n" + extract_function("Get-FrontParagraphs") + r'''
$wdActiveEndPageNumber = 3; $wdWithInTable = 12
$script:CountReads = 0
$paragraphs = [pscustomobject]@{}
$paragraphs | Add-Member ScriptProperty Count { $script:CountReads++; return 100 }
$paragraphs | Add-Member ScriptMethod Item {
    param($index)
    $range = [pscustomobject]@{ Text = "段落$index`r"; Start = $index; End = $index + 1; Style = 'Normal'; Page = $index }
    $range | Add-Member ScriptMethod Information { param($kind) if ($kind -eq 12) { return $false }; return $this.Page }
    return [pscustomobject]@{ Range = $range }
}
$items = @(Get-FrontParagraphs ([pscustomobject]@{ Paragraphs = $paragraphs }) 5)
if ($script:CountReads -ne 1) { throw 'Paragraph Count read repeatedly' }
if ($items.Count -ne 5 -or $items[4].Page -ne 5 -or $items[0].Index -ne 1) { throw 'Front paragraph results changed' }
''')

    def test_blank_audit_reused_only_without_mutations(self):
        run_powershell(extract_function("Repair-WordBlankPages") + r'''
$script:Audits = 0
$script:Outcome = ''
function Get-WordPageAudit($Document) {
    $script:Audits++
    return [pscustomobject]@{ page = 1; has_body_content = $true }
}
function Repair-TerminalBlankParagraph($Document, $Audit) {
    return [pscustomobject]@{ Changed = ($script:Outcome -eq 'REPAIRED'); Change = [pscustomobject]@{ page = 1 }; Diagnostic = if ($script:Outcome) { [pscustomobject]@{ outcome = $script:Outcome } } else { $null } }
}
foreach ($outcome in @('', 'RESTORED', 'REPAIRED')) {
    $script:Outcome = $outcome; $script:Audits = 0
    $result = Repair-WordBlankPages ([pscustomobject]@{})
    $expected = if ($outcome) { 2 } else { 1 }
    if ($script:Audits -ne $expected -or $result.AuditCount -ne $expected) { throw "Audit count for '$outcome': expected $expected, got $script:Audits" }
}
''')

    def test_typography_excludes_tables_toc_and_intervening_body(self):
        functions = "\n".join(extract_function(name) for name in ("Find-CoverBlock", "Set-CoverTypography"))
        run_powershell(functions + r'''
function Test-CoverTitle($Text, $Title) { return $Text -eq $Title }
function Test-ManualLabel($Text) { return $Text -eq '使用说明书' }
function Test-LegacyManualLabel($Text) { return (Test-ManualLabel $Text) }
function Get-FrontParagraphs($Document, $MaxPage) { return $script:Items }
$doc = [pscustomobject]@{}
$doc | Add-Member ScriptMethod Range { throw 'Protected material must not be formatted' }
function New-Item($Text, $Index) {
    return [pscustomobject]@{ Text = $Text; Index = $Index; Page = 1; InTable = $false; Style = 'Normal'; Start = $Index; End = $Index + 1 }
}
$script:Items = @((New-Item '示例软件' 1), (New-Item '使用说明书' 2))
if ($null -eq (Find-CoverBlock $script:Items '示例软件' 1 1 $true)) { throw 'Valid cover rejected' }
foreach ($index in 0..0) {
    foreach ($kind in @('Table', 'TOC')) {
        $script:Items = @((New-Item '示例软件' 1), (New-Item '使用说明书' 2))
        if ($kind -eq 'Table') { $script:Items[$index].InTable = $true } else { $script:Items[$index].Style = 'TOC1' }
        if ($null -ne (Find-CoverBlock $script:Items '示例软件' 1 1 $true)) { throw "Protected $kind accepted" }
        if ((Set-CoverTypography $doc '示例软件').Found) { throw "Protected $kind formatted" }
    }
}
$script:Items = @((New-Item '示例软件' 1), (New-Item '这里是正文内容。' 2), (New-Item '使用说明书' 3))
if ($null -ne (Find-CoverBlock $script:Items '示例软件' 1 1 $true)) { throw 'Label beyond body accepted as a cover block' }
''')

    def test_typography_styles_label_only_after_safe_page_one_reflow(self):
        functions = "\n".join(extract_function(name) for name in ("Find-CoverBlock", "Set-CoverTypography"))
        run_powershell(functions + r'''
$wdActiveEndSectionNumber = 8; $wdVerticalCenter = 1; $wdNumberAllNumbers = 0
$wdStyleNormal = 0; $wdOutlineLevelBodyText = 10; $wdAlignParagraphCenter = 1; $wdStatisticLines = 1
function Test-CoverTitle($Text, $Title) { return $Text -eq $Title }
function Test-ManualLabel($Text) { return $Text -eq '使用说明书' }
function Test-LegacyManualLabel($Text) { return (Test-ManualLabel $Text) }
function Get-FrontParagraphs($Document, $MaxPage) { return @($script:Items | Where-Object { $_.Page -le $MaxPage }) }
function New-MockRange($Start, $End) {
    $paragraphFormat = [pscustomobject]@{}
    $font = [pscustomobject]@{}
    $listFormat = [pscustomobject]@{}
    $listFormat | Add-Member ScriptMethod RemoveNumbers { param($kind) }
    $paragraphRange = [pscustomobject]@{ ListFormat = $listFormat; ParagraphFormat = [pscustomobject]@{}; Style = 'Normal' }
    $paragraph = [pscustomobject]@{ Range = $paragraphRange }
    $range = [pscustomobject]@{ Start = $Start; End = $End; Paragraphs = @($paragraph); ParagraphFormat = $paragraphFormat; Font = $font }
    $range | Add-Member ScriptMethod Information { param($kind) return 1 }
    $range | Add-Member ScriptMethod ComputeStatistics { param($kind) return 1 }
    return $range
}
function New-MockDocument([int]$TargetLabelPage, [bool]$BodyBetween = $false, [bool]$FormattedLabelPushesToPageTwo = $false, [int]$InitialLabelPage = 2) {
    $script:TargetLabelPage = $TargetLabelPage
    $script:FormattedLabelPushesToPageTwo = $FormattedLabelPushesToPageTwo
    $script:Repaginations = 0
    $script:CreatedRanges = @()
    $script:Items = @([pscustomobject]@{ Text = '示例软件'; Index = 1; Page = 1; InTable = $false; Style = 'Normal'; Start = 10; End = 11 })
    if ($BodyBetween) {
        $script:Items += [pscustomobject]@{ Text = '这里是正文'; Index = 2; Page = 1; InTable = $false; Style = 'Normal'; Start = 15; End = 16 }
    }
    $labelIndex = $script:Items.Count + 1
    $script:Items += [pscustomobject]@{ Text = '使用说明书'; Index = $labelIndex; Page = $InitialLabelPage; InTable = $false; Style = 'Normal'; Start = 20; End = 21 }
    $section = [pscustomobject]@{ PageSetup = [pscustomobject]@{} }
    $sections = [pscustomobject]@{}
    $sections | Add-Member ScriptMethod Item { param($index) return $section }
    $styles = [pscustomobject]@{}
    $styles | Add-Member ScriptMethod Item { param($index) return 'Normal' }
    $doc = [pscustomobject]@{ Sections = $sections; Styles = $styles; RepaginateCount = 0 }
    $doc | Add-Member ScriptMethod Range {
        param($start, $end)
        $script:CreatedRanges += [int]$start
        return New-MockRange $start $end
    }
    $doc | Add-Member ScriptMethod Repaginate {
        $this.RepaginateCount++
        $script:Repaginations++
        foreach ($item in $script:Items) {
            if ($item.Text -eq '使用说明书') {
                $item.Page = if ($script:FormattedLabelPushesToPageTwo -and 20 -in $script:CreatedRanges) { 2 } else { $script:TargetLabelPage }
            }
        }
    }
    return $doc
}
function Get-CoverLastLineUnitCount($Range, $Workload = $null) { return 3 }

$doc = New-MockDocument 2
$result = Set-CoverTypography $doc '示例软件'
if (-not $result.Found -or $result.LabelPresent -or $result.Label -ne '') { throw 'A label still on page two must not fail title formatting or be reported as formatted' }
if ($script:CreatedRanges.Count -ne 1 -or $script:CreatedRanges[0] -ne 10) { throw 'A label still on page two must not receive a formatting range' }

$doc = New-MockDocument 1
$result = Set-CoverTypography $doc '示例软件'
if (-not $result.Found -or -not $result.LabelPresent -or $result.Label -ne '使用说明书') { throw 'A label reflowed to page one was not recognized' }
if ($script:CreatedRanges.Count -ne 2 -or 20 -notin $script:CreatedRanges -or $doc.RepaginateCount -ne 2) { throw 'A safe reflowed label was not formatted and repaginated after pagination' }

$doc = New-MockDocument 1 $false $false 1
$result = Set-CoverTypography $doc '示例软件'
if (-not $result.Found -or -not $result.LabelPresent -or $result.LabelLayoutFailed) { throw 'An initially safe same-page cover unexpectedly failed' }
if ($doc.RepaginateCount -ne 1 -or $result.Workload.pagination_attempts -ne 1) { throw 'An initially safe same-page cover should repaginate once per size attempt' }

$doc = New-MockDocument 1 $false $true 1
$result = Set-CoverTypography $doc '示例软件'
if (-not $result.Found -or -not $result.LabelLayoutFailed -or $result.LabelPresent) { throw 'A label pushed to page two by formatting must be reported as a layout failure' }
if ($doc.RepaginateCount -ne 13 -or $result.Workload.pagination_attempts -ne 13) { throw 'A failed same-page convergence should use one pagination per size attempt' }

$doc = New-MockDocument 1 $true
$result = Set-CoverTypography $doc '示例软件'
if (-not $result.Found -or $result.LabelPresent) { throw 'A body paragraph between title and label must block label formatting, not title formatting' }
if (20 -in $script:CreatedRanges) { throw 'The label beyond intervening body text was formatted' }

foreach ($kind in @('Table', 'TOC')) {
    $doc = New-MockDocument 1
    if ($kind -eq 'Table') { $script:Items[1].InTable = $true } else { $script:Items[1].Style = 'TOC1' }
    $result = Set-CoverTypography $doc '示例软件'
    if (-not $result.Found -or $result.LabelPresent) { throw "Protected $kind label must not prevent safe title formatting or be reported as formatted" }
    if (20 -in $script:CreatedRanges) { throw "Protected $kind label was formatted" }
}
''')

    def test_last_line_equivalence_and_workload(self):
        functions = extract_function("Clean-Text") + "\n" + extract_function("Get-CoverLastLineUnitCount")
        run_powershell(functions + r'''
$wdFirstCharacterLineNumber = 10; $wdActiveEndPageNumber = 3
function Legacy-Units($Range) {
    $lastLine = -1; $lastLineText = ''
    foreach ($character in @($Range.Characters)) {
        $text = Clean-Text ([string]$character.Text)
        if ($text -eq '') { continue }
        try { $line = [int]$character.Information(10) } catch { continue }
        if ($line -gt $lastLine) { $lastLine = $line; $lastLineText = $text } elseif ($line -eq $lastLine) { $lastLineText += $text }
    }
    return @([regex]::Matches($lastLineText, '[\p{IsCJKUnifiedIdeographs}]|[A-Za-z0-9]+')).Count
}
function New-Range($Text, $Lines, $Pages, $FailureIndex = -1) {
    $chars = @()
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $c = [pscustomobject]@{ Text = [string]$Text[$i]; Line = $Lines[$i]; Page = $Pages[$i]; Fail = ($i -eq $FailureIndex) }
        $c | Add-Member ScriptMethod Information {
            param($kind)
            if ($kind -eq 3) { return $this.Page }
            $script:LineCalls++
            if ($this.Fail) { throw 'Unavailable' }
            return $this.Line
        }
        $chars += $c
    }
    return [pscustomobject]@{ Characters = $chars }
}
$cases = @(
    [pscustomobject]@{ Text = "甲乙ABC xyz`r"; Lines = @(1,1,2,2,2,2,2,2,2,2); Pages = @(1)*10; Fail = -1 },
    [pscustomobject]@{ Text = "甲`n 乙`r"; Lines = @(1,1,2,2,2); Pages = @(1)*5; Fail = -1 },
    [pscustomobject]@{ Text = '甲乙丙丁'; Lines = @(5,5,1,1); Pages = @(1,1,2,2); Fail = -1 },
    [pscustomobject]@{ Text = '甲乙丙丁'; Lines = @(1,1,2,2); Pages = @(1)*4; Fail = 2 },
    [pscustomobject]@{ Text = '甲乙丙丁'; Lines = @(1,1,0,0); Pages = @(1)*4; Fail = -1 }
)
foreach ($case in $cases) {
    $range = New-Range $case.Text $case.Lines $case.Pages $case.Fail
    $expected = Legacy-Units $range; $script:LineCalls = 0
    $actual = Get-CoverLastLineUnitCount $range
    if ($actual -ne $expected) { throw "Units changed: $actual vs $expected" }
}
$range = New-Range (('甲' * 100) + 'ABC') ((@(1)*100) + (@(2)*3)) (@(1)*103)
$collection = [pscustomobject]@{ Count = $range.Characters.Count; Items = $range.Characters }
$collection | Add-Member ScriptMethod Item { param($index) return $this.Items[$index - 1] }
$range.Characters = $collection
$script:LineCalls = 0
$workload = [ordered]@{ line_information_calls = 0; page_information_calls = 0; full_scan_fallbacks = 0 }
if ((Get-CoverLastLineUnitCount $range $workload) -ne 1) { throw 'Latin token split' }
if ($script:LineCalls -gt 5) { throw "Too many line calls: $script:LineCalls" }
if ($workload.line_information_calls -ne $script:LineCalls -or $workload.full_scan_fallbacks -ne 0) { throw 'Workload diagnostic mismatch' }
''')

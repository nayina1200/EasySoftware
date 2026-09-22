param(
    [Parameter(Mandatory = $true)]
    [string]$BaseDir,
    [switch]$AuditOnly,
    [string[]]$DuplicateTitles = @(),
    [string]$ReportPath = ""
)

$ErrorActionPreference = "Stop"

$wdActiveEndPageNumber = 3
$wdStatisticPages = 2
$wdStatisticLines = 1
$wdExportFormatPDF = 17
$wdAlignParagraphCenter = 1

function Clean-Text([string]$Text) {
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[\r\n\a\f\v]", "") -replace "\s+", " ").Trim()
}

function Test-ManualLabel([string]$Text) {
    $value = Clean-Text $Text
    return $value -match '^[【\[]?(使用说明书|使用说明|说明书|使用手册|用户手册|用户使用手册|用户操作手册|用户操作说明|操作说明书|操作说明|操作手册)[】\]]?$'
}

function Get-FrontParagraphs($Document, [int]$MaxPage = 4) {
    $items = @()
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $range = $paragraph.Range
        $page = [int]$range.Information($wdActiveEndPageNumber)
        if ($page -gt $MaxPage) { break }
        $items += [pscustomobject]@{
            Index = $i
            Page = $page
            Start = [int]$range.Start
            End = [int]$range.End
            Text = Clean-Text $range.Text
        }
    }
    return $items
}

function Find-CoverBlock($Items, [string]$Title, [int]$MinimumPage) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Text -ne $Title) { continue }
        $limit = [Math]::Min($Items.Count - 1, $i + 10)
        for ($j = $i + 1; $j -le $limit; $j++) {
            if ($Items[$j].Page -ne $item.Page) { break }
            if (Test-ManualLabel $Items[$j].Text) {
                return [pscustomobject]@{ TitleIndex = $i; LabelIndex = $j; Page = $item.Page }
            }
        }
    }
    return $null
}

function Remove-DuplicateCoverBlocks($Document, [string]$Title, [switch]$Preview) {
    $removed = 0
    for ($round = 0; $round -lt 8; $round++) {
        $items = @(Get-FrontParagraphs $Document 4)
        $block = Find-CoverBlock $items $Title 2
        if ($null -eq $block) { break }
        $removed++
        if ($Preview) { break }

        $startIndex = $block.TitleIndex
        while ($startIndex -gt 0 -and
               $items[$startIndex - 1].Page -eq $block.Page -and
               $items[$startIndex - 1].Text -eq "") {
            $startIndex--
        }

        $endIndex = $block.LabelIndex
        $blankCount = 0
        while ($endIndex + 1 -lt $items.Count -and
               $items[$endIndex + 1].Page -eq $block.Page -and
               $items[$endIndex + 1].Text -eq "" -and
               $blankCount -lt 3) {
            $endIndex++
            $blankCount++
        }

        $deleteRange = $Document.Range($items[$startIndex].Start, $items[$endIndex].End)
        $deleteRange.Delete() | Out-Null

        # Delete only the duplicate title block. Never delete the whole page:
        # the same page may continue with a TOC or body content below the old cover text.
    }
    return $removed
}

function Set-CoverTypography($Document, [string]$Title, [switch]$Preview) {
    $items = @(Get-FrontParagraphs $Document 1)
    $titleItem = $items | Where-Object { $_.Text -eq $Title } | Select-Object -First 1
    $labelItem = $items | Where-Object { Test-ManualLabel $_.Text } | Select-Object -First 1
    if ($null -eq $titleItem -or $null -eq $labelItem) {
        return [pscustomobject]@{ Found = $false; Size = 0; Lines = 0; Label = "" }
    }
    if ($Preview) {
        $titleRange = $Document.Range($titleItem.Start, $titleItem.End)
        return [pscustomobject]@{
            Found = $true
            Size = [double]$titleRange.Font.Size
            Lines = [int]$titleRange.ComputeStatistics($wdStatisticLines)
            Label = $labelItem.Text
        }
    }

    $titleRange = $Document.Range($titleItem.Start, $titleItem.End)
    $labelRange = $Document.Range($labelItem.Start, $labelItem.End)
    # Remove inherited list numbering from the complete cover title stack,
    # including the blank spacer paragraphs between title and label.
    $firstIndex = [Math]::Min($titleItem.Index, $labelItem.Index)
    $lastIndex = [Math]::Max($titleItem.Index, $labelItem.Index)
    for ($paragraphIndex = $firstIndex; $paragraphIndex -le $lastIndex; $paragraphIndex++) {
        $coverParagraphRange = $Document.Paragraphs.Item($paragraphIndex).Range
        # A numbered heading style can re-create visible numbers even after
        # ListFormat.RemoveNumbers(). Reset cover paragraphs to Word's built-in
        # Normal style first, then apply the required direct formatting below.
        try { $coverParagraphRange.Style = -1 } catch {}
        try { $coverParagraphRange.ListFormat.RemoveNumbers() } catch {}
        $coverParagraphRange.ParagraphFormat.LeftIndent = 0
        $coverParagraphRange.ParagraphFormat.RightIndent = 0
        $coverParagraphRange.ParagraphFormat.FirstLineIndent = 0
    }
    foreach ($range in @($titleRange, $labelRange)) {
        try { $range.ListFormat.RemoveNumbers() } catch {}
        $range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
        $range.ParagraphFormat.LeftIndent = 0
        $range.ParagraphFormat.RightIndent = 0
        $range.ParagraphFormat.FirstLineIndent = 0
        $range.Font.Name = "宋体"
        $range.Font.NameFarEast = "宋体"
        $range.Font.NameAscii = "SimSun"
        $range.Font.NameOther = "SimSun"
        $range.Font.Bold = -1
    }

    $size = 26
    $lines = 99
    while ($size -ge 14) {
        $titleRange.Font.Size = $size
        $labelRange.Font.Size = $size
        $Document.Repaginate()
        $lines = [int]$titleRange.ComputeStatistics($wdStatisticLines)
        if ($lines -le 1) { break }
        $size--
    }
    return [pscustomobject]@{ Found = $true; Size = $size; Lines = $lines; Label = $labelItem.Text }
}

$base = (Resolve-Path -LiteralPath $BaseDir).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$results = @()

try {
    $folders = Get-ChildItem -LiteralPath $base -Directory | Sort-Object Name
    foreach ($folder in $folders) {
        $docx = Get-ChildItem -LiteralPath $folder.FullName -File -Filter "*说明.docx" | Select-Object -First 1
        if ($null -eq $docx) { continue }
        $txt = Get-ChildItem -LiteralPath $folder.FullName -File -Filter "*.txt" | Select-Object -First 1
        $title = if ($null -ne $txt) { $txt.BaseName } else { $folder.Name }
        $pdf = [IO.Path]::ChangeExtension($docx.FullName, ".pdf")
        $doc = $null
        try {
            $doc = $word.Documents.Open($docx.FullName, $false, $AuditOnly.IsPresent)
            $duplicates = if ($AuditOnly) {
                Remove-DuplicateCoverBlocks $doc $title -Preview
            } elseif ($DuplicateTitles -contains $title) {
                Remove-DuplicateCoverBlocks $doc $title
            } else { 0 }
            $style = Set-CoverTypography $doc $title -Preview:$AuditOnly
            if (-not $AuditOnly) {
                $doc.Fields.Update() | Out-Null
                foreach ($section in $doc.Sections) {
                    foreach ($header in $section.Headers) {
                        if ($header.Exists) { $header.Range.Fields.Update() | Out-Null }
                    }
                }
                $doc.Save()
                $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF)
            }
            $results += [pscustomobject]@{
                Name = $title
                DuplicateCoverBlocks = $duplicates
                CoverFound = $style.Found
                Font = if ($style.Found) { "宋体" } else { "" }
                Bold = [bool]$style.Found
                FontSize = $style.Size
                TitleLines = $style.Lines
                Label = $style.Label
                Pdf = $pdf
                Status = "OK"
            }
        } catch {
            $results += [pscustomobject]@{
                Name = $title
                DuplicateCoverBlocks = -1
                CoverFound = $false
                Font = ""
                Bold = $false
                FontSize = 0
                TitleLines = 0
                Label = ""
                Pdf = $pdf
                Status = $_.Exception.Message
            }
        } finally {
            if ($null -ne $doc) {
                $doc.Close($false)
                [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
        }
    }
} finally {
    $word.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($ReportPath) {
    $results | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
}
$results | Format-Table Name, DuplicateCoverBlocks, CoverFound, FontSize, TitleLines, Label, Status -AutoSize

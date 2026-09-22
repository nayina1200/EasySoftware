param(
    [Parameter(Mandatory = $true)][string]$BaseDir,
    [Parameter(Mandatory = $true)][string]$PlanPath,
    [Parameter(Mandatory = $true)][string]$ReportPath,
    [string]$PidPath = "",
    [string]$ProgressPath = "",
    [int]$ProgressTotal = 0,
    [string]$ProgressPhase = "说明书处理"
)

$ErrorActionPreference = "Stop"
$wdActiveEndPageNumber = 3
$wdActiveEndSectionNumber = 2
$wdStatisticPages = 2
$wdStatisticLines = 1
$wdExportFormatPDF = 17
$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdAlignTabRight = 2
$wdTabLeaderSpaces = 0
$wdVerticalCenter = 1
$wdVerticalTop = 0
$wdHeaderPrimary = 1
$wdHeaderFirstPage = 2
$wdHeaderEvenPages = 3
$wdFieldPage = 33
$wdCollapseEnd = 0
$wdPageBreak = 7
$wdSectionBreakNextPage = 2
$wdVerticalPositionRelativeToPage = 6
$wdOutlineLevelBodyText = 10
$wdWithInTable = 12
$wdFindContinue = 1
$wdReplaceAll = 2
$wdNumberAllNumbers = 3
$wdStyleNormal = -1
$ProgressStartedAt = [DateTime]::UtcNow

function Clean-Text([string]$Text) {
    if ($null -eq $Text) { return "" }
    # Word exposes inline-shape anchors and page/section controls as non-printing
    # characters in Range.Text. They are not visible cover content.
    return (($Text -replace "[\x00-\x1F\x7F\uFFFC]", "") -replace "\s+", " ").Trim()
}

function Compact-Text([string]$Text) {
    return ((Clean-Text $Text) -replace '\s+', '')
}

function Strip-OrdinalPrefix([string]$Text) {
    return (Compact-Text $Text) -replace '^(?:\d+[、.．)）]|[（(](?:\d+|[一二三四五六七八九十]+)[）)]|第[一二三四五六七八九十]+[章节部分、，.]?|[一二三四五六七八九十]+[、.．)）])', ''
}

function Test-CoverTitle([string]$Text, [string]$Title) {
    $value = Strip-OrdinalPrefix $Text
    # A supplied cover often puts the version on the title line (“软件名 V1.0”).
    # Treat it as the cover title so it can be detected and repaired to the pure
    # software name; the version only belongs in the running header.
    $value = $value -replace '(?i)(?:\s*(?:V|版本)\s*\d+(?:\.\d+)*|\s*\d+\.\d+(?:\.\d+)*)\s*$', ''
    return $value -eq (Compact-Text $Title)
}

function Test-SectionHeading($Paragraph, [string]$Text) {
    try {
        $style = [string]$Paragraph.Range.Style
        if ($style -match '(?i)heading|标题|章|节|目录') { return $true }
        if ([int]$Paragraph.OutlineLevel -lt $wdOutlineLevelBodyText) { return $true }
    } catch {}
    $compact = Compact-Text $Text
    return ($compact.Length -le 80 -and $compact -match '^(?:第[一二三四五六七八九十]+[章节部分]|[一二三四五六七八九十]+[、.]|\d+[、.])')
}

function Test-ManualLabel([string]$Text) {
    $value = Compact-Text $Text
    return $value -match '^[【\[]?(使用说明书|使用手册|用户手册|用户使用手册|用户操作手册|用户操作说明|操作说明书|操作手册|软件设计说明书|软件设计说明)[】\]]?$'
}

function Test-LegacyManualLabel([string]$Text) {
    $value = Compact-Text $Text
    # Older supplied covers commonly prefix the same label with “软件”.  Treat
    # that as the same cover marker so a newly inserted normalized cover does
    # not leave a second legacy title block behind.
    $value = $value -replace '^[0-9一二三四五六七八九十]+[、\.)）]', ''
    $value = $value -replace '[《》「」『』【】\[\]（）\(\)"“”]', ''
    if ((Test-ManualLabel $value) -or $value -match '^[【\[]?(软件)?(使用说明书|使用手册|用户手册|用户使用手册|用户操作手册|用户操作说明|操作说明书|操作手册|软件说明书|软件设计说明书|软件设计说明|说明书)[】\]]?$') { return $true }
    $signals = @('书', '说明', '软件', '设计', '使用') | Where-Object { $value.Contains($_) }
    return @($signals).Count -ge 2 -and (@($signals) -contains '书' -or @($signals) -contains '说明' -or @($signals) -contains '使用')
}

function Test-Action($Actions, [string]$Name) {
    return @($Actions) -contains $Name
}

function Test-ImageCaption($Paragraph, [string]$Text) {
    if ($Text -match '^\s*(?:图|表|截图)\s*(?:[0-9０-９]+|[一二三四五六七八九十]+)') { return $true }
    try {
        $styleName = [string]$Paragraph.Range.Style
        return $styleName -match '(?i)caption|题注|图注'
    } catch { return $false }
}

function Write-WorkflowProgress([int]$Current, [string]$Title, [bool]$CurrentProjectActive = $false) {
    if ([string]::IsNullOrWhiteSpace($ProgressPath)) { return }
    $total = [Math]::Max($ProgressTotal, $Current)
    $elapsed = ([DateTime]::UtcNow - $ProgressStartedAt).TotalSeconds
    $estimate = if ($Current -gt 0) { ($elapsed / $Current) * [Math]::Max(0, $total - $Current) } else { $null }
    $state = [ordered]@{
        version = "3.4.26"; status = "RUNNING"; phase = $ProgressPhase
        current = $Current; total = $total; project_name = $Title
        elapsed_seconds = [Math]::Round($elapsed, 1)
        completed_projects = $Current; remaining_projects = [Math]::Max(0, $total - $Current)
        estimated_remaining_seconds = if ($null -ne $estimate) { [Math]::Round($estimate, 1) } else { $null }
        current_module = $ProgressPhase
        current_project_active = $CurrentProjectActive
        updated_at = [DateTime]::Now.ToString("o")
    }
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($ProgressPath))) | Out-Null
    $state | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $ProgressPath -Encoding UTF8
}

function Get-BodyTextParagraphs($Document, [string]$Title) {
    $items = @()
    $coverLastPage = 0
    $coverTitle = Compact-Text $Title
    for ($coverIndex = 1; $coverIndex -le $Document.Paragraphs.Count; $coverIndex++) {
        try {
            $coverParagraph = $Document.Paragraphs.Item($coverIndex)
            $coverText = Clean-Text $coverParagraph.Range.Text
            if ((Strip-OrdinalPrefix $coverText) -eq $coverTitle) {
                $coverLastPage = [Math]::Max($coverLastPage, [int]$coverParagraph.Range.Information($wdActiveEndPageNumber))
            } elseif (Test-LegacyManualLabel $coverText) {
                $coverLastPage = [Math]::Max($coverLastPage, [int]$coverParagraph.Range.Information($wdActiveEndPageNumber))
            }
        } catch {}
    }
    $firstPostCoverCandidate = $true
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        try {
        $paragraph = $Document.Paragraphs.Item($i)
        $text = Clean-Text $paragraph.Range.Text
        $isCaption = Test-ImageCaption $paragraph $text
        if ($text.Length -lt 2 -and -not $isCaption) { continue }
        if ((Strip-OrdinalPrefix $text) -eq $coverTitle) { continue }
        if (Test-LegacyManualLabel $text) { continue }
        if (Test-SectionHeading $paragraph $text) {
            try {
                $headingPage = [int]$paragraph.Range.Information($wdActiveEndPageNumber)
                if ($coverLastPage -gt 0 -and $headingPage -eq ($coverLastPage + 1)) { $firstPostCoverCandidate = $false }
            } catch {}
            continue
        }
        if ($coverLastPage -gt 0 -and $firstPostCoverCandidate) {
            try {
                $page = [int]$paragraph.Range.Information($wdActiveEndPageNumber)
                if ($page -eq ($coverLastPage + 1) -and $text.Length -le 80) { continue }
            } catch {}
            $firstPostCoverCandidate = $false
        }
        if ($text -match '^\s*(?:\d+[、.．)）]|[（(]\s*(?:\d+|[一二三四五六七八九十]+)\s*[）)]|第[一二三四五六七八九十]+[章节部分、，.]?|[一二三四五六七八九十]+[、.．)）]|[•●○◆▪])') { continue }
        try { if ($paragraph.Range.Information($wdWithInTable)) { continue } } catch {}
        # Headings and cover content are not ordinary body text. Word marks
        # ordinary body paragraphs with outline level 10.  Some supplied files
        # put short ordinary text (for example "登录") in an automatic list;
        # its visible content is still body text, so list membership alone must
        # not exclude it from the font repair.
        try {
            if ([string]$paragraph.OutlineLevel -ne [string]$wdOutlineLevelBodyText -and -not $isCaption) { continue }
        } catch { if (-not $isCaption) { continue } }
        $font = $paragraph.Range.Font
        $name = [string]$font.NameFarEast
        if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$font.Name }
        try {
            $size = [double]$font.Size
            $bold = $font.Bold
            $italic = $font.Italic
        } catch { continue }
        $mixed = [string]::IsNullOrWhiteSpace($name) -or $size -le 0 -or $size -ge 999999 -or [string]$bold -eq "9999999" -or [string]$italic -eq "9999999"
        $items += [pscustomobject]@{
            Paragraph = $paragraph
            FontName = $name
            FontNameFarEast = [string]$font.NameFarEast
            Size = $size
            Bold = ([string]$bold -eq "-1")
            Italic = ([string]$italic -eq "-1")
            Mixed = $mixed
            Signature = if ($mixed) { "__MIXED__$i" } else { ($name + "|" + ([Math]::Round($size, 2)) + "|" + ([string]$bold -eq "-1") + "|" + ([string]$italic -eq "-1")) }
        }
        } catch { continue }
    }
    return @($items)
}

function Normalize-BodyTextStyles($Document, [string]$Title) {
    $items = @(Get-BodyTextParagraphs $Document $Title)
    $referenceItems = @($items | Where-Object { -not $_.Mixed })
    if ($referenceItems.Count -lt 2) {
        return [pscustomobject]@{ Found = $false; Changed = $false; ChangedCount = 0; ReferenceFont = ""; ReferenceSize = 0 }
    }
    $referenceSignature = ($referenceItems | Group-Object Signature | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name | Select-Object -First 1).Name
    $reference = $referenceItems | Where-Object { $_.Signature -eq $referenceSignature } | Select-Object -First 1
    $changedCount = 0
    foreach ($item in $items | Where-Object { $_.Mixed -or $_.Signature -ne $referenceSignature }) {
        try {
            $target = $item.Paragraph.Range.Font
            $target.Name = $reference.FontName
            if (-not [string]::IsNullOrWhiteSpace($reference.FontNameFarEast)) { $target.NameFarEast = $reference.FontNameFarEast }
            $target.Size = $reference.Size
            $target.Bold = if ($reference.Bold) { -1 } else { 0 }
            $target.Italic = if ($reference.Italic) { -1 } else { 0 }
            $changedCount++
        } catch { continue }
    }
    return [pscustomobject]@{ Found = $true; Changed = ($changedCount -gt 0); ChangedCount = $changedCount; ReferenceFont = $reference.FontName; ReferenceSize = $reference.Size }
}

function Normalize-FirstPostCoverHeading($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 2)
    $cover = Find-CoverBlock $items $Title 1 1 $true
    if ($null -eq $cover) { return [pscustomobject]@{ Found = $false; Changed = $false; Text = "" } }
    $candidate = $null
    for ($index = $cover.LabelIndex + 1; $index -lt $items.Count; $index++) {
        $item = $items[$index]
        if ($item.Page -le $cover.Page -or $item.Text -eq "") { continue }
        if ((Test-CoverTitle $item.Text $Title) -or (Test-LegacyManualLabel $item.Text)) { continue }
        if ($item.Text.Length -gt 80 -or $item.Text -match '[。！？；;!?]') { break }
        $candidate = $item
        break
    }
    if ($null -eq $candidate) { return [pscustomobject]@{ Found = $false; Changed = $false; Text = "" } }

    $candidateParagraph = $Document.Paragraphs.Item($candidate.Index)
    $references = @()
    for ($index = $candidate.Index + 1; $index -le $Document.Paragraphs.Count; $index++) {
        $paragraph = $Document.Paragraphs.Item($index)
        $text = Clean-Text $paragraph.Range.Text
        if ($text.Length -lt 2 -or $text.Length -gt 80 -or $text -match '[。！？；:：;!?]') { continue }
        if (-not (Test-SectionHeading $paragraph $text)) {
            try { if ([string]$paragraph.Range.Font.Bold -ne "-1") { continue } } catch { continue }
        }
        try {
            $font = $paragraph.Range.Font
            $name = [string]$font.NameFarEast
            if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$font.Name }
            $size = [double]$font.Size
            $bold = [string]$font.Bold -eq "-1"
            $italic = [string]$font.Italic -eq "-1"
            if ([string]::IsNullOrWhiteSpace($name) -or $size -le 0 -or $size -ge 999999 -or [string]$font.Bold -eq "9999999") { continue }
            $references += [pscustomobject]@{ Paragraph = $paragraph; Name = $name; FarEast = [string]$font.NameFarEast; Size = $size; Bold = $bold; Italic = $italic; Signature = "$name|$size|$bold|$italic" }
        } catch { continue }
    }
    if ($references.Count -lt 2) { return [pscustomobject]@{ Found = $true; Changed = $false; Text = $candidate.Text } }
    $referenceSignature = ($references | Group-Object Signature | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name | Select-Object -First 1).Name
    $reference = $references | Where-Object { $_.Signature -eq $referenceSignature } | Select-Object -First 1
    try {
        $target = $candidateParagraph.Range.Font
        $same = ([string]$target.NameFarEast -eq $reference.FarEast -or ([string]$target.Name -eq $reference.Name -and [string]::IsNullOrWhiteSpace($reference.FarEast))) -and
            ([double]$target.Size -eq $reference.Size) -and ([string]$target.Bold -eq $(if ($reference.Bold) { "-1" } else { "0" })) -and ([string]$target.Italic -eq $(if ($reference.Italic) { "-1" } else { "0" }))
        if ($same) { return [pscustomobject]@{ Found = $true; Changed = $false; Text = $candidate.Text } }
        $target.Name = $reference.Name
        if (-not [string]::IsNullOrWhiteSpace($reference.FarEast)) { $target.NameFarEast = $reference.FarEast }
        $target.Size = $reference.Size
        $target.Bold = if ($reference.Bold) { -1 } else { 0 }
        $target.Italic = if ($reference.Italic) { -1 } else { 0 }
        return [pscustomobject]@{ Found = $true; Changed = $true; Text = $candidate.Text }
    } catch { return [pscustomobject]@{ Found = $true; Changed = $false; Text = $candidate.Text } }
}

function Get-TitleDoubleQuoteCount($Document, [string]$Title) {
    $total = 0
    $needles = @(('"' + $Title + '"'), ('“' + $Title + '”'))
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            $text = [string]$range.Text
            foreach ($needle in $needles) {
                $total += [regex]::Matches($text, [regex]::Escape($needle)).Count
            }
            $range = $range.NextStoryRange
        }
    }
    return $total
}

function Remove-TitleDoubleQuotes($Document, [string]$Title) {
    $removed = Get-TitleDoubleQuoteCount $Document $Title
    if ($removed -eq 0) { return 0 }
    $needles = @(('"' + $Title + '"'), ('“' + $Title + '”'))
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            foreach ($needle in $needles) {
                $find = $range.Find
                $find.ClearFormatting()
                $find.Replacement.ClearFormatting()
                $find.Text = $needle
                $find.Replacement.Text = $Title
                $find.Forward = $true
                $find.Wrap = $wdFindContinue
                $find.Format = $false
                [void]$find.Execute($needle, $false, $false, $false, $false, $false, $true, $wdFindContinue, $false, $Title, $wdReplaceAll, $false, $false, $false, $false)
            }
            $range = $range.NextStoryRange
        }
    }
    return $removed
}

function Replace-AllManualQuotes($Document) {
    $removed = 0
    $needles = @('"', '“', '”', '「', '」', '『', '』')
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            foreach ($needle in $needles) {
                $count = [regex]::Matches([string]$range.Text, [regex]::Escape($needle)).Count
                if ($count -eq 0) { continue }
                $find = $range.Find
                $find.ClearFormatting()
                $find.Replacement.ClearFormatting()
                $find.Text = $needle
                $find.Replacement.Text = ""
                $find.Forward = $true
                $find.Wrap = $wdFindContinue
                $find.Format = $false
                [void]$find.Execute($needle, $false, $false, $false, $false, $false, $true, $wdFindContinue, $false, "", $wdReplaceAll, $false, $false, $false, $false)
                $removed += $count
            }
            $range = $range.NextStoryRange
        }
    }
    return $removed
}

function Replace-ManualYears($Document) {
    $changed = 0
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            $matches = [regex]::Matches([string]$range.Text, '(?:19|20)\d{2}(?:[\s,，、]*(?:19|20)\d{2})*\s*年')
            for ($index = $matches.Count - 1; $index -ge 0; $index--) {
                $match = $matches[$index]
                $target = $range.Duplicate
                $target.Start = $range.Start + $match.Index
                $target.End = $target.Start + $match.Length
                $target.Text = "目前"
                $changed++
            }
            $range = $range.NextStoryRange
        }
    }
    return $changed
}

function Get-ManualPatternCount($Document, [string]$Pattern) {
    $total = 0
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            $total += [regex]::Matches([string]$range.Text, $Pattern).Count
            $range = $range.NextStoryRange
        }
    }
    return $total
}

function Get-ManualTextChanges($Document, [bool]$RemoveQuotes, [bool]$ReplaceYears) {
    # Collect the exact paragraph-level changes before Word mutates the text.
    # Paragraph indexes are unstable after Word updates fields, lists, or page
    # breaks, so they must never be used to pair before/after text in the report.
    $changes = @()
    if (-not $RemoveQuotes -and -not $ReplaceYears) { return @($changes) }
    for ($index = 1; $index -le $Document.Paragraphs.Count; $index++) {
        $before = Clean-Text $Document.Paragraphs.Item($index).Range.Text
        if ([string]::IsNullOrWhiteSpace($before)) { continue }
        $after = $before
        if ($RemoveQuotes) {
            $after = $after -replace '["“”「」『』]', ''
        }
        if ($ReplaceYears) {
            $after = $after -replace '(?:19|20)\d{2}(?:[\s,，、]*(?:19|20)\d{2})*\s*年', '目前'
        }
        if ($before -ne $after) {
            $changes += [pscustomobject]@{ paragraph = $index; before = $before; after = $after }
        }
    }
    return @($changes)
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

function Find-CoverBlock($Items, [string]$Title, [int]$MinimumPage, [int]$MaximumPage = 4, [bool]$IncludeLegacy = $false) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Page -gt $MaximumPage -or -not (Test-CoverTitle $item.Text $Title)) { continue }
        $limit = [Math]::Min($Items.Count - 1, $i + 6)
        for ($j = $i + 1; $j -le $limit; $j++) {
            if ($Items[$j].Page -ne $item.Page) { break }
            if ($Items[$j].Text -eq "") { continue }
            if ((Test-ManualLabel $Items[$j].Text) -or ($IncludeLegacy -and (Test-LegacyManualLabel $Items[$j].Text))) {
                return [pscustomobject]@{ TitleIndex = $i; LabelIndex = $j; Page = $item.Page }
            }
            # A body heading or paragraph means this is no longer a cover
            # candidate. Never scan past it to delete a later label.
            break
        }
    }
    return $null
}

function Find-IsolatedTitleCover($Items, [string]$Title, [int]$MinimumPage, [int]$MaximumPage = 10) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Page -gt $MaximumPage -or -not (Test-CoverTitle $item.Text $Title)) { continue }
        $pageText = @($Items | Where-Object { $_.Page -eq $item.Page -and $_.Text -ne "" })
        if ($pageText.Count -eq 1) { return [pscustomobject]@{ TitleIndex = $i; Page = $item.Page } }
    }
    return $null
}

function Find-IsolatedManualLabelCover($Items, [int]$MinimumPage, [int]$MaximumPage = 10) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Page -gt $MaximumPage -or -not (Test-LegacyManualLabel $item.Text)) { continue }
        $pageText = @($Items | Where-Object { $_.Page -eq $item.Page -and $_.Text -ne "" })
        if ($pageText.Count -eq 1) { return [pscustomobject]@{ LabelIndex = $i; Page = $item.Page } }
    }
    return $null
}

function Find-ResidualManualLabel($Items, [int]$MinimumPage, [int]$MaximumPage = 10) {
    foreach ($item in $Items) {
        if ($item.Page -ge $MinimumPage -and $item.Page -le $MaximumPage -and (Test-LegacyManualLabel $item.Text)) {
            return $item
        }
    }
    return $null
}

function Test-FirstPageCover($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    # Existing covers may say "说明书" or another legacy manual name. They
    # are still valid supplied covers and must not cause a duplicate insertion.
    return $null -ne (Find-CoverBlock $items $Title 1 1 $true)
}

function Test-CoverNeedsReplacement($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $block = Find-CoverBlock $items $Title 1 1 $true
    if ($null -eq $block) { return $false }
    $titleItem = $items[$block.TitleIndex]
    $labelItem = $items[$block.LabelIndex]
    $otherPageText = @($items | Where-Object {
        $_.Page -eq 1 -and $_.Index -ne $titleItem.Index -and $_.Index -ne $labelItem.Index -and $_.Text -ne ""
    })
    # When a supplied cover shares its page with the first heading or body,
    # preserve all original content. Typography plus a page break can make it
    # a valid cover without deleting title or body text.
    if ($otherPageText.Count -gt 0) { return $false }
    $titleRange = $Document.Range($titleItem.Start, $titleItem.End)
    $labelRange = $Document.Range($labelItem.Start, $labelItem.End)
    if ([int]$titleRange.ParagraphFormat.Alignment -ne $wdAlignParagraphCenter -or [int]$labelRange.ParagraphFormat.Alignment -ne $wdAlignParagraphCenter) { return $true }
    $titleSize = [double]$titleRange.Font.Size
    $labelSize = [double]$labelRange.Font.Size
    if ($titleSize -le 0 -or $labelSize -le 0) { return $true }
    return ([Math]::Abs($titleSize - $labelSize) -ge 4 -or [Math]::Max($titleSize, $labelSize) / [Math]::Min($titleSize, $labelSize) -ge 1.5)
}

function Remove-InvalidCoverFragments($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $fragments = @(
        $items | Where-Object { (Test-CoverTitle $_.Text $Title) -or (Test-LegacyManualLabel $_.Text) }
    )
    if ($fragments.Count -eq 0) { return $false }

    $fragmentIndexes = @($fragments | ForEach-Object { $_.Index })
    $otherText = @($items | Where-Object {
        $_.Text -ne "" -and $_.Index -notin $fragmentIndexes
    })
    $hasObjects = $false
    foreach ($shape in @($Document.InlineShapes)) {
        if ([int]$shape.Range.Information($wdActiveEndPageNumber) -eq 1) { $hasObjects = $true; break }
    }
    if (-not $hasObjects) {
        foreach ($table in @($Document.Tables)) {
            if ([int]$table.Range.Information($wdActiveEndPageNumber) -eq 1) { $hasObjects = $true; break }
        }
    }
    if ($otherText.Count -eq 0 -and -not $hasObjects) {
        $pageStart = [int]$Document.GoTo(1, 1, 1).Start
        $nextPage = $Document.GoTo(1, 1, 2)
        $pageEnd = if ($null -ne $nextPage -and [int]$nextPage.Start -gt $pageStart) { [int]$nextPage.Start } else { [int]$Document.Content.End }
        $Document.Range($pageStart, $pageEnd).Delete() | Out-Null
        return $true
    }

    # A page that contains real content is never a disposable cover page.  In
    # particular, keep a title or manual label when it shares the page with a
    # heading, body paragraph, table, or image.  A later normalized cover is
    # preferable to deleting supplied material.
    return $false
}

function Add-Cover($Document, [string]$Title) {
    $coverText = $Title + "`r`r`r使用说明书`r"
    $Document.Range(0, 0).InsertBefore($coverText)
    # A section break at document position zero fails for several supplied
    # templates. Insert it after the new cover text instead.
    $breakRange = $Document.Range($coverText.Length, $coverText.Length)
    $breakRange.InsertBreak($wdSectionBreakNextPage)
    # InsertBefore inherits the first paragraph's list template. Clear every
    # cover paragraph before applying cover typography.
    $coverRange = $Document.Range(0, [Math]::Min($Document.Content.End, $Title.Length + 12))
    foreach ($paragraph in $coverRange.Paragraphs) {
        try { $paragraph.Range.ListFormat.RemoveNumbers($wdNumberAllNumbers) } catch {}
        try { $paragraph.Range.Style = $Document.Styles.Item($wdStyleNormal) } catch {}
        try { $paragraph.Range.ParagraphFormat.OutlineLevel = $wdOutlineLevelBodyText } catch {}
    }
    $Document.Sections.Item(1).PageSetup.VerticalAlignment = $wdVerticalCenter
    if ($Document.Sections.Count -gt 1) { $Document.Sections.Item(2).PageSetup.VerticalAlignment = $wdVerticalTop }
}

function Remove-DuplicateCoverBlocks($Document, [string]$Title) {
    $removed = 0
    for ($round = 0; $round -lt 8; $round++) {
        # A newly inserted cover can move a legacy standalone cover beyond the
        # initial four pages when the original document begins with section
        # breaks.  It is still front matter, so inspect the first ten pages.
        $items = @(Get-FrontParagraphs $Document 10)
        $block = Find-CoverBlock $items $Title 2 10 $true
        $isolated = $null
        if ($null -eq $block) { $isolated = Find-IsolatedTitleCover $items $Title 2 10 }
        if ($null -eq $block -and $null -eq $isolated) { $isolated = Find-IsolatedManualLabelCover $items 2 10 }
        $residualLabel = $null
        if ($null -eq $block -and $null -eq $isolated) { $residualLabel = Find-ResidualManualLabel $items 2 10 }
        if ($null -eq $block -and $null -eq $isolated -and $null -eq $residualLabel) { break }

        if ($null -ne $residualLabel) {
            # The exact manual-label phrase is residual front matter when it
            # appears after a valid cover. Delete only that paragraph so any
            # heading, body text, table, or image sharing its page is retained.
            $Document.Range($residualLabel.Start, $residualLabel.End).Delete() | Out-Null
            $removed++
            continue
        }

        if ($null -ne $isolated) {
            $hasObjects = $false
            foreach ($shape in @($Document.InlineShapes)) {
                if ([int]$shape.Range.Information($wdActiveEndPageNumber) -eq $isolated.Page) { $hasObjects = $true; break }
            }
            if (-not $hasObjects) {
                foreach ($table in @($Document.Tables)) {
                    if ([int]$table.Range.Information($wdActiveEndPageNumber) -eq $isolated.Page) { $hasObjects = $true; break }
                }
            }
            if ($hasObjects) { break }
            $pageStart = [int]$Document.GoTo(1, 1, $isolated.Page).Start
            $nextPage = $Document.GoTo(1, 1, $isolated.Page + 1)
            $pageEnd = if ($null -ne $nextPage -and [int]$nextPage.Start -gt $pageStart) { [int]$nextPage.Start } else { [int]$Document.Content.End }
            $Document.Range($pageStart, $pageEnd).Delete() | Out-Null
            $removed++
            continue
        }

        $safe = $true
        for ($i = $block.TitleIndex + 1; $i -lt $block.LabelIndex; $i++) {
            if ($items[$i].Text -ne "") { $safe = $false; break }
        }
        if (-not $safe) { break }

        $startIndex = $block.TitleIndex
        while ($startIndex -gt 0 -and
               $items[$startIndex - 1].Page -eq $block.Page -and
               $items[$startIndex - 1].Text -eq "" -and
               ($block.TitleIndex - $startIndex) -lt 3) {
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

        $otherText = @($items | Where-Object {
            $_.Page -eq $block.Page -and ($_.Index -lt $items[$startIndex].Index -or $_.Index -gt $items[$endIndex].Index) -and $_.Text -ne ""
        })
        $hasObjects = $false
        foreach ($shape in @($Document.InlineShapes)) {
            if ([int]$shape.Range.Information($wdActiveEndPageNumber) -eq $block.Page) { $hasObjects = $true; break }
        }
        if (-not $hasObjects) {
            foreach ($table in @($Document.Tables)) {
                if ([int]$table.Range.Information($wdActiveEndPageNumber) -eq $block.Page) { $hasObjects = $true; break }
            }
        }
        if ($otherText.Count -eq 0 -and -not $hasObjects) {
            $pageStart = [int]$Document.GoTo(1, 1, $block.Page).Start
            $nextPage = $Document.GoTo(1, 1, $block.Page + 1)
            $pageEnd = if ($null -ne $nextPage -and [int]$nextPage.Start -gt $pageStart) { [int]$nextPage.Start } else { [int]$Document.Content.End }
            $Document.Range($pageStart, $pageEnd).Delete() | Out-Null
        } else {
            # This is not a standalone duplicate cover.  Do not delete a
            # title/label block from a page that also contains material.
            break
        }
        $removed++
        # Delete only the duplicate text block. Never delete the whole page.
    }
    return $removed
}

function Test-DuplicateCoverBlock($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 10)
    return ($null -ne (Find-CoverBlock $items $Title 2 10 $true) -or $null -ne (Find-IsolatedTitleCover $items $Title 2 10) -or $null -ne (Find-IsolatedManualLabelCover $items 2 10) -or $null -ne (Find-ResidualManualLabel $items 2 10))
}

function Set-CoverTypography($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $titleItem = $items | Where-Object { Test-CoverTitle $_.Text $Title } | Select-Object -First 1
    $labelItem = $items | Where-Object { Test-LegacyManualLabel $_.Text } | Select-Object -First 1
    if ($null -eq $titleItem -or $null -eq $labelItem) {
        return [pscustomobject]@{ Found = $false; Size = 0; Lines = 0; Label = "" }
    }

    $titleRange = $Document.Range($titleItem.Start, $titleItem.End)
    $labelRange = $Document.Range($labelItem.Start, $labelItem.End)
    try {
        $coverSection = [int]$titleRange.Information($wdActiveEndSectionNumber)
        $Document.Sections.Item($coverSection).PageSetup.VerticalAlignment = $wdVerticalCenter
    } catch {}
    foreach ($range in @($titleRange, $labelRange)) {
        foreach ($paragraph in $range.Paragraphs) {
            try { $paragraph.Range.ListFormat.RemoveNumbers($wdNumberAllNumbers) } catch {}
            try { $paragraph.Range.Style = $Document.Styles.Item($wdStyleNormal) } catch {}
            try { $paragraph.Range.ParagraphFormat.OutlineLevel = $wdOutlineLevelBodyText } catch {}
        }
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
    $titleRange.ParagraphFormat.SpaceBefore = 0
    $titleRange.ParagraphFormat.SpaceAfter = 0
    $labelRange.ParagraphFormat.SpaceBefore = 48
    $labelRange.ParagraphFormat.SpaceAfter = 0

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
    # Word's line-statistics can undercount a final wrapped CJK glyph on very
    # long centered titles. Keep a one-point margin for those borderline cases.
    if ($Title.Length -ge 20 -and $size -gt 14) {
        $candidate = $size - 1
        $titleRange.Font.Size = $candidate
        $labelRange.Font.Size = $candidate
        $Document.Repaginate()
        $candidateLines = [int]$titleRange.ComputeStatistics($wdStatisticLines)
        if ($candidateLines -le 1) {
            $size = $candidate
            $lines = $candidateLines
        } else {
            # The one-point visual margin must never reintroduce a wrapped title.
            $titleRange.Font.Size = $size
            $labelRange.Font.Size = $size
            $Document.Repaginate()
        }
    }
    return [pscustomobject]@{ Found = $true; Size = $size; Lines = $lines; Label = $labelItem.Text }
}

function Remove-CoverVersion($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $titleItem = $items | Where-Object { Test-CoverTitle $_.Text $Title } | Select-Object -First 1
    if ($null -eq $titleItem) { return [pscustomobject]@{ Found = $false; Changed = $false } }
    # Exclude the trailing paragraph mark so replacing the text keeps the
    # paragraph break and does not merge the title line with the label line.
    $titleRange = $Document.Range($titleItem.Start, [Math]::Max($titleItem.Start, $titleItem.End - 1))
    if ((Compact-Text $titleRange.Text) -eq (Compact-Text $Title)) {
        return [pscustomobject]@{ Found = $true; Changed = $false }
    }
    $titleRange.Text = $Title
    return [pscustomobject]@{ Found = $true; Changed = $true }
}

function Ensure-CoverPageBoundary($Document, [string]$Title) {
    # If an existing title/label block shares page 1 with body text or a TOC,
    # insert one page break after the label. Preserve all following content.
    $items = @(Get-FrontParagraphs $Document 1)
    $block = Find-CoverBlock $items $Title 1 1 $true
    if ($null -eq $block) {
        # Some supplied covers put the title and manual label in separate
        # paragraphs or apply Word list numbering. Find the first title and
        # the nearest manual label on page one without requiring adjacency.
        $titleItem = $items | Where-Object { (Compact-Text $_.Text) -eq (Compact-Text $Title) } | Select-Object -First 1
        $labelItem = $items | Where-Object { $_.Page -eq 1 -and (Test-ManualLabel $_.Text) } | Select-Object -First 1
        if ($null -eq $titleItem -or $null -eq $labelItem) { return $false }
        $block = [pscustomobject]@{ LabelIndex = [array]::IndexOf($items, $labelItem) }
    }
    $extra = @($items | Where-Object {
        $_.Page -eq 1 -and $_.Index -gt $items[$block.LabelIndex].Index -and $_.Text -ne ""
    })
    if ($extra.Count -eq 0) { return $false }
    $position = [int]$items[$block.LabelIndex].End - 1
    if ($position -lt 0) { return $false }
    $Document.Range($position, $position).InsertBreak($wdPageBreak)
    return $true
}

function Remove-BodyLeadingBlankParagraphs($Document, [string]$Title) {
    # The body must start at the very top of the page after the cover. Supplied
    # templates often carry empty paragraphs between the cover and the first
    # heading; remove those body-leading blanks so the first heading pins to the
    # top of page two. Paragraphs holding a page/section break are kept.
    $removed = 0
    for ($round = 0; $round -lt 10; $round++) {
        $items = @(Get-FrontParagraphs $Document 4)
        $cover = Find-CoverBlock $items $Title 1 1 $true
        if ($null -eq $cover) { break }
        $coverPage = [int]$items[$cover.LabelIndex].Page
        $target = $null
        for ($index = $cover.LabelIndex + 1; $index -lt $items.Count; $index++) {
            if ($items[$index].Text -ne "") { $target = $items[$index]; break }
        }
        if ($null -eq $target) { break }
        $deleted = 0
        foreach ($blank in @($items | Where-Object {
            $_.Index -gt $items[$cover.LabelIndex].Index -and $_.Index -lt $target.Index -and $_.Text -eq "" -and $_.Page -gt $coverPage
        } | Sort-Object Index -Descending)) {
            $raw = $Document.Range($blank.Start, $blank.End).Text
            $hasBreak = $raw.Contains([char]12)
            $pageBreakBefore = $false
            try { $pageBreakBefore = [bool]$Document.Paragraphs.Item($blank.Index).Range.ParagraphFormat.PageBreakBefore } catch { $pageBreakBefore = $true }
            if ($hasBreak -or $pageBreakBefore) { continue }
            $Document.Range($blank.Start, $blank.End).Delete() | Out-Null
            $deleted++
        }
        if ($deleted -eq 0) { break }
        $removed += $deleted
    }
    return $removed
}

function Story-HasContent($Story) {
    if (-not $Story.Exists) { return $false }
    if ((Clean-Text $Story.Range.Text) -ne "") { return $true }
    if (@(Get-PageFields $Story).Count -gt 0) { return $true }
    try { if ($Story.Range.InlineShapes.Count -gt 0 -or $Story.Range.ShapeRange.Count -gt 0) { return $true } } catch {}
    return $false
}

function Get-ActiveHeaders($Section) {
    $headers = @($Section.Headers.Item($wdHeaderPrimary))
    if ($Section.PageSetup.DifferentFirstPageHeaderFooter) { $headers += $Section.Headers.Item($wdHeaderFirstPage) }
    if ($Section.PageSetup.OddAndEvenPagesHeaderFooter) { $headers += $Section.Headers.Item($wdHeaderEvenPages) }
    return @($headers)
}

function Get-ActiveStories($Section) {
    $stories = @(Get-ActiveHeaders $Section)
    $stories += $Section.Footers.Item($wdHeaderPrimary)
    if ($Section.PageSetup.DifferentFirstPageHeaderFooter) { $stories += $Section.Footers.Item($wdHeaderFirstPage) }
    if ($Section.PageSetup.OddAndEvenPagesHeaderFooter) { $stories += $Section.Footers.Item($wdHeaderEvenPages) }
    return @($stories)
}

function Get-ActiveFooters($Section) {
    $footers = @($Section.Footers.Item($wdHeaderPrimary))
    if ($Section.PageSetup.DifferentFirstPageHeaderFooter) { $footers += $Section.Footers.Item($wdHeaderFirstPage) }
    if ($Section.PageSetup.OddAndEvenPagesHeaderFooter) { $footers += $Section.Footers.Item($wdHeaderEvenPages) }
    return @($footers)
}

function Story-HasPageField($Story) {
    return @(Get-PageFields $Story).Count -gt 0
}

function Get-PageFieldsFromRange($Range) {
    $fields = @()
    if ($null -eq $Range) { return @() }
    for ($index = 1; $index -le $Range.Fields.Count; $index++) {
        try {
            $field = $Range.Fields.Item($index)
            if ((Clean-Text $field.Code.Text) -match '(?i)\bPAGE\b') { $fields += $field }
        } catch {}
    }
    return @($fields)
}

function Get-PageFields($Story) {
    if (-not $Story.Exists) { return @() }
    # Existing page numbers are frequently stored in header text boxes.  Those
    # fields are not exposed by HeaderFooter.Range.Fields, so inspect shapes
    # first and keep their original layout ahead of any plain-header repair.
    $fields = @()
    try {
        $shapes = $Story.Range.ShapeRange
        for ($index = 1; $index -le $shapes.Count; $index++) {
            try { $fields += @(Get-PageFieldsFromRange $shapes.Item($index).TextFrame.TextRange) } catch {}
        }
    } catch {}
    $fields += @(Get-PageFieldsFromRange $Story.Range)
    return @($fields)
}

function Story-HasHeaderText($Story) {
    if (-not $Story.Exists) { return $false }
    $text = [string]$Story.Range.Text
    foreach ($field in @(Get-PageFields $Story)) {
        $text = $text.Replace([string]$field.Result.Text, "")
    }
    return -not [string]::IsNullOrWhiteSpace((Clean-Text $text))
}

function Get-PageFieldEntries($Section) {
    $entries = @()
    foreach ($story in @(Get-ActiveStories $Section)) {
        foreach ($field in @(Get-PageFields $story)) {
            $entries += [pscustomobject]@{ Story = $story; Field = $field; IsFooter = ($story.Index -ge $wdHeaderPrimary) }
        }
    }
    return @($entries)
}

function Remove-PageField($Field) {
    try { $Field.Delete(); return } catch {}
    try { $Field.Result.Delete() } catch {}
}

function Add-StandardHeader($Header, $Section, [string]$HeaderText) {
    $Header.LinkToPrevious = $false
    $range = $Header.Range
    $range.Text = $HeaderText
    $range.Font.Name = "宋体"
    $range.Font.NameFarEast = "宋体"
    $range.Font.NameAscii = "SimSun"
    $range.Font.NameOther = "SimSun"
    $range.Font.Size = 9
    $range.Font.Bold = 0
    $range.ParagraphFormat.Alignment = $wdAlignParagraphLeft
    $range.ParagraphFormat.TabStops.ClearAll()
    $width = $Section.PageSetup.PageWidth - $Section.PageSetup.LeftMargin - $Section.PageSetup.RightMargin
    [void]$range.ParagraphFormat.TabStops.Add($width, $wdAlignTabRight, $wdTabLeaderSpaces)
}

function Add-HeaderTextPreservingFields($Header, $Section, [string]$HeaderText) {
    $Header.LinkToPrevious = $false
    $fields = @(Get-PageFields $Header)
    if ($fields.Count -gt 0) {
        $insert = $Header.Range.Duplicate
        $insert.Collapse($wdCollapseStart)
        $insert.InsertBefore($HeaderText + "`t")
    } else {
        Add-StandardHeader $Header $Section $HeaderText
        return
    }
    $range = $Header.Range.Duplicate
    $range.End = [Math]::Min($range.End, $range.Start + $HeaderText.Length + 1)
    $range.Font.Name = "宋体"
    $range.Font.NameFarEast = "宋体"
    $range.Font.NameAscii = "SimSun"
    $range.Font.NameOther = "SimSun"
    $range.Font.Size = 9
    $range.Font.Bold = 0
}

function Add-PageField($Story) {
    $pageRange = $Story.Range.Duplicate
    if ($pageRange.End -gt $pageRange.Start) { $pageRange.End-- }
    $pageRange.Collapse($wdCollapseEnd)
    $pageRange.InsertAfter("`t")
    $pageRange.Collapse($wdCollapseEnd)
    [void]$Story.Range.Fields.Add($pageRange, $wdFieldPage)
}

function Get-PageFieldCount($Story) {
    return @(Get-PageFields $Story).Count
}

function Remove-ExtraPageFields($Header, $Footer) {
    $headerFields = @(Get-PageFields $Header)
    $footerFields = @(Get-PageFields $Footer)
    # Preserve a supplied footer number where available: it is commonly the
    # original format (for example "第 {PAGE} 页") and must win over any
    # header number added by an earlier repair attempt.
    $keepFooter = $footerFields.Count -gt 0
    $changed = $false
    for ($index = $headerFields.Count - 1; $index -ge 0; $index--) {
        if (-not $keepFooter -and $index -eq 0) { continue }
        Remove-PageField $headerFields[$index]
        $changed = $true
    }
    for ($index = $footerFields.Count - 1; $index -ge 0; $index--) {
        if ($keepFooter -and $index -eq 0) { continue }
        Remove-PageField $footerFields[$index]
        $changed = $true
    }
    return $changed
}

function Get-ExistingHeaderTemplate($Document) {
    foreach ($section in $Document.Sections) {
        foreach ($header in @(Get-ActiveHeaders $section)) {
            if (Story-HasHeaderText $header) { return $header }
        }
    }
    return $null
}

function Get-ExistingPageTemplate($Document) {
    # Prefer an existing footer page number because its literal format can be
    # "第 {PAGE} 页" or "-- {PAGE} --" and must remain untouched.
    foreach ($section in $Document.Sections) {
        foreach ($footer in @(Get-ActiveFooters $section)) {
            if ((Get-PageFieldCount $footer) -gt 0) { return [pscustomobject]@{ Story = $footer; IsFooter = $true } }
        }
    }
    foreach ($section in $Document.Sections) {
        foreach ($header in @(Get-ActiveHeaders $section)) {
            if ((Get-PageFieldCount $header) -gt 0) { return [pscustomobject]@{ Story = $header; IsFooter = $false } }
        }
    }
    return $null
}

function Copy-StoryTemplate($Source, $Target) {
    $Target.LinkToPrevious = $false
    $Target.Range.FormattedText = $Source.Range.FormattedText
}

function Test-HeaderPageLayout($Document) {
    foreach ($section in $Document.Sections) {
        $headers = @(Get-ActiveHeaders $section)
        $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            if (-not (Story-HasHeaderText $headers[$index])) { return $false }
            if ((Get-PageFieldCount $headers[$index]) + (Get-PageFieldCount $footers[$index]) -ne 1) { return $false }
        }
    }
    return $true
}

function Set-Headers($Document, [string]$Title) {
    $cleanTitle = ($Title -replace '(?i)\s*V\s*1\.0\s*$', '').Trim()
    $headerText = $cleanTitle + " V1.0"
    if (Test-HeaderPageLayout $Document) { return $false }
    $changed = $false
    $headerTemplate = Get-ExistingHeaderTemplate $Document
    $pageTemplate = Get-ExistingPageTemplate $Document
    for ($sectionIndex = 1; $sectionIndex -le $Document.Sections.Count; $sectionIndex++) {
        $section = $Document.Sections.Item($sectionIndex)
        $headers = @(Get-ActiveHeaders $section)
        $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            $header = $headers[$index]
            $footer = $footers[$index]
            if (-not (Story-HasHeaderText $header)) {
                if ($null -ne $headerTemplate) { Copy-StoryTemplate $headerTemplate $header }
                else { Add-StandardHeader $header $section $headerText }
                $changed = $true
            }
            if (Remove-ExtraPageFields $header $footer) { $changed = $true }
            # The copied header template may already contain PAGE. Recount only
            # after copying so a footer template never creates a second number.
            $pageCount = (Get-PageFieldCount $header) + (Get-PageFieldCount $footer)
            if ($pageCount -eq 0) {
                if ($null -ne $pageTemplate) {
                    $target = if ($pageTemplate.IsFooter) { $footer } else { $header }
                    if (-not (Story-HasContent $target)) { Copy-StoryTemplate $pageTemplate.Story $target }
                    else { Add-PageField $target }
                } else {
                    Add-PageField $header
                }
                $changed = $true
            }
        }
    }
    return $changed
}

function Audit-EffectiveCover($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $titleItem = $items | Where-Object { (Compact-Text $_.Text) -eq (Compact-Text $Title) } | Select-Object -First 1
    $labelItem = $items | Where-Object { Test-ManualLabel $_.Text } | Select-Object -First 1
    if ($null -eq $titleItem -or $null -eq $labelItem) {
        return [pscustomobject]@{ Found = $false; Font = ""; LabelFont = ""; Bold = $false; Size = 0; LabelSize = 0; Lines = 0; ExtraText = @(); TitleYRatio = -1; LabelYRatio = -1; GapPoints = -1 }
    }
    $titleRange = $Document.Range($titleItem.Start, $titleItem.End)
    $labelRange = $Document.Range($labelItem.Start, $labelItem.End)
    # List numbering is not included in Range.Text, so inspect the owning
    # paragraphs explicitly. A nonzero ListType means Word will render a number
    # or bullet before the cover text even when its visible text looks correct.
    $titleListType = 0
    $labelListType = 0
    try { $titleListType = [int]$Document.Paragraphs.Item($titleItem.Index).Range.ListFormat.ListType } catch {}
    try { $labelListType = [int]$Document.Paragraphs.Item($labelItem.Index).Range.ListFormat.ListType } catch {}
    $pageHeight = [double]$Document.Sections.Item(1).PageSetup.PageHeight
    $titleY = [double]$titleRange.Information($wdVerticalPositionRelativeToPage)
    $labelY = [double]$labelRange.Information($wdVerticalPositionRelativeToPage)
    $extraText = @($items | Where-Object {
        ($_.Text -ne "" -and (Compact-Text $_.Text) -ne (Compact-Text $Title) -and -not (Test-ManualLabel $_.Text) -and $_.Text -match '[\p{L}\p{N}]')
    } | ForEach-Object { $_.Text })
    return [pscustomobject]@{
        Found = $true
        Font = [string]$titleRange.Font.NameFarEast
        LabelFont = [string]$labelRange.Font.NameFarEast
        Bold = ([int]$titleRange.Font.Bold -eq -1 -and [int]$labelRange.Font.Bold -eq -1)
        Size = [double]$titleRange.Font.Size
        LabelSize = [double]$labelRange.Font.Size
        Lines = [int]$titleRange.ComputeStatistics($wdStatisticLines)
        TitleYRatio = if ($pageHeight -gt 0 -and $titleY -ge 0) { [Math]::Round($titleY / $pageHeight, 3) } else { -1 }
        LabelYRatio = if ($pageHeight -gt 0 -and $labelY -ge 0) { [Math]::Round($labelY / $pageHeight, 3) } else { -1 }
        GapPoints = if ($labelY -ge 0 -and $titleY -ge 0) { [Math]::Round($labelY - $titleY, 1) } else { -1 }
        ExtraText = $extraText
        TitleNumbered = ($titleListType -ne 0)
        LabelNumbered = ($labelListType -ne 0)
    }
}

function Audit-Headers($Document, [string]$Title) {
    $headerFound = $true
    $pageFound = $true
    foreach ($section in $Document.Sections) {
        $headers = @(Get-ActiveHeaders $section)
        $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            if (-not (Story-HasHeaderText $headers[$index])) { $headerFound = $false }
            if ((Get-PageFieldCount $headers[$index]) + (Get-PageFieldCount $footers[$index]) -ne 1) { $pageFound = $false }
        }
    }
    return [pscustomobject]@{ Header = $headerFound; Page = $pageFound }
}

function Safe-Name([string]$Value) {
    return ($Value -replace '[<>:"/\\|?*]', '_')
}

$base = (Resolve-Path -LiteralPath $BaseDir).Path
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$reportDir = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($ReportPath))
$backupRoot = Join-Path $reportDir "backups"
[IO.Directory]::CreateDirectory($backupRoot) | Out-Null
$tracePath = Join-Path $reportDir "word_trace_stable.log"
[IO.File]::WriteAllText($tracePath, "", (New-Object Text.UTF8Encoding($false)))

function Trace-Step([string]$Title, [string]$Step) {
    [IO.File]::AppendAllText(
        $tracePath,
        (([DateTime]::Now.ToString("HH:mm:ss.fff") + "`t" + $Title + "`t" + $Step + "`r`n")),
        (New-Object Text.UTF8Encoding($false))
    )
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
if (-not [string]::IsNullOrWhiteSpace($PidPath)) {
    if (-not ("EasySoftware.NativeMethods" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace EasySoftware {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
"@
    }
    [uint32]$wordPid = 0
    try { [void][EasySoftware.NativeMethods]::GetWindowThreadProcessId([IntPtr]$word.Hwnd, [ref]$wordPid) } catch {}
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($PidPath))) | Out-Null
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($PidPath), [string]$wordPid, (New-Object Text.UTF8Encoding($false)))
}
$results = @(); $progressIndex = 0

try {
    foreach ($item in @($plan.packages)) {
        $title = [string]$item.name
        Write-WorkflowProgress $progressIndex $title $true
        $folder = Join-Path $base ([string]$item.folder)
        [string]$docx = Join-Path $folder ([string]$item.docx)
        [string]$pdf = if (-not [string]::IsNullOrWhiteSpace([string]$item.pdf)) {
            Join-Path $folder ([string]$item.pdf)
        } else {
            [IO.Path]::ChangeExtension($docx, ".pdf")
        }
        $actions = @($item.actions)
        $backupFolder = Join-Path $backupRoot (Safe-Name $title)
        [IO.Directory]::CreateDirectory($backupFolder) | Out-Null
        $backupDocx = Join-Path $backupFolder ([IO.Path]::GetFileName($docx))
        $backupPdf = Join-Path $backupFolder ([IO.Path]::GetFileName($pdf))
        $hadOriginalPdf = Test-Path -LiteralPath $pdf
        Copy-Item -LiteralPath $docx -Destination $backupDocx -Force
        if ($hadOriginalPdf) { Copy-Item -LiteralPath $pdf -Destination $backupPdf -Force }

        $doc = $null
        $started = [DateTime]::UtcNow
        $applied = @()
        $removed = 0
        $failures = @()
        try {
            Trace-Step $title "open"
            $doc = $word.Documents.Open($docx, $false, $false)
            Trace-Step $title "opened"

            $coverAdded = $false
            Trace-Step $title "cover-check"
            $replaceInvalidCover = Test-CoverNeedsReplacement $doc $title
            if (Test-Action $actions "ADD_COVER") {
                # ADD_COVER is emitted only after the programmatic PDF audit has
                # determined that page 1 is not a clean, independent cover.
                if ($replaceInvalidCover) {
                    Trace-Step $title "cover-remove-invalid"
                    [void](Remove-InvalidCoverFragments $doc $title)
                }
                if (-not (Test-FirstPageCover $doc $title)) {
                    try {
                        Trace-Step $title "cover-add"
                        Add-Cover $doc $title
                        Trace-Step $title "cover-added"
                        $coverAdded = $true
                        if ($replaceInvalidCover) { $applied += "REPLACE_INVALID_COVER" } else { $applied += "ADD_COVER" }
                    } catch { $failures += ("补封面失败：" + $_.Exception.Message) }
                }
            } elseif (Test-Action $actions "ENSURE_COVER_FIRST_PAGE") {
                if ($replaceInvalidCover) {
                    [void](Remove-InvalidCoverFragments $doc $title)
                }
                if (-not (Test-FirstPageCover $doc $title)) {
                    try {
                        Add-Cover $doc $title
                        $coverAdded = $true
                        if ($replaceInvalidCover) { $applied += "REPLACE_INVALID_COVER" } else { $applied += "ADD_COVER" }
                    } catch { $failures += ("补封面失败：" + $_.Exception.Message) }
                }
            }

            if ((Test-Action $actions "REMOVE_DUPLICATE_COVER_TEXT") -or $coverAdded) {
                Trace-Step $title "cover-remove-duplicates"
                $removed = Remove-DuplicateCoverBlocks $doc $title
                Trace-Step $title "cover-duplicates-removed"
                if ($removed -gt 0) { $applied += "REMOVE_DUPLICATE_COVER_TEXT" }
            }

            if (Test-Action $actions "REMOVE_COVER_VERSION") {
                Trace-Step $title "cover-remove-version"
                $versionFix = Remove-CoverVersion $doc $title
                if ($versionFix.Found -and $versionFix.Changed) { $applied += "REMOVE_COVER_VERSION" }
                elseif (-not $versionFix.Found) { $failures += "未找到封面标题，无法去掉版本号" }
            }

            $style = $null
            if ((Test-Action $actions "NORMALIZE_COVER_STYLE") -or $coverAdded) {
                if (-not $coverAdded -and $replaceInvalidCover) {
                    [void](Remove-InvalidCoverFragments $doc $title)
                    Add-Cover $doc $title
                    $coverAdded = $true
                    $applied += "REPLACE_INVALID_COVER"
                }
                Trace-Step $title "cover-style"
                $style = Set-CoverTypography $doc $title
                Trace-Step $title "cover-styled"
                if (-not $style.Found) { $failures += "第一页未找到可排版的封面标题块" }
                else { $applied += "NORMALIZE_COVER_STYLE" }
            }

            Trace-Step $title "cover-boundary"
            if (Ensure-CoverPageBoundary $doc $title) {
                $applied += "SEPARATE_COVER_FROM_BODY"
            }
            Trace-Step $title "cover-boundary-done"
            Trace-Step $title "body-leading-blanks"
            $bodyBlanks = Remove-BodyLeadingBlankParagraphs $doc $title
            Trace-Step $title "body-leading-blanks-done"
            if ($bodyBlanks -gt 0) { $applied += "TRIM_BODY_LEADING_BLANKS" }

            $bodyTextStyle = $null
            if (Test-Action $actions "NORMALIZE_BODY_TEXT_STYLE") {
                Trace-Step $title "body-style"
                $bodyTextStyle = Normalize-BodyTextStyles $doc $title
                Trace-Step $title "body-styled"
                if ($bodyTextStyle.Changed) { $applied += "NORMALIZE_BODY_TEXT_STYLE" }
            }

            $firstPostCoverHeading = $null
            if (Test-Action $actions "NORMALIZE_FIRST_POST_COVER_HEADING") {
                $firstPostCoverHeading = Normalize-FirstPostCoverHeading $doc $title
                if ($firstPostCoverHeading.Changed) { $applied += "NORMALIZE_FIRST_POST_COVER_HEADING" }
            }

            $textChanges = @(Get-ManualTextChanges $doc (Test-Action $actions "REMOVE_ALL_DOUBLE_QUOTES") (Test-Action $actions "REPLACE_YEARS_WITH_CURRENT"))
            $titleQuotesRemoved = 0
            if (Test-Action $actions "REMOVE_TITLE_DOUBLE_QUOTES") {
                $titleQuotesRemoved = Remove-TitleDoubleQuotes $doc $title
                if ($titleQuotesRemoved -gt 0) { $applied += "REMOVE_TITLE_DOUBLE_QUOTES" }
            }
            $allQuotesRemoved = 0
            if (Test-Action $actions "REMOVE_ALL_DOUBLE_QUOTES") {
                $allQuotesRemoved = Replace-AllManualQuotes $doc
                if ($allQuotesRemoved -gt 0) { $applied += "REMOVE_ALL_DOUBLE_QUOTES" }
            }
            $yearsReplaced = 0
            if (Test-Action $actions "REPLACE_YEARS_WITH_CURRENT") {
                $yearsReplaced = Replace-ManualYears $doc
                if ($yearsReplaced -gt 0) { $applied += "REPLACE_YEARS_WITH_CURRENT" }
            }

            # Every saved repair must leave a deterministic header/PAGE structure,
            # including the cover page.
            $headerChanged = Set-Headers $doc $title
            if ($headerChanged) { $applied += "NORMALIZE_HEADER" }

            Trace-Step $title "fields"
            foreach ($section in $doc.Sections) {
                foreach ($header in $section.Headers) {
                    if ($header.Exists) { $header.Range.Fields.Update() | Out-Null }
                }
            }
            Trace-Step $title "save"
            $doc.Save()
            Trace-Step $title "export"
            $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF)
            Trace-Step $title "exported"
            $applied += "EXPORT_PDF"

            $coverAudit = Audit-EffectiveCover $doc $title
            $headerAudit = Audit-Headers $doc $title
            if (Test-DuplicateCoverBlock $doc $title) { $failures += "仍存在后续重复封面标题块" }
            if (-not $coverAudit.Found) { $failures += "第一页缺少有效封面标题块" }
            if (@($coverAudit.ExtraText).Count -gt 0) { $failures += "封面页仍混有目录或正文文字" }
            if ($coverAudit.Lines -gt 1) { $failures += "软件名称仍然换行" }
            if (-not $coverAudit.Bold) { $failures += "封面两行未同时加粗" }
            if ($coverAudit.Size -ne $coverAudit.LabelSize) { $failures += "封面两行字号不一致" }
            if ($coverAudit.Font -notmatch '宋体|SimSun' -or $coverAudit.LabelFont -notmatch '宋体|SimSun') { $failures += "封面两行未同时使用宋体" }
            if ($coverAudit.TitleYRatio -ge 0 -and ($coverAudit.TitleYRatio -lt 0.30 -or $coverAudit.TitleYRatio -gt 0.56)) { $failures += "软件名称未位于封面中段偏上" }
            if ($coverAudit.LabelYRatio -ge 0 -and ($coverAudit.LabelYRatio -lt 0.48 -or $coverAudit.LabelYRatio -gt 0.72)) { $failures += "用户手册/使用说明书未位于封面中部附近" }
            if ($coverAudit.GapPoints -ge 0 -and $coverAudit.GapPoints -lt 48) { $failures += "软件名称与用户手册/使用说明书之间留白不足" }
            if ($coverAudit.TitleNumbered) { $failures += "封面标题仍带自动编号" }
            if ($coverAudit.LabelNumbered) { $failures += "封面说明书行仍带自动编号" }
            if (-not ($headerAudit.Header -and $headerAudit.Page)) { $failures += "页眉或动态PAGE字段缺失" }
            if ((Test-Action $actions "REMOVE_TITLE_DOUBLE_QUOTES") -and (Get-TitleDoubleQuoteCount $doc $title) -gt 0) { $failures += "全文仍存在外带双引号的软件名称" }
            if ((Test-Action $actions "REMOVE_ALL_DOUBLE_QUOTES") -and (Get-ManualPatternCount $doc '["“”「」『』]') -gt 0) { $failures += "说明书全文仍存在双引号" }
            if ((Test-Action $actions "REPLACE_YEARS_WITH_CURRENT") -and (Get-ManualPatternCount $doc '(?:19|20)\d{2}(?:[\s,，、]*(?:19|20)\d{2})*\s*年') -gt 0) { $failures += "说明书全文仍存在年份表述" }
            # A QA finding must not undo successfully applied page headers,
            # page numbers, cover separation, or PDF export. Keep the repaired
            # file and report the remaining defect for review.

            $results += [pscustomobject]@{
                name = $title
                folder = [string]$item.folder
                status = if ($failures.Count -eq 0) { "OK" } else { "QA_FAILED" }
                changed = $true
                actions_requested = $actions
                actions_applied = @($applied | Select-Object -Unique)
                duplicate_blocks_removed = $removed
                title_double_quotes_removed = $titleQuotesRemoved
                all_double_quotes_removed = $allQuotesRemoved
                years_replaced = $yearsReplaced
                text_changes = @($textChanges)
                body_text_style_found = if ($null -ne $bodyTextStyle) { $bodyTextStyle.Found } else { $false }
                body_text_style_normalized = if ($null -ne $bodyTextStyle) { $bodyTextStyle.Changed } else { $false }
                body_text_style_changed_paragraphs = if ($null -ne $bodyTextStyle) { $bodyTextStyle.ChangedCount } else { 0 }
                body_text_reference_font = if ($null -ne $bodyTextStyle) { $bodyTextStyle.ReferenceFont } else { "" }
                body_text_reference_size = if ($null -ne $bodyTextStyle) { $bodyTextStyle.ReferenceSize } else { 0 }
                first_post_cover_heading_normalized = if ($null -ne $firstPostCoverHeading) { $firstPostCoverHeading.Changed } else { $false }
                first_post_cover_heading_text = if ($null -ne $firstPostCoverHeading) { $firstPostCoverHeading.Text } else { "" }
                font_size = $coverAudit.Size
                title_lines = $coverAudit.Lines
                title_y_ratio = $coverAudit.TitleYRatio
                label_y_ratio = $coverAudit.LabelYRatio
                title_label_gap_points = $coverAudit.GapPoints
                cover_extra_text = @($coverAudit.ExtraText)
                word_pages = [int]$doc.ComputeStatistics($wdStatisticPages)
                failures = $failures
                seconds = [Math]::Round(([DateTime]::UtcNow - $started).TotalSeconds, 3)
            }
        } catch {
            $message = $_.Exception.Message
            if ($null -ne $doc) {
                # Preserve any changes completed before the failing step.
                try { $doc.Save() } catch {}
                try { $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF) } catch {}
                try { $doc.Close($false) } catch {}
                [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                $doc = $null
            }
            $results += [pscustomobject]@{
                name = $title
                folder = [string]$item.folder
                status = "PRESERVED_FOR_REVIEW"
                changed = $true
                actions_requested = $actions
                actions_applied = @()
                duplicate_blocks_removed = 0
                title_double_quotes_removed = 0
                body_text_style_found = $false
                body_text_style_normalized = $false
                body_text_style_changed_paragraphs = 0
                body_text_reference_font = ""
                body_text_reference_size = 0
                font_size = 0
                title_lines = 0
                word_pages = 0
                failures = @($message)
                seconds = [Math]::Round(([DateTime]::UtcNow - $started).TotalSeconds, 3)
            }
            Trace-Step $title ("failed=" + $message)
        } finally {
            if ($null -ne $doc) {
                $doc.Close($false)
                [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
        }
        $progressIndex++
        Write-WorkflowProgress $progressIndex $title
    }
} finally {
    # Word may close its COM server immediately after exporting the final PDF.
    # Cleanup must not prevent the already-built report from being written.
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$report = [ordered]@{
        version = "3.4.26"
    created_at = [DateTime]::Now.ToString("o")
    base = $base
    packages = $results
}
[IO.Directory]::CreateDirectory($reportDir) | Out-Null
[IO.File]::WriteAllText(
    [IO.Path]::GetFullPath($ReportPath),
    ($report | ConvertTo-Json -Depth 8),
    (New-Object Text.UTF8Encoding($false))
)

if (@($results | Where-Object { $_.status -eq "ROLLED_BACK" }).Count -gt 0) { exit 1 }
exit 0

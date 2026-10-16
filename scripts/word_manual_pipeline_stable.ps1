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
$wdAlignParagraphRight = 2
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
$wdFirstCharacterLineNumber = 10
$wdOutlineLevelBodyText = 10
$wdWithInTable = 12
$wdFindContinue = 1
$wdReplaceAll = 2
$wdNumberAllNumbers = 3
$wdStyleNormal = -1
$ProgressStartedAt = [DateTime]::UtcNow
. (Join-Path $PSScriptRoot "cover_rules.ps1")

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
    return (Get-CoverClassification $Text $Title).kind -eq "cover_title"
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
    return (Get-CoverClassification $Text).kind -eq "manual_label"
}

function Test-LegacyManualLabel([string]$Text) {
    return Test-ManualLabel $Text
}

function Test-Action($Actions, [string]$Name) {
    return @($Actions) -contains $Name
}

function Get-BookmarkFieldDescriptor($Field) {
    $fieldType = [int]$Field.Type
    $typeName = switch ($fieldType) {
        3 { "REF" }
        37 { "PAGEREF" }
        72 { "NOTEREF" }
        default { return $null }
    }
    $code = [string]$Field.Code.Text
    $pattern = '^(?is)\s*' + [regex]::Escape($typeName) + '\s+(?:"((?:[^"]|"")+)"|([^\s\\]+))'
    $match = [regex]::Match($code, $pattern)
    if (-not $match.Success) {
        throw ("无法解析 {0} 字段目标：{1}" -f $typeName, $code.Trim())
    }
    $target = if ($match.Groups[1].Success) { $match.Groups[1].Value -replace '""', '"' } else { $match.Groups[2].Value }
    if ([string]::IsNullOrWhiteSpace($target)) {
        throw ("{0} 字段目标为空：{1}" -f $typeName, $code.Trim())
    }
    return [pscustomobject]@{ Type = $typeName; Target = $target }
}

function Resolve-MissingBookmarkField([string]$FieldType, [string]$Target, [string]$CachedResult, [bool]$BookmarkExists) {
    if ($BookmarkExists) {
        return [pscustomobject]@{ Action = "KEEP"; CachedResult = $CachedResult; Reason = "目标书签存在" }
    }
    $visibleResult = ($CachedResult -replace '[\x00-\x1F\x7F]', '').Trim()
    if ([string]::IsNullOrWhiteSpace($Target)) {
        return [pscustomobject]@{ Action = "FAIL"; CachedResult = $visibleResult; Reason = ("{0} 字段目标无法解析" -f $FieldType) }
    }
    if ([string]::IsNullOrWhiteSpace($visibleResult)) {
        return [pscustomobject]@{ Action = "FAIL"; CachedResult = $visibleResult; Reason = ("{0} 字段目标书签 {1} 不存在且缓存结果为空" -f $FieldType, $Target) }
    }
    return [pscustomobject]@{ Action = "UNLINK"; CachedResult = $visibleResult; Reason = "保留字段缓存结果" }
}

function Repair-MissingBookmarkFields($Document) {
    $changes = @()
    $fields = $Document.Fields
    for ($index = [int]$fields.Count; $index -ge 1; $index--) {
        $field = $fields.Item($index)
        $descriptor = Get-BookmarkFieldDescriptor $field
        if ($null -eq $descriptor) { continue }
        $bookmarkExists = [bool]$Document.Bookmarks.Exists([string]$descriptor.Target)
        $cachedResult = [string]$field.Result.Text
        $decision = Resolve-MissingBookmarkField $descriptor.Type $descriptor.Target $cachedResult $bookmarkExists
        if ($decision.Action -eq "FAIL") { throw $decision.Reason }
        if ($decision.Action -eq "KEEP") { continue }
        try {
            $field.Unlink()
        } catch {
            throw ("无法保留 {0} 字段 {1} 的缓存结果：{2}" -f $descriptor.Type, $descriptor.Target, $_.Exception.Message)
        }
        $changes += [pscustomobject]@{
            type = [string]$descriptor.Type
            target = [string]$descriptor.Target
            cached_result = [string]$decision.CachedResult
        }
    }
    return @($changes)
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
        version = "3.4.28"; status = "RUNNING"; phase = $ProgressPhase
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

function Get-FrontParagraphs($Document, [int]$MaxPage = 5) {
    $items = New-Object 'System.Collections.Generic.List[object]'
    $paragraphs = $Document.Paragraphs
    $paragraphCount = [int]$paragraphs.Count
    for ($i = 1; $i -le $paragraphCount; $i++) {
        $paragraph = $paragraphs.Item($i)
        $range = $paragraph.Range
        $page = [int]$range.Information($wdActiveEndPageNumber)
        if ($page -gt $MaxPage) { break }
        $items.Add([pscustomobject]@{
            Index = $i
            Page = $page
            Start = [int]$range.Start
            End = [int]$range.End
            Text = Clean-Text $range.Text
            InTable = [bool]$range.Information($wdWithInTable)
            Style = [string]$range.Style
        })
    }
    return $items.ToArray()
}

function Find-CoverBlock($Items, [string]$Title, [int]$MinimumPage, [int]$MaximumPage = 5, [bool]$IncludeLegacy = $false) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Page -gt $MaximumPage -or $item.InTable -or $item.Style -match '(?i)TOC|目录' -or -not (Test-CoverTitle $item.Text $Title)) { continue }
        $limit = [Math]::Min($Items.Count - 1, $i + 6)
        for ($j = $i + 1; $j -le $limit; $j++) {
            if ($Items[$j].Page -ne $item.Page) { break }
            if ($Items[$j].Text -eq "") { continue }
            if (-not $Items[$j].InTable -and $Items[$j].Style -notmatch '(?i)TOC|目录' -and ((Test-ManualLabel $Items[$j].Text) -or ($IncludeLegacy -and (Test-LegacyManualLabel $Items[$j].Text)))) {
                return [pscustomobject]@{ TitleIndex = $i; LabelIndex = $j; Page = $item.Page }
            }
            # A body heading or paragraph means this is no longer a cover
            # candidate. Never scan past it to delete a later label.
            break
        }
    }
    return $null
}

function Find-LeadingTableCover($Document, $Items, [string]$Title) {
    # A supplied cover may be a table. Recognize only a first-page table whose
    # first visible paragraph is the exact software title and whose same table
    # also contains a manual label. This prevents Add-Cover from inserting at
    # Range(0, 0), which Word treats as the first cell when the document starts
    # with a table.
    $firstIndex = -1
    for ($index = 0; $index -lt $Items.Count; $index++) {
        if ($Items[$index].Page -eq 1 -and $Items[$index].Text -ne "") {
            $firstIndex = $index
            break
        }
    }
    if ($firstIndex -lt 0) { return $null }
    $firstItem = $Items[$firstIndex]
    if (-not $firstItem.InTable -or $firstItem.Style -match '(?i)TOC|目录' -or -not (Test-CoverTitle $firstItem.Text $Title)) { return $null }

    try {
        $tableRange = $Document.Tables.Item(1).Range
        $tableStart = [int]$tableRange.Start
        $tableEnd = [int]$tableRange.End
    } catch { return $null }
    if ([int]$firstItem.Start -lt $tableStart -or [int]$firstItem.End -gt $tableEnd) { return $null }

    for ($index = $firstIndex + 1; $index -lt $Items.Count; $index++) {
        $item = $Items[$index]
        if ($item.Page -gt 1) { break }
        if (-not $item.InTable -or [int]$item.Start -lt $tableStart -or [int]$item.End -gt $tableEnd) { continue }
        if ($item.Style -notmatch '(?i)TOC|目录' -and (Test-LegacyManualLabel $item.Text)) {
            return [pscustomobject]@{ TitleIndex = $firstIndex; LabelIndex = $index; Page = 1; TableStart = $tableStart; TableEnd = $tableEnd }
        }
    }
    return $null
}

function Find-IsolatedTitleCover($Items, [string]$Title, [int]$MinimumPage, [int]$MaximumPage = 5) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Page -gt $MaximumPage -or $item.InTable -or $item.Style -match '(?i)TOC|目录' -or -not (Test-CoverTitle $item.Text $Title)) { continue }
        $pageText = @($Items | Where-Object { $_.Page -eq $item.Page -and $_.Text -ne "" })
        if ($pageText.Count -eq 1) { return [pscustomobject]@{ TitleIndex = $i; Page = $item.Page } }
    }
    return $null
}

function Find-IsolatedManualLabelCover($Items, [int]$MinimumPage, [int]$MaximumPage = 5) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Page -gt $MaximumPage -or $item.InTable -or $item.Style -match '(?i)TOC|目录' -or -not (Test-LegacyManualLabel $item.Text)) { continue }
        $pageText = @($Items | Where-Object { $_.Page -eq $item.Page -and $_.Text -ne "" })
        if ($pageText.Count -eq 1) { return [pscustomobject]@{ LabelIndex = $i; Page = $item.Page } }
    }
    return $null
}

function Find-ResidualManualLabel($Items, [int]$MinimumPage, [int]$MaximumPage = 5) {
    foreach ($item in $Items) {
        if ($item.Page -ge $MinimumPage -and $item.Page -le $MaximumPage -and -not $item.InTable -and $item.Style -notmatch '(?i)TOC|目录' -and (Test-LegacyManualLabel $item.Text)) {
            return $item
        }
    }
    return $null
}

function Test-CombinedCoverFragment([string]$Text, [string]$Title) {
    $value = Compact-Text $Text
    $coverTitle = Compact-Text $Title
    if ($value -eq "" -or $coverTitle -eq "" -or -not $value.StartsWith($coverTitle)) { return $false }
    $suffix = $value.Substring($coverTitle.Length)
    return ($suffix -match '^(?:软件)?(?:使用)?(?:用户)?(?:设计)?说明(?:书|文档)?$' -or $suffix -match '^(?:用户|使用)手册$')
}

function Test-PostCoverFragment([string]$Text, [string]$Title) {
    $value = Compact-Text $Text
    if (Test-CoverTitle $Text $Title) { return $true }
    if (Test-ManualLabel $Text) { return $true }
    # Only a complete document-control line is classified as cover metadata.
    return $value -match '^文档编号[:：].+系统状态[:：].+机密程度[:：].+$'
}

function Find-IsolatedCombinedCoverFragment($Items, [string]$Title, [int]$MinimumPage, [int]$MaximumPage = 5) {
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($item.Page -lt $MinimumPage -or $item.Page -gt $MaximumPage -or -not (Test-CombinedCoverFragment $item.Text $Title)) { continue }
        $pageText = @($Items | Where-Object { $_.Page -eq $item.Page -and $_.Text -ne "" })
        if ($pageText.Count -eq 1) { return [pscustomobject]@{ FragmentIndex = $i; Page = $item.Page } }
    }
    return $null
}

function Test-PageHasObjects($Document, [int]$Page) {
    foreach ($shape in @($Document.InlineShapes)) {
        if ([int]$shape.Range.Information($wdActiveEndPageNumber) -eq $Page) { return $true }
    }
    foreach ($table in @($Document.Tables)) {
        if ([int]$table.Range.Information($wdActiveEndPageNumber) -eq $Page) { return $true }
    }
    foreach ($shape in @($Document.Shapes)) {
        if ([int]$shape.Anchor.Information($wdActiveEndPageNumber) -eq $Page) { return $true }
    }
    return $false
}

function Test-RangeHasProtectedContent($Document, $Range, [bool]$CheckParagraphBreakBefore = $false, [bool]$AllowPageBreak = $false, [bool]$AllowSectionBreak = $false) {
    # Clean-Text hides control characters and image anchors. Check the raw
    # range before deleting an apparently blank paragraph or an entire page.
    if ([string]$Range.Text -match '\x0E' -or (-not $AllowPageBreak -and [string]$Range.Text -match '\x0C')) { return $true }
    if ($CheckParagraphBreakBefore -and [bool]$Range.ParagraphFormat.PageBreakBefore) { return $true }
    foreach ($section in @($Document.Sections)) {
        $end = [int]$section.Range.End
        if (-not $AllowSectionBreak -and $end -gt [int]$Range.Start -and $end -le [int]$Range.End) { return $true }
    }
    foreach ($shape in @($Document.InlineShapes)) {
        if ([int]$shape.Range.Start -lt [int]$Range.End -and [int]$shape.Range.End -gt [int]$Range.Start) { return $true }
    }
    foreach ($shape in @($Document.Shapes)) {
        if ([int]$shape.Anchor.Start -ge [int]$Range.Start -and [int]$shape.Anchor.Start -lt [int]$Range.End) { return $true }
    }
    foreach ($table in @($Document.Tables)) {
        if ([int]$table.Range.Start -lt [int]$Range.End -and [int]$table.Range.End -gt [int]$Range.Start) { return $true }
    }
    return $false
}

function Test-FirstPageCover($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    # Existing covers may say "说明书" or another legacy manual name. They
    # are still valid supplied covers and must not cause a duplicate insertion.
    return ($null -ne (Find-CoverBlock $items $Title 1 1 $true) -or $null -ne (Find-LeadingTableCover $Document $items $Title))
}

function Ensure-FirstPageCover($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $block = Find-CoverBlock $items $Title 1 1 $true
    if ($null -ne $block) { return [pscustomobject]@{ Changed = $false; Result = "EXISTING_COVER" } }
    if ($null -ne (Find-LeadingTableCover $Document $items $Title)) {
        return [pscustomobject]@{ Changed = $false; Result = "EXISTING_COVER" }
    }

    $firstText = $items | Where-Object { $_.Page -eq 1 -and $_.Text -ne "" } | Select-Object -First 1
    $titleItem = $items | Where-Object { $_.Page -eq 1 -and (Test-CoverTitle $_.Text $Title) } | Select-Object -First 1
    if ($null -ne $titleItem -and $null -ne $firstText -and $firstText.Index -eq $titleItem.Index -and -not $titleItem.InTable) {
        # A full name at the start of page one is a confirmed cover fragment.
        # Add only the missing label after its paragraph, leaving later content
        # in place so Ensure-CoverPageBoundary can move the body safely.
        $position = [int]$titleItem.End
        $Document.Range($position, $position).InsertBefore("使用说明书`r")
        $Document.Repaginate()
        return [pscustomobject]@{ Changed = $true; Result = "LABEL_ADDED" }
    }

    Add-Cover $Document $Title
    return [pscustomobject]@{ Changed = $true; Result = "COVER_ADDED" }
}

function Get-WordPageAudit($Document) {
    $Document.Repaginate()
    $pageCount = [int]$Document.ComputeStatistics($wdStatisticPages)
    $pages = @{}
    for ($page = 1; $page -le $pageCount; $page++) {
        $pages[$page] = [ordered]@{
            page = $page; text = @(); inline_images = 0; floating_shapes = 0; tables = 0
            page_breaks = 0; section_breaks = 0; page_break_before = 0; section_starts = @()
        }
    }
    for ($index = 1; $index -le $Document.Paragraphs.Count; $index++) {
        try {
            $paragraph = $Document.Paragraphs.Item($index)
            $range = $paragraph.Range
            $page = [int]$range.Information($wdActiveEndPageNumber)
            if (-not $pages.ContainsKey($page)) { continue }
            $text = Clean-Text $range.Text
            if ($text -ne "") { $pages[$page].text += $text }
            if ([string]$range.Text -match '\x0C') { $pages[$page].page_breaks++ }
            try { if ([bool]$paragraph.Range.ParagraphFormat.PageBreakBefore) { $pages[$page].page_break_before++ } } catch {}
            try { if ([bool]$range.Information($wdWithInTable)) { $pages[$page].tables++ } } catch {}
        } catch {}
    }
    foreach ($shape in @($Document.InlineShapes)) {
        try { $page = [int]$shape.Range.Information($wdActiveEndPageNumber); if ($pages.ContainsKey($page)) { $pages[$page].inline_images++ } } catch {}
    }
    foreach ($shape in @($Document.Shapes)) {
        try { $page = [int]$shape.Anchor.Information($wdActiveEndPageNumber); if ($pages.ContainsKey($page)) { $pages[$page].floating_shapes++ } } catch {}
    }
    foreach ($section in @($Document.Sections)) {
        try {
            $sectionStart = [int]$section.Range.Start
            $page = [int]$Document.Range($sectionStart, [Math]::Min($sectionStart + 1, [int]$Document.Content.End)).Information($wdActiveEndPageNumber)
            if ($pages.ContainsKey($page)) {
                $pages[$page].section_breaks++
                $kind = [int]$section.PageSetup.SectionStart
                $startKind = switch ($kind) { 3 { "even_page" } 4 { "odd_page" } 2 { "new_page" } default { "continuous_or_column" } }
                $pages[$page].section_starts += $startKind
            }
        } catch {}
    }
    return @($pages.Values | Sort-Object -Property @{ Expression = { [int]$_.page }; Ascending = $true } | ForEach-Object {
        [pscustomobject]@{
            page = $_.page; visible_text = @($_.text); inline_images = $_.inline_images
            floating_shapes = $_.floating_shapes; tables = $_.tables; page_breaks = $_.page_breaks
            section_breaks = $_.section_breaks; page_break_before = $_.page_break_before
            section_starts = @($_.section_starts)
            has_body_content = (@($_.text).Count -gt 0 -or $_.inline_images -gt 0 -or $_.floating_shapes -gt 0 -or $_.tables -gt 0)
        }
    })
}

function Repair-TerminalBlankParagraph($Document, $PageAudit) {
    $auditPages = @($PageAudit | Sort-Object -Property @{ Expression = { [int]$_.page }; Ascending = $true })
    $pageCount = if ($auditPages.Count -gt 0) {
        [int](($auditPages | Measure-Object -Property page -Maximum).Maximum)
    } else { 0 }
    if ($auditPages.Count -lt 2) {
        return [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = $null }
    }
    $terminalAuditPage = $auditPages | Where-Object { [int]$_.page -eq $pageCount } | Select-Object -First 1
    if ($pageCount -lt 2 -or $null -eq $terminalAuditPage -or $terminalAuditPage.has_body_content) {
        return [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = $null }
    }

    $targetPage = $pageCount
    $diagnostic = [ordered]@{
        page = $targetPage; cause = "末尾空白页终段落格式检查"; outcome = "BLOCKED"
        reason = "未满足严格修复条件"; paragraph_text = ""; previous_page = 0
        pages_before = $pageCount; pages_after = $pageCount
        pages_after_attempt = $null; terminal_page_after_attempt = $null
        terminal_text_codes_after_attempt = @()
        restore_verified = $false; restore_failures = @()
        page_count_after_restore = $null; terminal_page_after_restore = $null
        font_size_before = $null; line_spacing_rule_before = $null; line_spacing_before = $null
        space_before_before = $null; space_after_before = $null
    }
    $reason = ""
    $paragraph = $null
    $range = $null
    $previous = $null
    $paragraphCount = 0
    try { $paragraphCount = [int]$Document.Paragraphs.Count } catch { $reason = "无法读取段落集合" }
    if (-not $reason -and $paragraphCount -lt 2) { $reason = "缺少前一段正文" }
    if (-not $reason) {
        try {
            $paragraph = $Document.Paragraphs.Item($paragraphCount)
            $previous = $Document.Paragraphs.Item($paragraphCount - 1)
            $range = $paragraph.Range
            $rawText = [string]$range.Text
            $diagnostic.paragraph_text = $rawText
            if ($rawText -cne "`r") { $reason = "末尾段落文本不是单独回车符" }
        } catch { $reason = "无法读取末尾段落" }
    }
    if (-not $reason) {
        try {
            if ([bool]$range.Information($wdWithInTable)) { $reason = "末尾段落位于表格内" }
            elseif ([bool]$range.ParagraphFormat.PageBreakBefore) { $reason = "末尾段落设置了段前分页" }
            elseif ([bool]$range.ParagraphFormat.KeepWithNext -or [bool]$range.ParagraphFormat.KeepTogether) { $reason = "末尾段落带有 keep 分页约束" }
            elseif (Test-RangeHasProtectedContent $Document $range $true $false $true) { $reason = "末尾段落包含控制符、对象或受保护内容" }
        } catch { $reason = "无法核验末尾段落的保护属性" }
    }
    if (-not $reason) {
        try {
            $sectionCount = [int]$Document.Sections.Count
            $lastSection = $Document.Sections.Item($sectionCount)
            $lastSectionStart = [int]$lastSection.Range.Start
            if ($lastSectionStart -gt [int]$range.Start -and $lastSectionStart -le [int]$range.End) {
                $reason = "末尾段落包含分节起始边界"
            }
            for ($sectionIndex = 1; -not $reason -and $sectionIndex -lt $sectionCount; $sectionIndex++) {
                $sectionEnd = [int]$Document.Sections.Item($sectionIndex).Range.End
                if ($sectionEnd -gt [int]$range.Start -and $sectionEnd -le [int]$range.End) {
                    $reason = "末尾段落包含分节结束边界"
                }
            }
            if (-not $reason -and $sectionCount -gt 1) {
                $paragraphSection = [int]$range.Information($wdActiveEndSectionNumber)
                $previousSection = [int]$previous.Range.Information($wdActiveEndSectionNumber)
                if ($paragraphSection -ne $previousSection) { $reason = "末尾段落与前一段分属不同节" }
            }
        } catch { $reason = "无法核验末尾段落的分节边界" }
    }
    if (-not $reason) {
        try {
            $previousText = Clean-Text ([string]$previous.Range.Text)
            $previousPage = [int]$previous.Range.Information($wdActiveEndPageNumber)
            $terminalPage = [int]$range.Information($wdActiveEndPageNumber)
            $diagnostic.previous_page = $previousPage
            if ([string]::IsNullOrWhiteSpace($previousText)) { $reason = "前一段没有可见正文" }
            elseif (-not [bool]@($auditPages | Where-Object { [int]$_.page -eq ($targetPage - 1) -and $_.has_body_content })) { $reason = "前一页没有正文内容" }
            elseif ($previousPage -ne ($targetPage - 1)) { $reason = "前一段不在紧邻的前一页" }
            elseif ($terminalPage -ne $targetPage) { $reason = "末尾空段未独占最后一页" }
        } catch { $reason = "无法核验末尾段落与前一页正文的位置" }
    }

    $format = $null
    $oldFormat = $null
    if (-not $reason) {
        try {
            $format = $range.ParagraphFormat
            $oldFormat = [ordered]@{
                font_size = [double]$range.Font.Size
                line_spacing_rule = [int]$format.LineSpacingRule
                line_spacing = [double]$format.LineSpacing
                space_before = [double]$format.SpaceBefore
                space_after = [double]$format.SpaceAfter
            }
            if ($oldFormat.font_size -le 0 -or $oldFormat.font_size -ge 999999 -or
                $oldFormat.line_spacing_rule -lt 0 -or $oldFormat.line_spacing_rule -gt 5 -or
                $oldFormat.line_spacing -le 0 -or $oldFormat.line_spacing -ge 999999 -or
                [Math]::Abs($oldFormat.space_before) -ge 999999 -or [Math]::Abs($oldFormat.space_after) -ge 999999) {
                $reason = "末尾段落格式不是可恢复的单一值"
            }
            foreach ($key in $oldFormat.Keys) { $diagnostic[($key + "_before")] = $oldFormat[$key] }
        } catch { $reason = "无法读取末尾段落的原始格式" }
    }
    if ($reason) {
        $diagnostic.reason = $reason
        return [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = [pscustomobject]$diagnostic }
    }

    try {
        $range.Font.Size = [double]1
        $format.LineSpacingRule = 4
        $format.LineSpacing = [double]1
        $format.SpaceBefore = [double]0
        $format.SpaceAfter = [double]0
        $Document.Repaginate()
        $pagesAfter = [int]$Document.ComputeStatistics($wdStatisticPages)
        $terminalAfter = $Document.Paragraphs.Item($paragraphCount).Range
        $terminalPageAfter = [int]$terminalAfter.Information($wdActiveEndPageNumber)
        $pageCountMatches = ($pagesAfter -eq ($pageCount - 1))
        $terminalPageMatches = ($terminalPageAfter -eq ($targetPage - 1))
        $terminalTextMatches = ([string]$terminalAfter.Text -ceq [string][char]13)
        $diagnostic.pages_after_attempt = $pagesAfter
        $diagnostic.terminal_page_after_attempt = $terminalPageAfter
        $diagnostic.terminal_text_matches_after_attempt = $terminalTextMatches
        $diagnostic.terminal_text_codes_after_attempt = @(([string]$terminalAfter.Text).ToCharArray() | ForEach-Object { [int][char]$_ })
        if ($pageCountMatches -and $terminalPageMatches -and $terminalTextMatches) {
            $diagnostic.outcome = "REPAIRED"
            $diagnostic.reason = "末尾空段已收缩并回到前一页"
            $diagnostic.pages_after = $pagesAfter
            $change = [pscustomobject]@{
                page = $targetPage; cause = "末尾空白页终段落格式收缩"
                action = "保留末尾空段并收缩格式至前一页"
                before_pages = $pageCount; after_pages = $pagesAfter
                previous_page = [int]$diagnostic.previous_page
                font_size_before = [double]$oldFormat.font_size
            }
            return [pscustomobject]@{ Changed = $true; Change = $change; Diagnostic = [pscustomobject]$diagnostic }
        }
        $reason = "重新分页后总页数未恰减一，或末尾空段未回到前一页"
    } catch {
        $reason = "收缩末尾空段后重新分页失败：" + $_.Exception.Message
    }

    try {
        $range.Font.Size = [double]$oldFormat.font_size
        $format.LineSpacingRule = [int]$oldFormat.line_spacing_rule
        $format.LineSpacing = [double]$oldFormat.line_spacing
        $format.SpaceBefore = [double]$oldFormat.space_before
        $format.SpaceAfter = [double]$oldFormat.space_after
        $Document.Repaginate()
        $pagesAfterRestore = [int]$Document.ComputeStatistics($wdStatisticPages)
        $terminalAfterRestore = $Document.Paragraphs.Item($paragraphCount).Range
        $terminalPageAfterRestore = [int]$terminalAfterRestore.Information($wdActiveEndPageNumber)
        $restoredFormat = [ordered]@{
            font_size = [double]$terminalAfterRestore.Font.Size
            line_spacing_rule = [int]$terminalAfterRestore.ParagraphFormat.LineSpacingRule
            line_spacing = [double]$terminalAfterRestore.ParagraphFormat.LineSpacing
            space_before = [double]$terminalAfterRestore.ParagraphFormat.SpaceBefore
            space_after = [double]$terminalAfterRestore.ParagraphFormat.SpaceAfter
        }
        $diagnostic.pages_after = $pagesAfterRestore
        $diagnostic.page_count_after_restore = $pagesAfterRestore
        $diagnostic.terminal_page_after_restore = $terminalPageAfterRestore
        if ($pagesAfterRestore -ne $pageCount) { $diagnostic.restore_failures += "恢复后页数为 $pagesAfterRestore，应为 $pageCount" }
        if ($terminalPageAfterRestore -ne $targetPage) { $diagnostic.restore_failures += "恢复后末段位于第 $terminalPageAfterRestore 页，应回到第 $targetPage 页" }
        if ([string]$terminalAfterRestore.Text -cne [string]$diagnostic.paragraph_text) { $diagnostic.restore_failures += "恢复后末段文本与修复前不同" }
        if ([Math]::Abs($restoredFormat.font_size - $oldFormat.font_size) -gt 0.01) { $diagnostic.restore_failures += "恢复后字体大小与原值不同" }
        if ($restoredFormat.line_spacing_rule -ne $oldFormat.line_spacing_rule) { $diagnostic.restore_failures += "恢复后行距规则与原值不同" }
        if ([Math]::Abs($restoredFormat.line_spacing - $oldFormat.line_spacing) -gt 0.01) { $diagnostic.restore_failures += "恢复后行距与原值不同" }
        if ([Math]::Abs($restoredFormat.space_before - $oldFormat.space_before) -gt 0.01) { $diagnostic.restore_failures += "恢复后段前距与原值不同" }
        if ([Math]::Abs($restoredFormat.space_after - $oldFormat.space_after) -gt 0.01) { $diagnostic.restore_failures += "恢复后段后距与原值不同" }
    } catch {
        throw ("末尾空段修复失败且无法完成原状态恢复核验：" + $_.Exception.Message)
    }
    if ($diagnostic.restore_failures.Count -gt 0) {
        throw ("末尾空段修复失败且原状态恢复核验未通过：" + ($diagnostic.restore_failures -join "；"))
    }
    $diagnostic.restore_verified = $true
    $diagnostic.outcome = "RESTORED"
    $diagnostic.reason = $reason
    return [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = [pscustomobject]$diagnostic }
}

function Repair-WordBlankPages($Document, [int[]]$TargetPages = @()) {
    $changes = @()
    $diagnostics = @()
    $pageAudit = @(Get-WordPageAudit $Document)
    $auditCount = 1
    $auditDirty = $false
    $maxRounds = [Math]::Max(1, $pageAudit.Count)
    for ($round = 0; $round -lt $maxRounds; $round++) {
        if ($auditDirty) {
            $pageAudit = @(Get-WordPageAudit $Document)
            $auditCount++
            $auditDirty = $false
        }
        $emptyPages = @($pageAudit | Where-Object { -not $_.has_body_content } | ForEach-Object { [int]$_.page })
        if ($TargetPages.Count -gt 0) { $emptyPages = @($emptyPages | Where-Object { $_ -in $TargetPages }) }
        if ($emptyPages.Count -eq 0) { break }
        $changed = $false
        foreach ($emptyPage in $emptyPages) {
            $nextPage = $emptyPage + 1
            foreach ($section in @($Document.Sections)) {
                try {
                    $sectionStart = [int]$section.Range.Start
                    $startPage = [int]$Document.Range($sectionStart, [Math]::Min($sectionStart + 1, [int]$Document.Content.End)).Information($wdActiveEndPageNumber)
                    $sectionKind = [int]$section.PageSetup.SectionStart
                    if ($startPage -eq $nextPage -and ($sectionKind -eq 3 -or $sectionKind -eq 4)) {
                        $auditDirty = $true
                        $section.PageSetup.SectionStart = 2
                        $changes += [pscustomobject]@{ page = $emptyPage; cause = if ($sectionKind -eq 4) { "奇数页分节" } else { "偶数页分节" }; action = "改为下一页分节" }
                        $changed = $true
                        break
                    }
                } catch {}
            }
        }
        if (-not $changed) {
            foreach ($emptyPage in $emptyPages) {
                $blankPageAudit = $pageAudit | Where-Object { [int]$_.page -eq $emptyPage } | Select-Object -First 1
                $sectionStarts = @($blankPageAudit.section_starts)
                if ($null -eq $blankPageAudit -or $blankPageAudit.has_body_content -or
                    [int]$blankPageAudit.page_breaks -ne 1 -or [int]$blankPageAudit.section_breaks -ne 1 -or
                    [int]$blankPageAudit.page_break_before -ne 0 -or $sectionStarts.Count -ne 1 -or
                    $sectionStarts[0] -ne "new_page" -or @($blankPageAudit.visible_text).Count -ne 0 -or
                    [int]$blankPageAudit.tables -ne 0 -or [int]$blankPageAudit.inline_images -ne 0 -or
                    [int]$blankPageAudit.floating_shapes -ne 0) { continue }

                $sectionStartsOnPage = @()
                foreach ($section in @($Document.Sections)) {
                    try {
                        if ([int]$section.PageSetup.SectionStart -ne 2) { continue }
                        $sectionStart = [int]$section.Range.Start
                        $sectionStartPage = [int]$Document.Range(
                            $sectionStart,
                            [Math]::Min($sectionStart + 1, [int]$Document.Content.End)
                        ).Information($wdActiveEndPageNumber)
                        if ($sectionStartPage -eq $emptyPage) {
                            $sectionStartsOnPage += [pscustomobject]@{ Start = $sectionStart; Section = $section }
                        }
                    } catch {}
                }
                if ($sectionStartsOnPage.Count -ne 1) { continue }
                $sectionStart = [int]$sectionStartsOnPage[0].Start
                $eligibleBreaks = @()
                for ($index = 1; $index -le $Document.Paragraphs.Count; $index++) {
                    try {
                        $paragraph = $Document.Paragraphs.Item($index)
                        $paragraphRange = $paragraph.Range
                        if ([int]$paragraphRange.Start -ne $sectionStart -or
                            [int]$paragraphRange.End -ne ($sectionStart + 2) -or
                            [int]$paragraphRange.Information($wdActiveEndPageNumber) -ne $emptyPage -or
                            [bool]$paragraphRange.Information($wdWithInTable) -or
                            [bool]$paragraphRange.ParagraphFormat.PageBreakBefore) { continue }

                        $paragraphCodes = @(([string]$paragraphRange.Text).ToCharArray() | ForEach-Object { [int][char]$_ })
                        if ($paragraphCodes.Count -ne 2 -or $paragraphCodes[0] -ne 12 -or $paragraphCodes[1] -ne 13 -or
                            (Test-RangeHasProtectedContent $Document $paragraphRange $true $true $true)) { continue }

                        $breakRange = $Document.Range($sectionStart, $sectionStart + 1)
                        if ([int]$breakRange.Start -ne $sectionStart -or [int]$breakRange.End -ne ($sectionStart + 1) -or
                            [string]$breakRange.Text -cne [string][char]12 -or
                            [int]$breakRange.Information($wdActiveEndPageNumber) -ne $emptyPage -or
                            [bool]$breakRange.Information($wdWithInTable)) { continue }

                        $eligibleBreaks += [pscustomobject]@{ Range = $breakRange }
                    } catch {}
                }
                if ($eligibleBreaks.Count -ne 1) { continue }

                # Remove only the duplicate form-feed at the exact start of a
                # next-page section. Keep its paragraph mark and section break.
                $auditDirty = $true
                $eligibleBreaks[0].Range.Delete() | Out-Null
                $changes += [pscustomobject]@{
                    page = $emptyPage
                    cause = "下一页分节起始页重复显式分页符"
                    action = "删除分节起点分页符字符并保留段落标记和分节"
                }
                $changed = $true
                break
            }
        }
        if (-not $changed) { break }
        $Document.Repaginate()
    }
    if ($auditDirty) {
        $pageAudit = @(Get-WordPageAudit $Document)
        $auditCount++
    }
    $beforeTerminalAudit = $pageAudit
    $terminalPage = if ($beforeTerminalAudit.Count -gt 0) { [int]$beforeTerminalAudit[-1].page } else { 0 }
    $terminalRepair = if ($TargetPages.Count -eq 0 -or $terminalPage -in $TargetPages) {
        Repair-TerminalBlankParagraph $Document $beforeTerminalAudit
    } else {
        [pscustomobject]@{ Changed = $false; Change = $null; Diagnostic = $null }
    }
    if ($null -ne $terminalRepair.Diagnostic) { $diagnostics += $terminalRepair.Diagnostic }
    if ($terminalRepair.Changed) { $changes += $terminalRepair.Change }
    # A failed compaction still mutates and restores Word layout. Re-audit it.
    $terminalMutated = $terminalRepair.Changed -or ($null -ne $terminalRepair.Diagnostic -and
        $terminalRepair.Diagnostic.outcome -in @('REPAIRED', 'RESTORED'))
    $finalPages = if ($terminalMutated) {
        $auditCount++
        @(Get-WordPageAudit $Document)
    } else { $beforeTerminalAudit }
    $unresolvedPages = @($finalPages | Where-Object { -not $_.has_body_content } | ForEach-Object {
        [pscustomobject]@{ page = [int]$_.page; cause = "无法安全识别或修复空白页成因" }
    })
    return [pscustomobject]@{ Changed = ($changes.Count -gt 0); Changes = $changes; Pages = $finalPages; UnresolvedPages = $unresolvedPages; Diagnostics = $diagnostics; AuditCount = $auditCount }
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
    if ($otherText.Count -eq 0 -and -not (Test-PageHasObjects $Document 1)) {
        $pageStart = [int]$Document.GoTo(1, 1, 1).Start
        $nextPage = $Document.GoTo(1, 1, 2)
        $pageEnd = if ($null -ne $nextPage -and [int]$nextPage.Start -gt $pageStart) { [int]$nextPage.Start } else { [int]$Document.Content.End }
        if (Test-RangeHasProtectedContent $Document ($Document.Range($pageStart, $pageEnd))) { return $false }
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
    $items = @(Get-FrontParagraphs $Document 1)
    $firstItem = $items | Where-Object { $_.Page -eq 1 } | Select-Object -First 1
    if ($null -ne $firstItem -and $firstItem.InTable -and $null -eq (Find-LeadingTableCover $Document $items $Title)) {
        throw "文档以普通表格开头，无法安全插入封面，请人工处理"
    }
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
    Set-CoverVerticalAlignment $Document 1
}

function Set-CoverVerticalAlignment($Document, [int]$CoverSection) {
    $sectionCount = [int]$Document.Sections.Count
    # Default every section to top alignment. Only a confirmed first section
    # separated from the body may use vertical centering for its cover page.
    for ($sectionIndex = 1; $sectionIndex -le $sectionCount; $sectionIndex++) {
        $Document.Sections.Item($sectionIndex).PageSetup.VerticalAlignment = $wdVerticalTop
    }
    if ($sectionCount -gt 1 -and $CoverSection -eq 1) {
        $Document.Sections.Item(1).PageSetup.VerticalAlignment = $wdVerticalCenter
    }
}

function Remove-DuplicateCoverBlocks($Document, [string]$Title) {
    $removed = 0
    $minimumPage = 2
    for ($round = 0; $round -lt 8; $round++) {
        # Cover title and label cleanup is intentionally capped at the first
        # five actual Word pages. An unsafe page is skipped while later pages
        # continue to be checked.
        $items = @(Get-FrontParagraphs $Document 5)
        $exactTitle = $null
        $exactTitleRange = $null
        foreach ($item in $items) {
            if ($item.Page -lt 2 -or $item.Page -gt 5 -or $item.InTable -or
                $item.Style -match '(?i)TOC|目录' -or
                (Compact-Text $item.Text) -cne (Compact-Text $Title)) { continue }
            # Paragraph marks may carry section breaks. Inspect only the visible
            # title text here so the cleanup can remove that text while keeping
            # the structural boundary intact.
            $candidateTextEnd = [Math]::Max([int]$item.Start, [int]$item.End - 1)
            $candidateRange = $Document.Range([int]$item.Start, $candidateTextEnd)
            if ([bool]$candidateRange.Information($wdWithInTable) -or
                (Test-RangeHasProtectedContent $Document $candidateRange $true $false $true)) { continue }
            $exactTitle = $item
            $exactTitleRange = $candidateRange
            break
        }
        if ($null -ne $exactTitle) {
            # Keep the paragraph mark: Word may store a section break on it.
            # Removing only the visible title lets the blank-page repair handle
            # the now-empty page without deleting the section boundary.
            $titleTextEnd = [Math]::Max([int]$exactTitle.Start, [int]$exactTitle.End - 1)
            $Document.Range([int]$exactTitle.Start, $titleTextEnd).Delete() | Out-Null
            $removed++
            $minimumPage = 2
            continue
        }

        # A duplicate title/label can share page 2 with the first body text or
        # contents. Remove only leading, exact-match front-matter paragraphs.
        $pageTwoText = @($items | Where-Object { $_.Page -eq 2 -and $_.Text -ne "" })
        $leadingFragment = if ($pageTwoText.Count -gt 0) { $pageTwoText[0] } else { $null }
        if ($null -ne $leadingFragment -and $leadingFragment.Page -eq 2 -and -not $leadingFragment.InTable -and
            $leadingFragment.Style -notmatch '(?i)TOC|目录' -and
            (Test-PostCoverFragment $leadingFragment.Text $Title)) {
            $bodyText = @($pageTwoText | Where-Object { -not (Test-PostCoverFragment $_.Text $Title) })
            $fragmentRange = $Document.Range($leadingFragment.Start, $leadingFragment.End)
            if ($bodyText.Count -gt 0 -and -not [bool]$fragmentRange.Information($wdWithInTable) -and
                -not (Test-RangeHasProtectedContent $Document $fragmentRange $true)) {
                $Document.Range($leadingFragment.Start, $leadingFragment.End).Delete() | Out-Null
                $removed++
                $minimumPage = 2
                continue
            }
        }
        $block = Find-CoverBlock $items $Title $minimumPage 5 $true
        $isolated = $null
        if ($null -eq $block) { $isolated = Find-IsolatedTitleCover $items $Title $minimumPage 5 }
        if ($null -eq $block -and $null -eq $isolated) { $isolated = Find-IsolatedManualLabelCover $items $minimumPage 5 }
        if ($null -eq $block -and $null -eq $isolated) { $isolated = Find-IsolatedCombinedCoverFragment $items $Title $minimumPage 5 }
        $residualLabel = $null
        if ($null -eq $block -and $null -eq $isolated) { $residualLabel = Find-ResidualManualLabel $items $minimumPage 5 }
        if ($null -eq $block -and $null -eq $isolated -and $null -eq $residualLabel) { break }

        if ($null -ne $residualLabel) {
            $otherText = @($items | Where-Object {
                $_.Page -eq $residualLabel.Page -and $_.Index -ne $residualLabel.Index -and $_.Text -ne ""
            })
            $hasObjects = $false
            foreach ($shape in @($Document.InlineShapes)) {
                if ([int]$shape.Range.Information($wdActiveEndPageNumber) -eq $residualLabel.Page) { $hasObjects = $true; break }
            }
            if (-not $hasObjects) {
                foreach ($table in @($Document.Tables)) {
                    if ([int]$table.Range.Information($wdActiveEndPageNumber) -eq $residualLabel.Page) { $hasObjects = $true; break }
                }
            }
            # A residual label on a content page is ambiguous. Keep it intact
            # and let the final QA report it instead of deleting supplied text.
            if ($otherText.Count -gt 0 -or $hasObjects) { $minimumPage = [Math]::Max($minimumPage, $residualLabel.Page + 1); continue }
            # Preserve a section break that Word may store on the paragraph mark.
            $labelTextEnd = [Math]::Max([int]$residualLabel.Start, [int]$residualLabel.End - 1)
            $Document.Range([int]$residualLabel.Start, $labelTextEnd).Delete() | Out-Null
            $removed++
            $minimumPage = 2
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
            if ($hasObjects -or (Test-PageHasObjects $Document $isolated.Page)) { $minimumPage = [Math]::Max($minimumPage, $isolated.Page + 1); continue }
            $pageStart = [int]$Document.GoTo(1, 1, $isolated.Page).Start
            $nextPage = $Document.GoTo(1, 1, $isolated.Page + 1)
            $pageEnd = if ($null -ne $nextPage -and [int]$nextPage.Start -gt $pageStart) { [int]$nextPage.Start } else { [int]$Document.Content.End }
            if (Test-RangeHasProtectedContent $Document ($Document.Range($pageStart, $pageEnd)) $false $true $true) { $minimumPage = [Math]::Max($minimumPage, $isolated.Page + 1); continue }
            $Document.Range($pageStart, $pageEnd).Delete() | Out-Null
            $removed++
            $minimumPage = 2
            continue
        }

        $safe = $true
        for ($i = $block.TitleIndex + 1; $i -lt $block.LabelIndex; $i++) {
            if ($items[$i].Text -ne "") { $safe = $false; break }
        }
        if (-not $safe) { $minimumPage = [Math]::Max($minimumPage, $block.Page + 1); continue }

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
        if ($otherText.Count -eq 0 -and -not $hasObjects -and -not (Test-PageHasObjects $Document $block.Page)) {
            $pageStart = [int]$Document.GoTo(1, 1, $block.Page).Start
            $nextPage = $Document.GoTo(1, 1, $block.Page + 1)
            $pageEnd = if ($null -ne $nextPage -and [int]$nextPage.Start -gt $pageStart) { [int]$nextPage.Start } else { [int]$Document.Content.End }
            if (Test-RangeHasProtectedContent $Document ($Document.Range($pageStart, $pageEnd)) $false $true $true) { $minimumPage = [Math]::Max($minimumPage, $block.Page + 1); continue }
            $Document.Range($pageStart, $pageEnd).Delete() | Out-Null
        } else {
            # This is not a standalone duplicate cover.  Do not delete a
            # title/label block from a page that also contains material.
            $minimumPage = [Math]::Max($minimumPage, $block.Page + 1)
            continue
        }
        $removed++
        $minimumPage = 2
        # Delete only the duplicate text block. Never delete the whole page.
    }
    $residualTitles = @(
        Get-FrontParagraphs $Document 5 |
            Where-Object { $_.Page -ge 2 -and $_.Page -le 5 -and (Compact-Text $_.Text) -ceq (Compact-Text $Title) } |
            ForEach-Object { "p$($_.Page)@$($_.Start)-$($_.End)" }
    )
    Trace-Step $Title ("duplicate-residuals:" + ($residualTitles -join ","))
    return $removed
}

function Test-DuplicateCoverBlock($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 5)
    foreach ($item in $items) {
        if ($item.Page -ge 2 -and $item.Page -le 5 -and -not $item.InTable -and
            $item.Style -notmatch '(?i)TOC|目录' -and
            (Compact-Text $item.Text) -ceq (Compact-Text $Title)) { return $true }
    }
    $pageTwoText = @($items | Where-Object { $_.Page -eq 2 -and $_.Text -ne "" })
    if ($pageTwoText.Count -gt 0 -and -not $pageTwoText[0].InTable -and
        $pageTwoText[0].Style -notmatch '(?i)TOC|目录' -and
        (Test-PostCoverFragment $pageTwoText[0].Text $Title)) {
        $bodyText = @($pageTwoText | Where-Object { -not (Test-PostCoverFragment $_.Text $Title) })
        $firstRange = $Document.Range($pageTwoText[0].Start, $pageTwoText[0].End)
        if ($bodyText.Count -gt 0 -and -not [bool]$firstRange.Information($wdWithInTable) -and
            -not (Test-RangeHasProtectedContent $Document $firstRange $true)) { return $true }
    }
    $block = Find-CoverBlock $items $Title 2 5 $true
    if ($null -ne $block) {
        $titleIndex = $items[$block.TitleIndex].Index
        $labelIndex = $items[$block.LabelIndex].Index
        $otherText = @($items | Where-Object { $_.Page -eq $block.Page -and $_.Text -ne "" -and $_.Index -ne $titleIndex -and $_.Index -ne $labelIndex })
        if ($otherText.Count -eq 0 -and -not (Test-PageHasObjects $Document $block.Page)) { return $true }
    }
    foreach ($isolated in @(
        (Find-IsolatedTitleCover $items $Title 2 5),
        (Find-IsolatedManualLabelCover $items 2 5),
        (Find-IsolatedCombinedCoverFragment $items $Title 2 5)
    )) {
        if ($null -ne $isolated -and -not (Test-PageHasObjects $Document $isolated.Page)) { return $true }
    }
    $residual = Find-ResidualManualLabel $items 2 5
    if ($null -ne $residual) {
        $otherText = @($items | Where-Object { $_.Page -eq $residual.Page -and $_.Index -ne $residual.Index -and $_.Text -ne "" })
        if ($otherText.Count -eq 0 -and -not (Test-PageHasObjects $Document $residual.Page)) { return $true }
    }
    return $false
}

function Get-CoverLastLineUnitCount($Range, $Workload = $null) {
    $characters = $Range.Characters
    $characterCount = [int]$characters.Count
    $lastLineText = ""
    $fastPath = $false
    try {
        if ($characterCount -eq 0) { return 0 }
        $first = if ($characters -is [array]) { $characters[0] } else { $characters.Item(1) }
        $last = if ($characters -is [array]) { $characters[-1] } else { $characters.Item($characterCount) }
        if ($null -ne $Workload) { $Workload.page_information_calls++ }
        $firstPage = [int]$first.Information($wdActiveEndPageNumber)
        if ($null -ne $Workload) { $Workload.page_information_calls++ }
        $lastPage = [int]$last.Information($wdActiveEndPageNumber)
        # Word line numbers restart on a new page. Only a single-page range
        # has the monotonic line numbers needed to stop the reverse scan.
        if ($firstPage -gt 0 -and $firstPage -eq $lastPage) {
            $lastLine = -1
            $parts = New-Object 'System.Collections.Generic.List[string]'
            for ($index = $characterCount; $index -ge 1; $index--) {
                $character = if ($characters -is [array]) { $characters[$index - 1] } else { $characters.Item($index) }
                $text = Clean-Text ([string]$character.Text)
                if ($text -eq "") { continue }
                if ($null -ne $Workload) { $Workload.line_information_calls++ }
                $line = [int]$character.Information($wdFirstCharacterLineNumber)
                if ($line -le 0) { throw 'Word line information unavailable' }
                if ($lastLine -lt 0) { $lastLine = $line }
                if ($line -lt $lastLine) { break }
                if ($line -gt $lastLine) { throw 'Word line information is not monotonic' }
                $parts.Add($text)
            }
            $parts.Reverse()
            $lastLineText = $parts -join ''
            $fastPath = $true
        }
    } catch { $fastPath = $false }
    if (-not $fastPath) {
        if ($null -ne $Workload) { $Workload.full_scan_fallbacks++ }
        $lastLine = -1
        $lastLineText = ""
        foreach ($character in @($Range.Characters)) {
            $text = Clean-Text ([string]$character.Text)
            if ($text -eq "") { continue }
            try {
                if ($null -ne $Workload) { $Workload.line_information_calls++ }
                $line = [int]$character.Information($wdFirstCharacterLineNumber)
            } catch { continue }
            if ($line -gt $lastLine) {
                $lastLine = $line
                $lastLineText = $text
            } elseif ($line -eq $lastLine) {
                $lastLineText += $text
            }
        }
    }
    if ($lastLineText -eq "") { return 0 }
    $units = [regex]::Matches($lastLineText, '[\p{IsCJKUnifiedIdeographs}]|[A-Za-z0-9]+')
    return @($units).Count
}

function Set-CoverTypography($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $titleItem = $items | Where-Object {
        $_.Page -eq 1 -and -not $_.InTable -and $_.Style -notmatch '(?i)TOC|目录' -and (Test-CoverTitle $_.Text $Title)
    } | Select-Object -First 1
    if ($null -eq $titleItem) {
        return [pscustomobject]@{ Found = $false; Size = 0; Lines = 0; Label = ""; LabelLayoutFailed = $false }
    }
    $block = Find-CoverBlock $items $Title 1 1 $true
    $labelItem = if ($null -ne $block) { $items[$block.LabelIndex] } else { $null }

    # Replace stacks of empty paragraphs between the existing title and label
    # with one deterministic spacing value. This keeps supplied wording intact
    # while preventing the label from being pushed onto the lower page or a
    # second page.
    if ($null -ne $labelItem -and $labelItem.Index -gt ($titleItem.Index + 1)) {
        $between = @($items | Where-Object { $_.Index -gt $titleItem.Index -and $_.Index -lt $labelItem.Index })
        if (@($between | Where-Object { $_.Text -ne "" }).Count -eq 0) {
            foreach ($blank in @($between | Sort-Object Index -Descending)) {
                $blankRange = $Document.Range($blank.Start, $blank.End)
                if (-not (Test-RangeHasProtectedContent $Document $blankRange $true)) {
                    $blankRange.Delete() | Out-Null
                }
            }
            $items = @(Get-FrontParagraphs $Document 1)
            $titleItem = $items | Where-Object {
                $_.Page -eq 1 -and -not $_.InTable -and $_.Style -notmatch '(?i)TOC|目录' -and (Test-CoverTitle $_.Text $Title)
            } | Select-Object -First 1
            if ($null -eq $titleItem) { return [pscustomobject]@{ Found = $false; Size = 0; Lines = 0; Label = ""; LabelLayoutFailed = $false } }
            $block = Find-CoverBlock $items $Title 1 1 $true
            $labelItem = if ($null -ne $block) { $items[$block.LabelIndex] } else { $null }
        }
    }

    $titleRange = $Document.Range($titleItem.Start, $titleItem.End)
    $labelRange = $null
    $formatLabelRange = {
        param($Range)
        foreach ($paragraph in $Range.Paragraphs) {
            try { $paragraph.Range.ListFormat.RemoveNumbers($wdNumberAllNumbers) } catch {}
            try { $paragraph.Range.Style = $Document.Styles.Item($wdStyleNormal) } catch {}
            try { $paragraph.Range.ParagraphFormat.OutlineLevel = $wdOutlineLevelBodyText } catch {}
        }
        $Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
        $Range.ParagraphFormat.LeftIndent = [double]0
        $Range.ParagraphFormat.RightIndent = [double]0
        $Range.ParagraphFormat.FirstLineIndent = [double]0
        $Range.Font.Name = "宋体"
        $Range.Font.NameFarEast = "宋体"
        $Range.Font.NameAscii = "SimSun"
        $Range.Font.NameOther = "SimSun"
        $Range.Font.Bold = -1
        $Range.ParagraphFormat.SpaceBefore = [double]48
        $Range.ParagraphFormat.SpaceAfter = [double]0
    }
    # An initially safe title/label block can be formatted before the size loop;
    # each subsequent size pass then needs only one repagination.
    if ($null -ne $labelItem) {
        $labelRange = $Document.Range($labelItem.Start, $labelItem.End)
        & $formatLabelRange $labelRange
    }
    try {
        $coverSection = [int]$titleRange.Information($wdActiveEndSectionNumber)
        Set-CoverVerticalAlignment $Document $coverSection
    } catch {
        # A failed or missing section boundary must not leave body pages centered.
        for ($sectionIndex = 1; $sectionIndex -le $Document.Sections.Count; $sectionIndex++) {
            $Document.Sections.Item($sectionIndex).PageSetup.VerticalAlignment = $wdVerticalTop
        }
    }
    foreach ($paragraph in $titleRange.Paragraphs) {
        try { $paragraph.Range.ListFormat.RemoveNumbers($wdNumberAllNumbers) } catch {}
        try { $paragraph.Range.Style = $Document.Styles.Item($wdStyleNormal) } catch {}
        try { $paragraph.Range.ParagraphFormat.OutlineLevel = $wdOutlineLevelBodyText } catch {}
    }
    $titleRange.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $titleRange.ParagraphFormat.LeftIndent = [double]0
    $titleRange.ParagraphFormat.RightIndent = [double]0
    $titleRange.ParagraphFormat.FirstLineIndent = [double]0
    $titleRange.Font.Name = "宋体"
    $titleRange.Font.NameFarEast = "宋体"
    $titleRange.Font.NameAscii = "SimSun"
    $titleRange.Font.NameOther = "SimSun"
    $titleRange.Font.Bold = -1
    $titleRange.ParagraphFormat.SpaceBefore = [double]0
    $titleRange.ParagraphFormat.SpaceAfter = [double]0

    $workload = [ordered]@{ pagination_attempts = 0; line_information_calls = 0; page_information_calls = 0; full_scan_fallbacks = 0 }
    $size = 26
    $lines = 99
    $orphanedLastLine = $false
    $labelWasSafe = ($null -ne $labelItem)
    $labelLayoutFailed = $false
    while ($size -ge 14) {
        $titleRange.Font.Size = [double]$size
        if ($null -ne $labelRange) { $labelRange.Font.Size = [double]$size }
        $workload.pagination_attempts++
        $Document.Repaginate()

        # A label can initially spill onto page two. Consider formatting it
        # only after this title-size pass has repaginated it into a safe block
        # on page one, with no body paragraph between title and label.
        $items = @(Get-FrontParagraphs $Document 2)
        $block = Find-CoverBlock $items $Title 1 1 $true
        $labelItem = if ($null -ne $block) { $items[$block.LabelIndex] } else { $null }
        if ($null -ne $labelItem) {
            $labelWasSafe = $true
            $labelLayoutFailed = $false
            if ($null -eq $labelRange) {
                $labelRange = $Document.Range($labelItem.Start, $labelItem.End)
                & $formatLabelRange $labelRange
                $labelRange.Font.Size = [double]$size
                # A label that returned from page two is formatted only after
                # the title pass confirmed a safe block. Recheck once because
                # its new spacing/font can change pagination.
                $workload.pagination_attempts++
                $Document.Repaginate()
                $items = @(Get-FrontParagraphs $Document 2)
                $block = Find-CoverBlock $items $Title 1 1 $true
                if ($null -ne $block) {
                    $labelItem = $items[$block.LabelIndex]
                } else {
                    $labelItem = $null
                    $labelLayoutFailed = $true
                }
            }
        } elseif ($labelWasSafe) {
            # Keep the formatted range available for the next smaller-size
            # pass, but never report success while it remains off page one.
            $labelLayoutFailed = $true
        }

        $lines = [int]$titleRange.ComputeStatistics($wdStatisticLines)
        $orphanedLastLine = ($lines -eq 2 -and (Get-CoverLastLineUnitCount $titleRange $workload) -le 2)
        if ($lines -le 2 -and -not $orphanedLastLine -and -not $labelLayoutFailed) { break }
        if ($size -eq 14) { break }
        $size--
    }
    return [pscustomobject]@{ Found = $true; Size = $size; Lines = $lines; OrphanedLastLine = $orphanedLastLine; Label = if ($null -ne $labelItem) { $labelItem.Text } else { "" }; LabelPresent = ($null -ne $labelItem); LabelLayoutFailed = $labelLayoutFailed; Workload = [pscustomobject]$workload }
}

function Remove-CoverVersion($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    if ($null -ne (Find-LeadingTableCover $Document $items $Title)) {
        return [pscustomobject]@{ Found = $true; Changed = $false; ProtectedTable = $true }
    }
    $titleItem = $items | Where-Object { Test-CoverTitle $_.Text $Title } | Select-Object -First 1
    if ($null -eq $titleItem) { return [pscustomobject]@{ Found = $false; Changed = $false; ProtectedTable = $false } }
    # Exclude the trailing paragraph mark so replacing the text keeps the
    # paragraph break and does not merge the title line with the label line.
    $titleRange = $Document.Range($titleItem.Start, [Math]::Max($titleItem.Start, $titleItem.End - 1))
    if ((Compact-Text $titleRange.Text) -eq (Compact-Text $Title)) {
        return [pscustomobject]@{ Found = $true; Changed = $false; ProtectedTable = $false }
    }
    $titleRange.Text = $Title
    return [pscustomobject]@{ Found = $true; Changed = $true; ProtectedTable = $false }
}

function Ensure-CoverPageBoundary($Document, [string]$Title) {
    # If an existing title/label block shares page 1 with body text, a TOC,
    # an image, or a table, insert one page break after the label. Preserve all
    # following content and the document's existing page setup.
    $items = @(Get-FrontParagraphs $Document 1)
    # A table-based cover is already protected as a unit. Inserting a break at
    # a cell paragraph's end would split or otherwise rewrite the source table.
    $firstPageItem = $items | Where-Object { $_.Page -eq 1 } | Select-Object -First 1
    $leadingTableCover = Find-LeadingTableCover $Document $items $Title
    if ($null -ne $firstPageItem -and $firstPageItem.InTable -and $null -eq $leadingTableCover) { return $false }
    if ($null -ne $leadingTableCover) { return $false }
    $block = Find-CoverBlock $items $Title 1 1 $true
    if ($null -eq $block) {
        # Some supplied covers put the title and manual label in separate
        # paragraphs or apply Word list numbering. Find the first title and
        # the nearest manual label on page one without requiring adjacency.
        $titleItem = $items | Where-Object { -not $_.InTable -and (Compact-Text $_.Text) -eq (Compact-Text $Title) } | Select-Object -First 1
        $labelItem = $items | Where-Object { $_.Page -eq 1 -and -not $_.InTable -and (Test-ManualLabel $_.Text) } | Select-Object -First 1
        if ($null -eq $titleItem -or $null -eq $labelItem) { return $false }
        $block = [pscustomobject]@{ LabelIndex = [array]::IndexOf($items, $labelItem) }
    }
    $extra = @($items | Where-Object {
        $_.Page -eq 1 -and $_.Index -gt $items[$block.LabelIndex].Index -and $_.Text -ne ""
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
    if ($extra.Count -eq 0 -and -not $hasObjects) { return $false }
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
        $items = @(Get-FrontParagraphs $Document 5)
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
            if ($hasBreak -or $pageBreakBefore -or (Test-RangeHasProtectedContent $Document ($Document.Range($blank.Start, $blank.End)) $true)) { continue }
            $Document.Range($blank.Start, $blank.End).Delete() | Out-Null
            $deleted++
        }
        if ($deleted -eq 0) { break }
        $removed += $deleted
    }
    return $removed
}

function Remove-RedundantPostCoverFrontMatter($Document, [string]$Title) {
    # Keep the first section/page boundary that separates the cover. Remove only
    # later front-matter fragments and explicit page breaks before real body
    # content; otherwise a supplied cover can consume pages 1-2 while the body
    # starts on page 3.
    $items = @(Get-FrontParagraphs $Document 5)
    $cover = Find-CoverBlock $items $Title 1 1 $true
    if ($null -eq $cover) { return 0 }
    $label = $items[$cover.LabelIndex]
    $coverSection = 1
    try { $coverSection = [int]$Document.Range($label.Start, $label.End).Information($wdActiveEndSectionNumber) } catch {}
    $protectedBoundaryFound = $false
    # In OOXML a section break is commonly stored on the cover label paragraph
    # itself. Word then reports no section-number change on the following empty
    # page-break paragraph, so detect a section end that falls inside the label
    # range before scanning later paragraphs.
    for ($sectionIndex = 1; $sectionIndex -lt $Document.Sections.Count; $sectionIndex++) {
        $sectionEnd = [int]$Document.Sections.Item($sectionIndex).Range.End
        if ($sectionEnd -ge $label.Start -and $sectionEnd -le ($label.End + 1)) {
            $protectedBoundaryFound = $true
            break
        }
    }
    $candidates = @()
    $realBody = $null
    for ($index = $cover.LabelIndex + 1; $index -lt $items.Count; $index++) {
        $item = $items[$index]
        $range = $Document.Range($item.Start, $item.End)
        $section = $coverSection
        try { $section = [int]$range.Information($wdActiveEndSectionNumber) } catch {}
        if (-not $protectedBoundaryFound -and $section -gt $coverSection) {
            $protectedBoundaryFound = $true
            continue
        }
        if (Test-RangeHasProtectedContent $Document $range $true) {
            $realBody = $item
            break
        }
        $isFragment = (Test-CoverTitle $item.Text $Title) -or (Test-LegacyManualLabel $item.Text) -or (Test-CombinedCoverFragment $item.Text $Title)
        if ($item.Text -eq "" -or $isFragment) {
            if ($protectedBoundaryFound) { $candidates += $item }
            continue
        }
        $realBody = $item
        break
    }
    if (-not $protectedBoundaryFound -or $null -eq $realBody -or $realBody.Page -lt 3) { return 0 }
    $removed = 0
    foreach ($item in @($candidates | Sort-Object Index -Descending)) {
        $candidateRange = $Document.Range($item.Start, $item.End)
        if (Test-RangeHasProtectedContent $Document $candidateRange $true) { continue }
        $candidateRange.Delete() | Out-Null
        $removed++
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

function Test-PageLabelParagraph($Paragraph, $Locator) {
    $fields = @(Get-PageFieldsFromRange $Paragraph.Range)
    if ($fields.Count -ne 1) { return $false }
    $text = [string]$Paragraph.Range.Text
    $resultText = [string]$fields[0].Result.Text
    if ($resultText) { $text = $text.Replace($resultText, "") }
    $visible = Clean-Text $text
    $prefix = Clean-Text ([string]$Locator.prefix)
    $suffix = Clean-Text ([string]$Locator.suffix)
    if ($prefix -and -not $visible.StartsWith($prefix, [StringComparison]::Ordinal)) { return $false }
    if ($suffix -and -not $visible.EndsWith($suffix, [StringComparison]::Ordinal)) { return $false }
    return $true
}

function Set-PageLabelParagraphAlignment($Document, $Locators) {
    $changed = 0
    foreach ($locator in @($Locators)) {
        foreach ($sectionIndex in @($locator.section_indices)) {
            if ([int]$sectionIndex -lt 1 -or [int]$sectionIndex -gt $Document.Sections.Count) { continue }
            $section = $Document.Sections.Item([int]$sectionIndex)
            $stories = if ([string]$locator.story_type -eq "footer") { @(Get-ActiveFooters $section) } else { @(Get-ActiveHeaders $section) }
            foreach ($story in $stories) {
                $paragraphs = @()
                try { $paragraphs += @($story.Range.Paragraphs) } catch {}
                try {
                    $shapes = $story.Range.ShapeRange
                    for ($shapeIndex = 1; $shapeIndex -le $shapes.Count; $shapeIndex++) {
                        try { $paragraphs += @($shapes.Item($shapeIndex).TextFrame.TextRange.Paragraphs) } catch {}
                    }
                } catch {}
                foreach ($paragraph in $paragraphs) {
                    try {
                        if (Test-PageLabelParagraph $paragraph $locator) {
                            $paragraph.Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
                            $changed++
                        }
                    } catch {}
                }
            }
        }
    }
    return $changed
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

function Set-Headers($Document, [string]$Title, [bool]$SetHeaderText = $true, [bool]$SetPageNumbers = $true) {
    $cleanTitle = ($Title -replace '(?i)\s*V\s*1\.0\s*$', '').Trim()
    $headerText = $cleanTitle + " V1.0"
    if ($SetHeaderText -and $SetPageNumbers -and (Test-HeaderPageLayout $Document)) {
        return [pscustomobject]@{ Header = $false; PageNumbers = $false }
    }
    $headerChanged = $false
    $pageNumbersChanged = $false
    $pageTemplate = if ($SetPageNumbers) { Get-ExistingPageTemplate $Document } else { $null }
    for ($sectionIndex = 1; $sectionIndex -le $Document.Sections.Count; $sectionIndex++) {
        $section = $Document.Sections.Item($sectionIndex)
        $headers = @(Get-ActiveHeaders $section)
        $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            $header = $headers[$index]
            $footer = $footers[$index]
            if ($SetHeaderText -and -not (Story-HasHeaderText $header)) {
                # Add only the title text. Copying a whole header template here
                # could also copy its PAGE field when page numbers were not selected.
                Add-HeaderTextPreservingFields $header $section $headerText
                $headerChanged = $true
            }
            if ($SetPageNumbers) {
                if (Remove-ExtraPageFields $header $footer) { $pageNumbersChanged = $true }
                $pageCount = (Get-PageFieldCount $header) + (Get-PageFieldCount $footer)
                if ($pageCount -eq 0) {
                    $target = if ($null -ne $pageTemplate -and $pageTemplate.IsFooter) { $footer } else { $header }
                    if ($null -ne $pageTemplate -and $pageTemplate.IsFooter -and -not (Story-HasContent $target)) {
                        # A footer template carries page-number presentation; a
                        # header template may also contain the manual title.
                        Copy-StoryTemplate $pageTemplate.Story $target
                    } else {
                        Add-PageField $target
                    }
                    $pageNumbersChanged = $true
                }
            }
        }
    }
    return [pscustomobject]@{ Header = $headerChanged; PageNumbers = $pageNumbersChanged }
}

function Audit-EffectiveCover($Document, [string]$Title) {
    $items = @(Get-FrontParagraphs $Document 1)
    $block = Find-CoverBlock $items $Title 1 1 $true
    if ($null -eq $block) {
        $tableBlock = Find-LeadingTableCover $Document $items $Title
        if ($null -ne $tableBlock) {
            # The supplied table is protected as a unit, so acknowledge the
            # title/label without demanding typography edits inside its cells.
            return [pscustomobject]@{ Found = $true; ProtectedTable = $true; LabelPresent = $true; Label = $items[$tableBlock.LabelIndex].Text; Font = ""; LabelFont = ""; Bold = $false; Size = 0; LabelSize = 0; Lines = 0; OrphanedLastLine = $false; ExtraText = @(); TitleYRatio = -1; LabelYRatio = -1; GapPoints = -1; TitleNumbered = $false; LabelNumbered = $false }
        }
        return [pscustomobject]@{ Found = $false; ProtectedTable = $false; LabelPresent = $false; Label = ""; Font = ""; LabelFont = ""; Bold = $false; Size = 0; LabelSize = 0; Lines = 0; OrphanedLastLine = $false; ExtraText = @(); TitleYRatio = -1; LabelYRatio = -1; GapPoints = -1 }
    }
    $titleItem = $items[$block.TitleIndex]
    $labelItem = $items[$block.LabelIndex]
    $titleRange = $Document.Range($titleItem.Start, $titleItem.End)
    $labelRange = $Document.Range($labelItem.Start, $labelItem.End)
    # List numbering is not included in Range.Text, so inspect the owning
    # paragraphs explicitly. A nonzero ListType means Word will render a number
    # or bullet before the cover text even when its visible text looks correct.
    $titleListType = 0
    $labelListType = 0
    try { $titleListType = [int]$Document.Paragraphs.Item($titleItem.Index).Range.ListFormat.ListType } catch {}
    if ($null -ne $labelItem) { try { $labelListType = [int]$Document.Paragraphs.Item($labelItem.Index).Range.ListFormat.ListType } catch {} }
    $pageHeight = [double]$Document.Sections.Item(1).PageSetup.PageHeight
    $titleY = [double]$titleRange.Information($wdVerticalPositionRelativeToPage)
    $labelY = if ($null -ne $labelRange) { [double]$labelRange.Information($wdVerticalPositionRelativeToPage) } else { -1 }
    $extraText = @($items | Where-Object {
        ($_.Page -eq 1 -and $_.Text -ne "" -and -not (Test-CoverTitle $_.Text $Title) -and -not (Test-LegacyManualLabel $_.Text) -and $_.Text -match '[\p{L}\p{N}]')
    } | ForEach-Object { $_.Text })
    $titleLines = [int]$titleRange.ComputeStatistics($wdStatisticLines)
    return [pscustomobject]@{
        Found = $true
        ProtectedTable = $false
        LabelPresent = ($null -ne $labelRange)
        Label = $labelItem.Text
        Font = [string]$titleRange.Font.NameFarEast
        LabelFont = if ($null -ne $labelRange) { [string]$labelRange.Font.NameFarEast } else { "" }
        Bold = ([int]$titleRange.Font.Bold -eq -1 -and ($null -eq $labelRange -or [int]$labelRange.Font.Bold -eq -1))
        Size = [double]$titleRange.Font.Size
        LabelSize = if ($null -ne $labelRange) { [double]$labelRange.Font.Size } else { 0 }
        Lines = $titleLines
        OrphanedLastLine = ($titleLines -eq 2 -and (Get-CoverLastLineUnitCount $titleRange) -le 2)
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
    $cleanTitle = ($Title -replace '(?i)\s*V\s*1\.0\s*$', '').Trim()
    foreach ($section in $Document.Sections) {
        $headers = @(Get-ActiveHeaders $section)
        $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            $headerText = [string]$headers[$index].Range.Text
            foreach ($field in @(Get-PageFields $headers[$index])) {
                try { $headerText = $headerText.Replace([string]$field.Result.Text, "") } catch {}
            }
            $headerText = Clean-Text $headerText
            if ($headerText.IndexOf($cleanTitle, [StringComparison]::OrdinalIgnoreCase) -lt 0 -or $headerText -notmatch '(?i)V\s*1\.0') { $headerFound = $false }
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

function Trace-DuplicateTitlePositions($Document, [string]$Title, [string]$Stage) {
    $positions = @(
        Get-FrontParagraphs $Document 5 |
            Where-Object { $_.Page -ge 2 -and $_.Page -le 5 -and -not $_.InTable -and $_.Style -notmatch '(?i)TOC|目录' -and (Compact-Text $_.Text) -ceq (Compact-Text $Title) } |
            ForEach-Object { "p$($_.Page)@$($_.Start)-$($_.End)" }
    )
    Trace-Step $Title ("duplicate-position-" + $Stage + ":" + $(if ($positions.Count) { $positions -join "," } else { "none" }))
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
        $coverActionNames = @("ADD_COVER", "ENSURE_COVER_FIRST_PAGE", "AUDIT_ALL_PAGES", "AUDIT_PDF_ALL_PAGES", "REPAIR_BLANK_PAGES", "REMOVE_DUPLICATE_COVER_TEXT", "REMOVE_COVER_VERSION", "NORMALIZE_COVER_STYLE", "REPLACE_INVALID_COVER")
        $doCover = @($actions | Where-Object { $coverActionNames -contains $_ }).Count -gt 0
        # Word page mapping is the escalation path. Routine cover work uses the
        # exported PDF screen and does not map every paragraph through COM.
        $fullPageAudit = (Test-Action $actions "AUDIT_ALL_PAGES") -or (Test-Action $actions "REPAIR_BLANK_PAGES")
        $doHeader = Test-Action $actions "NORMALIZE_HEADER"
        $doPageNumbers = Test-Action $actions "NORMALIZE_PAGE_NUMBERS"
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
        $brokenBookmarkFieldsUnlinked = @()
        $failures = @()
        $blankRepair = $null
        try {
            Trace-Step $title "open"
            $doc = $word.Documents.Open($docx, $false, $false)
            Trace-Step $title "opened"
            $originalFrontPageText = if ($doCover) { @((Get-FrontParagraphs $doc 1 | Where-Object { $_.Text -ne "" } | ForEach-Object { $_.Text })) } else { @() }
            $originalCoverTextConflicts = @($originalFrontPageText | Where-Object {
                -not (Test-CoverTitle $_ $title) -and -not (Test-LegacyManualLabel $_)
            })
            $tableCoverPreserved = $doCover -and ($null -ne (Find-LeadingTableCover $doc @(Get-FrontParagraphs $doc 1) $title))

            $coverAdded = $false
            Trace-Step $title "cover-check"
            $replaceInvalidCover = if ($doCover -and -not $tableCoverPreserved) { Test-CoverNeedsReplacement $doc $title } else { $false }
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
                try {
                    $coverEnsure = Ensure-FirstPageCover $doc $title
                    if ($coverEnsure.Changed) {
                        $coverAdded = $true
                        if ($coverEnsure.Result -eq "LABEL_ADDED") { $applied += "ADD_MISSING_COVER_LABEL" }
                        elseif ($coverEnsure.Result -eq "COVER_ADDED") { $applied += "ADD_COVER" }
                    }
                } catch { $failures += ("补全封面失败：" + $_.Exception.Message) }
            }

            if ($doCover -and ((Test-Action $actions "REMOVE_DUPLICATE_COVER_TEXT") -or $coverAdded -or (Test-DuplicateCoverBlock $doc $title))) {
                Trace-Step $title "cover-remove-duplicates"
                $removed = Remove-DuplicateCoverBlocks $doc $title
                Trace-Step $title "cover-duplicates-removed"
                if ($removed -gt 0) { $applied += "REMOVE_DUPLICATE_COVER_TEXT" }
            }
            if ($doCover) { Trace-DuplicateTitlePositions $doc $title "after-remove" }

            if (Test-Action $actions "REMOVE_COVER_VERSION") {
                Trace-Step $title "cover-remove-version"
                $versionFix = Remove-CoverVersion $doc $title
                if ($versionFix.ProtectedTable) { $failures += "封面位于受保护表格中，无法安全去掉版本号，请人工处理" }
                elseif ($versionFix.Found -and $versionFix.Changed) { $applied += "REMOVE_COVER_VERSION" }
                elseif (-not $versionFix.Found) { $failures += "未找到封面标题，无法去掉版本号" }
            }

            $style = $null
            if (((Test-Action $actions "NORMALIZE_COVER_STYLE") -or $coverAdded) -and -not $tableCoverPreserved) {
                if (-not $coverAdded -and $replaceInvalidCover) {
                    [void](Remove-InvalidCoverFragments $doc $title)
                    Add-Cover $doc $title
                    $coverAdded = $true
                    $applied += "REPLACE_INVALID_COVER"
                }
                Trace-Step $title "cover-style"
                $style = Set-CoverTypography $doc $title
                Trace-Step $title "cover-styled"
                if ($null -ne $style.Workload) {
                    Trace-Step $title ("cover-workload=" + ($style.Workload | ConvertTo-Json -Compress))
                }
                Trace-DuplicateTitlePositions $doc $title "after-style"
                if (-not $style.Found) { $failures += "第一页未找到可排版的封面标题块" }
                else { $applied += "NORMALIZE_COVER_STYLE" }
                if ($style.LabelLayoutFailed) { $failures += "封面说明书标签格式化后无法与标题保持在第一页" }
            }

            if ($doCover) {
                Trace-Step $title "cover-boundary"
                if (Ensure-CoverPageBoundary $doc $title) {
                    $applied += "SEPARATE_COVER_FROM_BODY"
                }
                Trace-Step $title "cover-boundary-done"
                Trace-DuplicateTitlePositions $doc $title "after-boundary"
                Trace-Step $title "cover-frontmatter-cleanup"
                $frontMatterRemoved = Remove-RedundantPostCoverFrontMatter $doc $title
                Trace-Step $title "cover-frontmatter-cleaned"
                Trace-DuplicateTitlePositions $doc $title "after-frontmatter"
                if ($frontMatterRemoved -gt 0) { $applied += "TRIM_REDUNDANT_POST_COVER_FRONTMATTER" }
                # Page breaks and section boundaries can repaginate a supplied
                # duplicate cover onto page 2 after the initial pass. Run the
                # exact-match cleanup again after the cover/body boundary is
                # established. Routine blank-page screening runs on the PDF;
                # Word repair is reserved for explicit or anomaly-triggered work.
                Trace-Step $title "cover-late-duplicate-cleanup"
                $lateDuplicateRemoved = Remove-DuplicateCoverBlocks $doc $title
                if ($lateDuplicateRemoved -gt 0) {
                    $removed += $lateDuplicateRemoved
                    if ($applied -notcontains "REMOVE_DUPLICATE_COVER_TEXT") { $applied += "REMOVE_DUPLICATE_COVER_TEXT" }
                }
                Trace-Step $title "cover-late-duplicate-cleanup-done"
                Trace-DuplicateTitlePositions $doc $title "after-late-remove"
                Trace-Step $title "body-leading-blanks"
                $bodyBlanks = Remove-BodyLeadingBlankParagraphs $doc $title
                Trace-Step $title "body-leading-blanks-done"
                if ($bodyBlanks -gt 0) { $applied += "TRIM_BODY_LEADING_BLANKS" }
                if ($fullPageAudit) {
                    Trace-Step $title "blank-page-audit"
                    $blankRepair = if (Test-Action $actions "REPAIR_BLANK_PAGES") {
                        Repair-WordBlankPages -Document $doc -TargetPages @($item.retry_blank_pages)
                    } else {
                        Repair-WordBlankPages $doc
                    }
                    Trace-Step $title "blank-page-audit-done"
                    if ($blankRepair.Changed) { $applied += "REPAIR_WORD_BLANK_PAGES" }
                }
                Trace-DuplicateTitlePositions $doc $title "after-blank-repair"
            }

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

            $headerChanges = Set-Headers $doc $title $doHeader $doPageNumbers
            if ($doHeader -and $headerChanges.Header) { $applied += "NORMALIZE_HEADER" }
            if ($doPageNumbers -and $headerChanges.PageNumbers) { $applied += "NORMALIZE_PAGE_NUMBERS" }
            if (Test-Action $actions "REPAIR_TRUNCATED_PAGE_LABELS") {
                $labelChanges = Set-PageLabelParagraphAlignment $doc @($item.page_label_templates)
                if ($labelChanges -gt 0) {
                    $applied += "REPAIR_TRUNCATED_PAGE_LABELS"
                } else {
                    $failures += "已确认截断的页码模板未能在对应 Word 段落中唯一定位"
                }
            }

            Trace-Step $title "fields"
            foreach ($section in $doc.Sections) {
                foreach ($header in $section.Headers) {
                    if ($header.Exists) { $header.Range.Fields.Update() | Out-Null }
                }
            }
            if (Test-Action $actions "EXPORT_PDF") {
                Trace-Step $title "missing-bookmark-field-audit"
                $brokenBookmarkFieldsUnlinked = @(Repair-MissingBookmarkFields $doc)
                Trace-Step $title ("missing-bookmark-fields-unlinked:" + $brokenBookmarkFieldsUnlinked.Count)
            }
            Trace-Step $title "save"
            $doc.Save()
            $wordPageAudit = if ($doCover -and $fullPageAudit) { @(Get-WordPageAudit $doc) } else { @() }
            if (Test-Action $actions "EXPORT_PDF") {
                Trace-Step $title "export"
                $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF)
                Trace-Step $title "exported"
                $applied += "EXPORT_PDF"
            }

            $coverAudit = if ($doCover) { Audit-EffectiveCover $doc $title } else { $null }
            $headerAudit = Audit-Headers $doc $title
            if ($doCover) { Trace-DuplicateTitlePositions $doc $title "before-final-qa" }
            if ($doCover -and (Test-DuplicateCoverBlock $doc $title)) { $failures += "仍存在后续重复封面标题块" }
            $unresolvedBlankPages = if ($null -ne $blankRepair) {
                @($blankRepair.UnresolvedPages | ForEach-Object { [int]$_.page })
            } else {
                @()
            }
            if ($doCover -and $unresolvedBlankPages.Count -gt 0) { $failures += "全书仍存在无法安全修复的空白页: $($unresolvedBlankPages -join ',')" }
            if ($doCover -and -not $coverAudit.Found) { $failures += "第一页缺少有效封面标题块" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and @($coverAudit.ExtraText).Count -gt 0) { $failures += "封面页仍混有目录或正文文字" }
            if ($doCover -and $coverAudit.Lines -gt 2) { $failures += "软件名称超过两行" }
            if ($doCover -and $coverAudit.OrphanedLastLine) { $failures += "软件名称第二行仅剩一至两个词" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and -not $coverAudit.Bold) { $failures += "封面两行未同时加粗" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and $coverAudit.LabelPresent -and $coverAudit.Size -ne $coverAudit.LabelSize) { $failures += "封面两行字号不一致" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and ($coverAudit.Font -notmatch '宋体|SimSun' -or ($coverAudit.LabelPresent -and $coverAudit.LabelFont -notmatch '宋体|SimSun'))) { $failures += "封面标题或说明书行未使用宋体" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and $coverAudit.TitleYRatio -ge 0 -and ($coverAudit.TitleYRatio -lt 0.30 -or $coverAudit.TitleYRatio -gt 0.56)) { $failures += "软件名称未位于封面中段偏上" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and $coverAudit.LabelPresent -and $coverAudit.LabelYRatio -ge 0 -and ($coverAudit.LabelYRatio -lt 0.48 -or $coverAudit.LabelYRatio -gt 0.72)) { $failures += "用户手册/使用说明书未位于封面中部附近" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and $coverAudit.LabelPresent -and $coverAudit.GapPoints -ge 0 -and $coverAudit.GapPoints -lt 48) { $failures += "软件名称与用户手册/使用说明书之间留白不足" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and $coverAudit.TitleNumbered) { $failures += "封面标题仍带自动编号" }
            if ($doCover -and -not $coverAudit.ProtectedTable -and $coverAudit.LabelPresent -and $coverAudit.LabelNumbered) { $failures += "封面说明书行仍带自动编号" }
            if ($doHeader -and -not $headerAudit.Header) { $failures += "页眉缺失" }
            if ($doPageNumbers -and -not $headerAudit.Page) { $failures += "动态PAGE字段缺失或重复" }
            if ((Test-Action $actions "REMOVE_TITLE_DOUBLE_QUOTES") -and (Get-TitleDoubleQuoteCount $doc $title) -gt 0) { $failures += "全文仍存在外带双引号的软件名称" }
            if ((Test-Action $actions "REMOVE_ALL_DOUBLE_QUOTES") -and (Get-ManualPatternCount $doc '["“”「」『』]') -gt 0) { $failures += "说明书全文仍存在双引号" }
            if ((Test-Action $actions "REPLACE_YEARS_WITH_CURRENT") -and (Get-ManualPatternCount $doc '(?:19|20)\d{2}(?:[\s,，、]*(?:19|20)\d{2})*\s*年') -gt 0) { $failures += "说明书全文仍存在年份表述" }
            # A QA finding must not undo successfully applied page headers,
            # page numbers, cover separation, or PDF export. Keep the repaired
            # file and report the remaining defect for review.

            $results += [pscustomobject]@{
                name = $title
                cover_rule_version = [int]$script:CoverRuleData.version
                name_source = [string]$item.name_source
                docx_title = [string]$item.docx_title
                name_conflicts = @($item.name_conflicts)
                folder = [string]$item.folder
                status = if ($failures.Count -eq 0) { "OK" } else { "QA_FAILED" }
                changed = $true
                actions_requested = $actions
                actions_applied = @($applied | Select-Object -Unique)
                duplicate_blocks_removed = $removed
                broken_bookmark_fields_unlinked = @($brokenBookmarkFieldsUnlinked)
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
                cover_typography_workload = if ($null -ne $style) { $style.Workload } else { $null }
                font_size = if ($doCover) { $coverAudit.Size } else { 0 }
                title_lines = if ($doCover) { $coverAudit.Lines } else { 0 }
                title_y_ratio = if ($doCover) { $coverAudit.TitleYRatio } else { -1 }
                label_y_ratio = if ($doCover) { $coverAudit.LabelYRatio } else { -1 }
                title_label_gap_points = if ($doCover) { $coverAudit.GapPoints } else { -1 }
                cover_label_present = if ($doCover) { $coverAudit.LabelPresent } else { $false }
                cover_label = if ($doCover) { $coverAudit.Label } else { "" }
                cover_label_layout_failed = if ($null -ne $style) { [bool]$style.LabelLayoutFailed } else { $false }
                cover_extra_text = if ($doCover) { @($coverAudit.ExtraText) } else { @() }
                original_cover_text = @($originalFrontPageText)
                cover_text_conflicts = @($originalCoverTextConflicts)
                front_page_text = if ($doCover) { @((Get-FrontParagraphs $doc 1 | Where-Object { $_.Text -ne "" } | ForEach-Object { $_.Text })) } else { @() }
                word_pages = [int]$doc.ComputeStatistics($wdStatisticPages)
                page_audit_mode = if ($fullPageAudit) { "full_word_com" } elseif ($doCover) { "deferred_to_pdf_screening" } else { "not_requested" }
                page_audit = $wordPageAudit
                word_blank_page_candidates = if ($doCover) { @($wordPageAudit | Where-Object { -not $_.has_body_content }) } else { @() }
                blank_page_repairs = if ($null -ne $blankRepair) { @($blankRepair.Changes) } else { @() }
                unresolved_blank_pages = if ($null -ne $blankRepair) { @($blankRepair.UnresolvedPages | ForEach-Object { [int]$_.page }) } else { @() }
                unresolved_blank_page_details = if ($null -ne $blankRepair) { @($blankRepair.UnresolvedPages) } else { @() }
                word_blank_page_audit_count = if ($null -ne $blankRepair) { $blankRepair.AuditCount } else { $null }
                word_blank_page_diagnostics = if ($null -ne $blankRepair) { @($blankRepair.Diagnostics) } else { @() }
                attempts = @([pscustomobject]@{ attempt = 1; phase = if ($fullPageAudit) { "Word处理、导出和完整页面映射" } else { "Word封面处理、保存和导出；页面筛查转交PDF" }; status = if ($failures.Count -eq 0) { "PASSED" } else { "FAILED" }; failures = @($failures) })
                failures = $failures
                seconds = [Math]::Round(([DateTime]::UtcNow - $started).TotalSeconds, 3)
            }
        } catch {
            $message = $_.Exception.Message
            if ($null -ne $doc) {
                # Preserve any changes completed before the failing step.
                try { $doc.Save() } catch {}
                if (Test-Action $actions "EXPORT_PDF") { try { $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF) } catch {} }
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
                broken_bookmark_fields_unlinked = @($brokenBookmarkFieldsUnlinked)
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
        version = "3.4.28"
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

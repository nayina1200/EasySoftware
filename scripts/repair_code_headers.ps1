param(
    [Parameter(Mandatory = $true)][string]$BaseDir,
    [Parameter(Mandatory = $true)][string]$PlanPath,
    [Parameter(Mandatory = $true)][string]$ReportPath,
    [string]$PidPath = "",
    [string]$ProgressPath = "",
    [int]$ProgressTotal = 0,
    [string]$ProgressPhase = "代码处理"
)

$ErrorActionPreference = "Stop"
$wdHeaderPrimary = 1
$wdHeaderFirstPage = 2
$wdHeaderEvenPages = 3
$wdFieldPage = 33
$wdCollapseEnd = 0
$wdExportFormatPDF = 17
$wdFontSize9 = 9
$wdAlignParagraphLeft = 0
$wdAlignTabRight = 2
$wdTabLeaderSpaces = 0
$ProgressStartedAt = [DateTime]::UtcNow

function Clean([string]$value) { return (($value -replace "[\x00-\x1F\x7F\uFFFC]", "") -replace "\s+", " ").Trim() }
function Safe([string]$value) { return ($value -replace '[<>:"/\\|?*]', '_') }
function Write-WorkflowProgress([int]$Current, [string]$Title, [bool]$CurrentProjectActive = $false) {
    if ([string]::IsNullOrWhiteSpace($ProgressPath)) { return }
    $total = [Math]::Max($ProgressTotal, $Current)
    $elapsed = ([DateTime]::UtcNow - $ProgressStartedAt).TotalSeconds
    $estimate = if ($Current -gt 0) { ($elapsed / $Current) * [Math]::Max(0, $total - $Current) } else { $null }
    $state = [ordered]@{
        version = "3.4.23"; status = "RUNNING"; phase = $ProgressPhase
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
function Story-HasContent($story) {
    if (-not $story.Exists) { return $false }
    if ((Clean $story.Range.Text) -ne "") { return $true }
    if (@($story.Range.Fields).Count -gt 0) { return $true }
    try { if ($story.Range.InlineShapes.Count -gt 0 -or $story.Range.ShapeRange.Count -gt 0) { return $true } } catch {}
    return $false
}
function Get-ActiveHeaders($section) {
    $headers = @($section.Headers.Item($wdHeaderPrimary))
    if ($section.PageSetup.DifferentFirstPageHeaderFooter) { $headers += $section.Headers.Item($wdHeaderFirstPage) }
    if ($section.PageSetup.OddAndEvenPagesHeaderFooter) { $headers += $section.Headers.Item($wdHeaderEvenPages) }
    return @($headers)
}
function Get-ActiveStories($section) {
    $stories = @(Get-ActiveHeaders $section)
    $stories += $section.Footers.Item($wdHeaderPrimary)
    if ($section.PageSetup.DifferentFirstPageHeaderFooter) { $stories += $section.Footers.Item($wdHeaderFirstPage) }
    if ($section.PageSetup.OddAndEvenPagesHeaderFooter) { $stories += $section.Footers.Item($wdHeaderEvenPages) }
    return @($stories)
}
function Get-ActiveFooters($section) {
    $footers = @($section.Footers.Item($wdHeaderPrimary))
    if ($section.PageSetup.DifferentFirstPageHeaderFooter) { $footers += $section.Footers.Item($wdHeaderFirstPage) }
    if ($section.PageSetup.OddAndEvenPagesHeaderFooter) { $footers += $section.Footers.Item($wdHeaderEvenPages) }
    return @($footers)
}
function Get-PageFieldCount($section) {
    $count = 0
    foreach ($story in @(Get-ActiveStories $section)) {
        if (-not $story.Exists) { continue }
        $count += @($story.Range.Fields | Where-Object { (Clean $_.Code.Text) -match '(?i)\bPAGE\b' }).Count
    }
    return $count
}
function Get-PageFields($story) {
    if (-not $story.Exists) { return @() }
    return @($story.Range.Fields | Where-Object { (Clean $_.Code.Text) -match '(?i)\bPAGE\b' })
}
function Story-HasHeaderText($story) {
    if (-not $story.Exists) { return $false }
    $text = [string]$story.Range.Text
    foreach ($field in @(Get-PageFields $story)) { $text = $text.Replace([string]$field.Result.Text, "") }
    return -not [string]::IsNullOrWhiteSpace((Clean $text))
}
function Get-PageFieldEntries($section) {
    $entries = @()
    foreach ($story in @(Get-ActiveHeaders $section)) {
        foreach ($field in @(Get-PageFields $story)) {
            $entries += [pscustomobject]@{ Story = $story; Field = $field; IsFooter = $false }
        }
    }
    $footers = @($section.Footers.Item($wdHeaderPrimary))
    if ($section.PageSetup.DifferentFirstPageHeaderFooter) { $footers += $section.Footers.Item($wdHeaderFirstPage) }
    if ($section.PageSetup.OddAndEvenPagesHeaderFooter) { $footers += $section.Footers.Item($wdHeaderEvenPages) }
    foreach ($story in $footers) {
        foreach ($field in @(Get-PageFields $story)) {
            $entries += [pscustomobject]@{ Story = $story; Field = $field; IsFooter = $true }
        }
    }
    return @($entries)
}
function Remove-PageField($field) {
    try { $field.Delete(); return } catch {}
    try { $field.Result.Delete() } catch {}
}
function Add-HeaderTextPreservingFields($header, $section, [string]$headerText) {
    $header.LinkToPrevious = $false
    $fields = @(Get-PageFields $header)
    if ($fields.Count -gt 0) {
        $insert = $header.Range.Duplicate
        $insert.Collapse($wdCollapseStart)
        $insert.InsertBefore($headerText + "`t")
    } else {
        $header.Range.Text = $headerText
    }
    $header.Range.Font.Name = "宋体"
    $header.Range.Font.NameFarEast = "宋体"
    $header.Range.Font.NameAscii = "SimSun"
    $header.Range.Font.NameOther = "SimSun"
    $header.Range.Font.Size = $wdFontSize9
    $header.Range.Font.Bold = 0
    $header.Range.ParagraphFormat.Alignment = $wdAlignParagraphLeft
    $header.Range.ParagraphFormat.TabStops.ClearAll()
    $width = $section.PageSetup.PageWidth - $section.PageSetup.LeftMargin - $section.PageSetup.RightMargin
    [void]$header.Range.ParagraphFormat.TabStops.Add($width, $wdAlignTabRight, $wdTabLeaderSpaces)
}
function Add-PageField($story) {
    # Add the field through the collapsed insertion range. Calling Fields.Add
    # on the full story range can silently place the field outside the header
    # or footer in some Word documents, leaving the exported PDF unnumbered.
    $pageRange = $story.Range.Duplicate
    if ($pageRange.End -gt $pageRange.Start) { $pageRange.End-- }
    $pageRange.Collapse($wdCollapseEnd)
    $pageRange.InsertAfter("`t")
    $pageRange.Collapse($wdCollapseEnd)
    [void]$pageRange.Fields.Add($pageRange, $wdFieldPage)
}
function Get-PageFieldCount($story) { return @(Get-PageFields $story).Count }
function Remove-ExtraPageFields($header, $footer) {
    $entries = @(); foreach ($field in @(Get-PageFields $header)) { $entries += $field }; foreach ($field in @(Get-PageFields $footer)) { $entries += $field }
    $changed = $false
    for ($index = 1; $index -lt $entries.Count; $index++) { Remove-PageField $entries[$index]; $changed = $true }
    return $changed
}
function Get-ExistingHeaderTemplate($document) {
    foreach ($section in $document.Sections) { foreach ($header in @(Get-ActiveHeaders $section)) { if (Story-HasHeaderText $header) { return $header } } }
    return $null
}
function Get-ExistingPageTemplate($document) {
    foreach ($section in $document.Sections) { foreach ($footer in @(Get-ActiveFooters $section)) { if ((Get-PageFieldCount $footer) -gt 0) { return [pscustomobject]@{ Story = $footer; IsFooter = $true } } } }
    foreach ($section in $document.Sections) { foreach ($header in @(Get-ActiveHeaders $section)) { if ((Get-PageFieldCount $header) -gt 0) { return [pscustomobject]@{ Story = $header; IsFooter = $false } } } }
    return $null
}
function Copy-StoryTemplate($source, $target) { $target.LinkToPrevious = $false; $target.Range.FormattedText = $source.Range.FormattedText }
function Test-HeaderPageLayout($document) {
    foreach ($section in $document.Sections) {
        $headers = @(Get-ActiveHeaders $section); $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            if (-not (Story-HasHeaderText $headers[$index])) { return $false }
            if ((Get-PageFieldCount $headers[$index]) + (Get-PageFieldCount $footers[$index]) -ne 1) { return $false }
        }
    }
    return $true
}
function Set-CodeHeader($document, [string]$title) {
    $cleanTitle = ($title -replace '(?i)\s*V\s*1\.0\s*$', '').Trim()
    $headerText = $cleanTitle + " V1.0"
    if (Test-HeaderPageLayout $document) { return $false }
    $changed = $false; $headerTemplate = Get-ExistingHeaderTemplate $document; $pageTemplate = Get-ExistingPageTemplate $document
    foreach ($section in $document.Sections) {
        $headers = @(Get-ActiveHeaders $section); $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            $header = $headers[$index]; $footer = $footers[$index]
            if (-not (Story-HasHeaderText $header)) {
                if ($null -ne $headerTemplate) { Copy-StoryTemplate $headerTemplate $header } else { Add-HeaderTextPreservingFields $header $section $headerText }
                $changed = $true
            }
            if (Remove-ExtraPageFields $header $footer) { $changed = $true }
            # A copied header can already include PAGE; count after copying to
            # preserve it instead of adding a second page number in the footer.
            $pageCount = (Get-PageFieldCount $header) + (Get-PageFieldCount $footer)
            if ($pageCount -eq 0) {
                if ($null -ne $pageTemplate) {
                    $target = if ($pageTemplate.IsFooter) { $footer } else { $header }
                    if (-not (Story-HasContent $target)) { Copy-StoryTemplate $pageTemplate.Story $target } else { Add-PageField $target }
                } else { Add-PageField $header }
                $changed = $true
            }
        }
    }
    return $changed
}
function Update-CodeHeaderFields($document) {
    try { $document.Application.Options.UpdateFieldsAtPrint = $true } catch {}
    try { [void]$document.Fields.Update() } catch {}
    foreach ($section in $document.Sections) {
        foreach ($story in @($section.Headers) + @($section.Footers)) {
            try { [void]$story.Range.Fields.Update() } catch {}
        }
    }
}
function Audit-CodeHeader($document, [string]$title) {
    $headerFound = $true; $pageFound = $true
    foreach ($section in $document.Sections) {
        $headers = @(Get-ActiveHeaders $section); $footers = @(Get-ActiveFooters $section)
        for ($index = 0; $index -lt $headers.Count; $index++) {
            if (-not (Story-HasHeaderText $headers[$index])) { $headerFound = $false }
            if ((Get-PageFieldCount $headers[$index]) + (Get-PageFieldCount $footers[$index]) -ne 1) { $pageFound = $false }
        }
    }
    return [pscustomobject]@{ Header = $headerFound; Page = $pageFound }
}

$base = (Resolve-Path -LiteralPath $BaseDir).Path
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$reportDir = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($ReportPath))
$backupRoot = Join-Path $reportDir "code_backups"
[IO.Directory]::CreateDirectory($backupRoot) | Out-Null
$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
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
        $docx = Join-Path $folder ([string]$item.docx)
        $pdf = if (-not [string]::IsNullOrWhiteSpace([string]$item.pdf)) { Join-Path $folder ([string]$item.pdf) } else { [IO.Path]::ChangeExtension($docx, ".pdf") }
        $backup = Join-Path $backupRoot (Safe $title)
        [IO.Directory]::CreateDirectory($backup) | Out-Null
        Copy-Item -LiteralPath $docx -Destination (Join-Path $backup ([IO.Path]::GetFileName($docx))) -Force
        if (Test-Path -LiteralPath $pdf) { Copy-Item -LiteralPath $pdf -Destination (Join-Path $backup ([IO.Path]::GetFileName($pdf))) -Force }
        $doc = $null; $failures = @(); $started = [DateTime]::UtcNow
        $temporaryPdf = Join-Path $folder (([IO.Path]::GetFileNameWithoutExtension($pdf)) + ".easysoftware-" + [Guid]::NewGuid().ToString("N") + ".tmp.pdf")
        try {
            $doc = $word.Documents.Open($docx, $false, $false)
            $headerChanged = Set-CodeHeader $doc $title
            Update-CodeHeaderFields $doc
            # Export to a sidecar first. Replacing an existing PDF directly can
            # block Word when a browser or PDF reader still has it open.
            # Forcing Repaginate here lays out every code page first and can take
            # minutes on long source documents without improving the output.
            $doc.Save(); $doc.ExportAsFixedFormat($temporaryPdf, $wdExportFormatPDF)
            if (-not (Test-Path -LiteralPath $temporaryPdf -PathType Leaf)) { throw "Word未生成临时代码PDF" }
            $audit = Audit-CodeHeader $doc $title
            if (-not ($audit.Header -and $audit.Page)) { $failures += "code header or PAGE incomplete" }
            if ($failures.Count -gt 0) { throw ("quality check failed: " + ($failures -join "; ")) }
            $wordPages = [int]$doc.ComputeStatistics(2)
            if ($null -ne $doc) { $doc.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null; $doc = $null }
            Move-Item -LiteralPath $temporaryPdf -Destination $pdf -Force
            $status = "OK"
            $result = New-Object PSObject
            $result | Add-Member NoteProperty name $title
            $result | Add-Member NoteProperty folder ([string]$item.folder)
            $result | Add-Member NoteProperty status $status
            $result | Add-Member NoteProperty changed $true
            $actions = @("EXPORT_CODE_PDF")
            if ($headerChanged) { $actions = @("NORMALIZE_CODE_HEADER", "EXPORT_CODE_PDF") }
            $result | Add-Member NoteProperty actions_applied $actions
            $result | Add-Member NoteProperty failures $failures
            $result | Add-Member NoteProperty word_pages $wordPages
            $result | Add-Member NoteProperty seconds ([Math]::Round(([DateTime]::UtcNow-$started).TotalSeconds,3))
            $results += $result
        } catch {
            $message = $_.Exception.Message
            if (Test-Path -LiteralPath $temporaryPdf) { Remove-Item -LiteralPath $temporaryPdf -Force -ErrorAction SilentlyContinue }
            if ($null -ne $doc) { try {$doc.Close($false)} catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null; $doc=$null }
            # Keep the repaired artifacts for review; the workflow must report
            # the defect instead of silently restoring the original files.
            $result = New-Object PSObject
            $result | Add-Member NoteProperty name $title
            $result | Add-Member NoteProperty folder ([string]$item.folder)
            $result | Add-Member NoteProperty status "PRESERVED_FOR_REVIEW"
            $result | Add-Member NoteProperty changed $false
            $result | Add-Member NoteProperty actions_applied @()
            $result | Add-Member NoteProperty failures @($message)
            $result | Add-Member NoteProperty word_pages 0
            $result | Add-Member NoteProperty seconds ([Math]::Round(([DateTime]::UtcNow-$started).TotalSeconds,3))
            $results += $result
        } finally {
            if (Test-Path -LiteralPath $temporaryPdf) { Remove-Item -LiteralPath $temporaryPdf -Force -ErrorAction SilentlyContinue }
            if($null -ne $doc){$doc.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($doc)|Out-Null}
        }
        $progressIndex++
        Write-WorkflowProgress $progressIndex $title
    }
} finally {
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
}
[pscustomobject]@{
    version = "3.4.23"
    created_at = [DateTime]::Now.ToString("o")
    base = $base
    packages = $results
} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReportPath -Encoding UTF8

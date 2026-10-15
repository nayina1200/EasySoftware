param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [string]$TemporaryPdfPath = "",
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [string]$PidPath = "",
    [switch]$CountOnly,
    [int]$MinimumPages = 0,
    [int]$MinimumPageLines = 0,
    [string]$SupplementPath = ""
)

$ErrorActionPreference = "Stop"
$word = $null
$doc = $null
$started = [DateTime]::UtcNow

function Write-Result([hashtable]$Value) {
    $Value.seconds = [Math]::Round(([DateTime]::UtcNow - $started).TotalSeconds, 2)
    [IO.File]::WriteAllText(
        [IO.Path]::GetFullPath($ResultPath),
        ($Value | ConvertTo-Json -Depth 8),
        (New-Object Text.UTF8Encoding($false))
    )
}

try {
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
        [IO.File]::WriteAllText([IO.Path]::GetFullPath($PidPath), [string]$wordPid, (New-Object Text.UTF8Encoding($false)))
    }

    $ensureMinimum = $MinimumPages -gt 0 -and $MinimumPageLines -gt 0
    $doc = $word.Documents.Open((Resolve-Path -LiteralPath $DocxPath).Path, $false, [bool]($CountOnly -and -not $ensureMinimum))
    if ($CountOnly) {
        $supplemented = 0
        $supplementLines = @()
        if ($ensureMinimum) {
            if ([string]::IsNullOrWhiteSpace($SupplementPath) -or -not (Test-Path -LiteralPath $SupplementPath -PathType Leaf)) {
                throw "缺少代码补充说明文件"
            }
            $supplementLines = [IO.File]::ReadAllLines([IO.Path]::GetFullPath($SupplementPath), [Text.Encoding]::UTF8)
        }
        for ($attempt = 0; $attempt -lt 4; $attempt++) {
            $doc.Repaginate()
            $pages = [int]$doc.ComputeStatistics(2)
            $pageLines = 0
            if ($pages -ge 60) {
                $start = [int]$doc.GoTo(1, 1, 60).Start
                $end = if ($pages -ge 61) { [int]$doc.GoTo(1, 1, 61).Start } else { [int]$doc.Content.End }
                $pageRange = $doc.Range($start, $end)
                $pageLines = [int]$pageRange.ComputeStatistics(1)
            }
            if (-not $ensureMinimum -or ($pages -ge $MinimumPages -and $pageLines -ge $MinimumPageLines)) { break }
            if ($attempt -eq 3) { break }
            $needed = [Math]::Max(80, ($MinimumPages - $pages + 1) * 75)
            $remaining = $supplementLines.Length - $supplemented
            if ($remaining -le 0) { break }
            $take = [Math]::Min($needed, $remaining)
            $text = ($supplementLines[$supplemented..($supplemented + $take - 1)] -join "`r") + "`r"
            if ($pages -ge 60) {
                $insertAt = [int]$doc.GoTo(1, 1, 60).Paragraphs.Item(1).Range.Start
            } else {
                $insertAt = [int]$doc.Content.End - 1
                $text = "`r" + $text
            }
            $doc.Range($insertAt, $insertAt).InsertBefore($text)
            $inserted = $doc.Range($insertAt, $insertAt + $text.Length - 1)
            $inserted.Font.Name = "Consolas"
            $inserted.Font.Size = 8
            $inserted.ParagraphFormat.SpaceBefore = 0
            $inserted.ParagraphFormat.SpaceAfter = 0
            $inserted.ParagraphFormat.LineSpacingRule = 0
            $supplemented += $take
            $doc.Save()
        }
        if ($ensureMinimum -and ($pages -lt $MinimumPages -or $pageLines -lt $MinimumPageLines)) {
            throw "补足后仍不足：共 $pages 页，第 60 页 $pageLines 行"
        }
        Write-Result @{ status = "OK"; docx = [IO.Path]::GetFullPath($DocxPath); pages = $pages; page_60_lines = $pageLines; supplemented_lines = $supplemented }
        exit 0
    }
    try { [void]$doc.Fields.Update() } catch {}
    foreach ($section in $doc.Sections) {
        foreach ($story in @($section.Headers) + @($section.Footers)) {
            try { [void]$story.Range.Fields.Update() } catch {}
        }
    }
    $doc.Save()
    if ($ensureMinimum) {
        $doc.Repaginate()
        $pages = [int]$doc.ComputeStatistics(2)
        $pageLines = 0
        if ($pages -ge 60) {
            $start = [int]$doc.GoTo(1, 1, 60).Start
            $end = if ($pages -ge 61) { [int]$doc.GoTo(1, 1, 61).Start } else { [int]$doc.Content.End }
            $pageLines = [int]$doc.Range($start, $end).ComputeStatistics(1)
        }
        if ($pages -lt $MinimumPages -or $pageLines -lt $MinimumPageLines) {
            throw "导出前排版不达标：共 $pages 页，第 60 页 $pageLines 行"
        }
    }
    $doc.ExportAsFixedFormat(([IO.Path]::GetFullPath($TemporaryPdfPath)), 17)
    if (-not (Test-Path -LiteralPath $TemporaryPdfPath -PathType Leaf) -or (Get-Item -LiteralPath $TemporaryPdfPath).Length -eq 0) {
        throw "Word 未生成有效 PDF"
    }
    Write-Result @{ status = "OK"; docx = [IO.Path]::GetFullPath($DocxPath); pdf = [IO.Path]::GetFullPath($TemporaryPdfPath); pages = $pages; page_60_lines = $pageLines }
    exit 0
} catch {
    Write-Result @{ status = "FAILED"; docx = [IO.Path]::GetFullPath($DocxPath); error = $_.Exception.Message }
    exit 2
} finally {
    if ($null -ne $doc) { try { $doc.Close($false) } catch {}; try { [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {} }
    if ($null -ne $word) { try { $word.Quit() } catch {}; try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {} }
}

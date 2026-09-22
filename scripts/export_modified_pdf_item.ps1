param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$TemporaryPdfPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [string]$PidPath = ""
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

    $doc = $word.Documents.Open((Resolve-Path -LiteralPath $DocxPath).Path, $false, $false)
    try { [void]$doc.Fields.Update() } catch {}
    foreach ($section in $doc.Sections) {
        foreach ($story in @($section.Headers) + @($section.Footers)) {
            try { [void]$story.Range.Fields.Update() } catch {}
        }
    }
    $doc.Save()
    $doc.ExportAsFixedFormat(([IO.Path]::GetFullPath($TemporaryPdfPath)), 17)
    if (-not (Test-Path -LiteralPath $TemporaryPdfPath -PathType Leaf) -or (Get-Item -LiteralPath $TemporaryPdfPath).Length -eq 0) {
        throw "Word 未生成有效 PDF"
    }
    Write-Result @{ status = "OK"; docx = [IO.Path]::GetFullPath($DocxPath); pdf = [IO.Path]::GetFullPath($TemporaryPdfPath) }
    exit 0
} catch {
    Write-Result @{ status = "FAILED"; docx = [IO.Path]::GetFullPath($DocxPath); error = $_.Exception.Message }
    exit 2
} finally {
    if ($null -ne $doc) { try { $doc.Close($false) } catch {}; try { [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {} }
    if ($null -ne $word) { try { $word.Quit() } catch {}; try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {} }
}

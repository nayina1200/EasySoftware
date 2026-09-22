param(
    [Parameter(Mandatory = $true)]
    [string]$BaseDir
)

$ErrorActionPreference = "Stop"
$wdExportFormatPDF = 17
$base = (Resolve-Path -LiteralPath $BaseDir).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($docx in Get-ChildItem -LiteralPath $base -Recurse -File -Filter "*说明.docx" | Sort-Object FullName) {
        $doc = $null
        try {
            $doc = $word.Documents.Open($docx.FullName, $false, $false)
            $doc.Fields.Update() | Out-Null
            foreach ($section in $doc.Sections) {
                foreach ($header in $section.Headers) {
                    if ($header.Exists) { $header.Range.Fields.Update() | Out-Null }
                }
            }
            $doc.Save()
            $pdf = [IO.Path]::ChangeExtension($docx.FullName, ".pdf")
            $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF)
            Write-Output ("{0}`tOK" -f $docx.BaseName)
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


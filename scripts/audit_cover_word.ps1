param([Parameter(Mandatory = $true)][string]$BaseDir)
$ErrorActionPreference = "Stop"
function Clean([string]$s) { return (($s -replace "[\r\n\a\f\v]", "") -replace "\s+", " ").Trim() }
function IsLabel([string]$s) { return (Clean $s) -match '^[【\[]?(使用说明书|使用说明|说明书|使用手册|用户手册|用户使用手册|用户操作手册|用户操作说明|操作说明书|操作说明|操作手册)[】\]]?$' }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$rows = @()
try {
  foreach($folder in Get-ChildItem -LiteralPath $BaseDir -Directory | Sort-Object Name) {
    $docx=Get-ChildItem -LiteralPath $folder.FullName -File -Filter '*说明.docx' | Select-Object -First 1
    if($null -eq $docx){continue}
    $txt=Get-ChildItem -LiteralPath $folder.FullName -File -Filter '*.txt' | Select-Object -First 1
    $title=if($txt){$txt.BaseName}else{$folder.Name}
    $doc=$word.Documents.Open($docx.FullName,$false,$true)
    try {
      $tp=$null;$lp=$null
      for($i=1;$i -le [Math]::Min(20,$doc.Paragraphs.Count);$i++){
        $p=$doc.Paragraphs.Item($i);$t=Clean $p.Range.Text
        if($null -eq $tp -and $t -eq $title){$tp=$p.Range.Duplicate}
        if($null -eq $lp -and (IsLabel $t)){$lp=$p.Range.Duplicate}
      }
      $rows += [pscustomobject]@{
        Name=$title
        TitleFont=if($tp){$tp.Font.NameFarEast}else{''}
        LabelFont=if($lp){$lp.Font.NameFarEast}else{''}
        TitleBold=if($tp){$tp.Font.Bold}else{0}
        LabelBold=if($lp){$lp.Font.Bold}else{0}
        TitleSize=if($tp){$tp.Font.Size}else{0}
        LabelSize=if($lp){$lp.Font.Size}else{0}
        TitleLines=if($tp){$tp.ComputeStatistics(1)}else{0}
      }
    } finally {$doc.Close($false);[Runtime.InteropServices.Marshal]::ReleaseComObject($doc)|Out-Null}
  }
} finally {$word.Quit();[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null}
$rows | Format-Table -AutoSize

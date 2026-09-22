param(
    [Parameter(Mandatory = $true)][string]$BaseDir,
    [string[]]$AddCoverTitles = @(),
    [switch]$SkipHeaderNormalization
)
$ErrorActionPreference = 'Stop'
$wdSectionBreakNextPage=2;$wdAlignCenter=1;$wdAlignLeft=0;$wdAlignRight=2
$wdAlignTabRight=2;$wdTabLeaderSpaces=0;$wdVerticalCenter=1;$wdHeaderPrimary=1
$wdFieldPage=33;$wdCollapseEnd=0;$wdLines=1;$wdExportPDF=17

function Add-Cover($doc,[string]$title){
  $doc.Range(0,0).InsertBreak($wdSectionBreakNextPage)
  $text=$title+"`r`r`r使用说明书";$doc.Range(0,0).InsertBefore($text)
  $tr=$doc.Paragraphs.Item(1).Range.Duplicate;$sr=$doc.Paragraphs.Item(4).Range.Duplicate
  foreach($r in @($tr,$sr)){
    try{$r.ListFormat.RemoveNumbers()}catch{}
    $r.Font.Name='宋体';$r.Font.NameFarEast='宋体';$r.Font.NameAscii='SimSun';$r.Font.NameOther='SimSun'
    $r.Font.Bold=-1;$r.ParagraphFormat.Alignment=$wdAlignCenter
  }
  $size=26
  while($size -ge 14){$tr.Font.Size=$size;$sr.Font.Size=$size;$doc.Repaginate();if($tr.ComputeStatistics($wdLines)-le 1){break};$size--}
  $doc.Sections.Item(1).PageSetup.VerticalAlignment=$wdVerticalCenter
  return $size
}

function Set-Header($doc,[string]$title){
  foreach($section in $doc.Sections){
    $section.PageSetup.OddAndEvenPagesHeaderFooter=0;$section.PageSetup.DifferentFirstPageHeaderFooter=0
    $h=$section.Headers.Item($wdHeaderPrimary);$h.LinkToPrevious=$false
    $r=$h.Range;$r.Text=$title+' V1.0'+"`t";$r.Font.Name='宋体';$r.Font.NameFarEast='宋体';$r.Font.Size=9;$r.Font.Bold=0
    $r.ParagraphFormat.Alignment=$wdAlignLeft;$r.ParagraphFormat.TabStops.ClearAll()
    $width=$section.PageSetup.PageWidth-$section.PageSetup.LeftMargin-$section.PageSetup.RightMargin
    [void]$r.ParagraphFormat.TabStops.Add($width,$wdAlignTabRight,$wdTabLeaderSpaces)
    $pr=$h.Range.Duplicate;if($pr.End -gt $pr.Start){$pr.End--};$pr.Collapse($wdCollapseEnd)
    [void]$doc.Fields.Add($pr,$wdFieldPage);$h.PageNumbers.RestartNumberingAtSection=$false
  }
}

$word=New-Object -ComObject Word.Application;$word.Visible=$false;$word.DisplayAlerts=0
try{
  foreach($folder in Get-ChildItem -LiteralPath $BaseDir -Directory|Sort-Object Name){
    $docx=Get-ChildItem -LiteralPath $folder.FullName -File -Filter '*说明.docx'|Select-Object -First 1
    if(-not $docx){continue};$txt=Get-ChildItem -LiteralPath $folder.FullName -File -Filter '*.txt'|Select-Object -First 1
    $title=if($txt){$txt.BaseName}else{$folder.Name};$doc=$word.Documents.Open($docx.FullName,$false,$false)
    try{
      $added=$false;$size=$null
      if($AddCoverTitles -contains $title){$size=Add-Cover $doc $title;$added=$true}
      if(-not $SkipHeaderNormalization){Set-Header $doc $title}
      $doc.Fields.Update()|Out-Null;$doc.Save();$pdf=[IO.Path]::ChangeExtension($docx.FullName,'.pdf')
      $doc.ExportAsFixedFormat($pdf,$wdExportPDF)
      [pscustomobject]@{Name=$title;CoverAdded=$added;FontSize=$size;HeaderNormalized=(-not $SkipHeaderNormalization);Status='OK'}
    }finally{$doc.Close($false);[Runtime.InteropServices.Marshal]::ReleaseComObject($doc)|Out-Null}
  }
}finally{$word.Quit();[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null}


param(
  [Parameter(Mandatory=$true)][string]$OriginalDocx,
  [Parameter(Mandatory=$true)][string]$RevisedDocx,
  [Parameter(Mandatory=$true)][string]$CommentsCsv,
  [Parameter(Mandatory=$true)][string]$OutputDocx
)
$ErrorActionPreference = "Stop"
$wdFormatDocumentDefault = 16
$wdFindContinue = 1

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $orig = $word.Documents.Open((Resolve-Path $OriginalDocx))
  $rev  = $word.Documents.Open((Resolve-Path $RevisedDocx))
  $orig.Compare($rev)
  $comp = $word.ActiveDocument

  if (Test-Path $CommentsCsv) {
    $rows = Import-Csv -Path $CommentsCsv
    foreach ($row in $rows) {
      $anchor = $row.ancre_textuelle
      $note   = $row.commentaire
      if ([string]::IsNullOrWhiteSpace($anchor) -or [string]::IsNullOrWhiteSpace($note)) { continue }
      $rng = $comp.Content
      $find = $rng.Find
      $find.ClearFormatting()
      $find.Text = $anchor
      $find.Forward = $true
      $find.Wrap = $wdFindContinue
      if ($find.Execute()) { $comp.Comments.Add($rng, $note) | Out-Null }
    }
  }

  $outDir = Split-Path -Parent $OutputDocx
  if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
  $comp.SaveAs([ref]$OutputDocx, [ref]$wdFormatDocumentDefault)
} finally {
  foreach ($d in @($word.Documents)) { try { $d.Close($false) } catch {} }
  $word.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}


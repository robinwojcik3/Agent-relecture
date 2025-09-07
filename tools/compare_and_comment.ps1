param(
  [Parameter(Mandatory=$true)][string]$OriginalDocx,
  [Parameter(Mandatory=$true)][string]$RevisedDocx,
  [Parameter(Mandatory=$true)][string]$CommentsCsv,
  [Parameter(Mandatory=$true)][string]$OutputDocx
)
$ErrorActionPreference = "Stop"
$wdFormatDocumentDefault = 16
$wdFindContinue = 1

# Resolve input paths early with LiteralPath (handles spaces/accents)
$origPath = (Resolve-Path -LiteralPath $OriginalDocx).Path
$revPath  = (Resolve-Path -LiteralPath $RevisedDocx).Path
$csvPath  = $null
if (Test-Path -LiteralPath $CommentsCsv) { $csvPath = (Resolve-Path -LiteralPath $CommentsCsv).Path }

# Stage files into a short TEMP folder to avoid MAX_PATH/encoding issues with Office COM
$stage = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("word_compare_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $stage | Out-Null
$stageOrig = Join-Path $stage 'orig.docx'
$stageRev  = Join-Path $stage 'rev.docx'
$stageCsv  = Join-Path $stage 'comments.csv'
Copy-Item -LiteralPath $origPath -Destination $stageOrig -Force
Copy-Item -LiteralPath $revPath  -Destination $stageRev  -Force
if ($csvPath) { Copy-Item -LiteralPath $csvPath -Destination $stageCsv -Force }

$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  try {
    $orig = $word.Documents.Open($stageOrig)
    $rev  = $word.Documents.Open($stageRev)
    $orig.Compare($rev)
    $comp = $word.ActiveDocument

    if (Test-Path -LiteralPath $stageCsv) {
      $rows = Import-Csv -LiteralPath $stageCsv
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

    # Save to a short temp path first, then copy to target (avoids Office SaveAs long path limits)
    $stageOut = Join-Path $stage 'out.docx'
    $comp.SaveAs([ref]$stageOut, [ref]$wdFormatDocumentDefault)
  } finally {
    foreach ($d in @($word.Documents)) { try { $d.Close($false) } catch {} }
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
}
catch {
  # Emit a clear, actionable error to stderr so Python can display it
  $errMsg = ($_.Exception | Out-String)
  if ([string]::IsNullOrWhiteSpace($errMsg)) { $errMsg = ($_ | Out-String) }
  Write-Error "Echec dans compare_and_comment.ps1 : $errMsg"
  Write-Error "Chemins: `n  Original = $origPath`n  Revised  = $revPath`n  CSV      = $csvPath`n  Output   = $OutputDocx"
  exit 1
}

# Ensure destination directory exists and copy result
$outDir = Split-Path -Parent $OutputDocx
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
Copy-Item -LiteralPath $stageOut -Destination $OutputDocx -Force | Out-Null

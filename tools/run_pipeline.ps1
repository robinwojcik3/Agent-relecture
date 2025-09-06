$ErrorActionPreference = "Stop"
$root   = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo   = Split-Path -Parent $root
$work   = Join-Path $repo "work"
$output = Join-Path $repo "output"

$sessionFile = Join-Path $work "session.json"
if (-not (Test-Path $sessionFile)) { throw "work/session.json manquant. Lancez Start.py d’abord." }
$session = Get-Content $sessionFile | ConvertFrom-Json
$origRel = $session.original_docx
$orig    = Join-Path $repo $origRel
$revMd   = Join-Path $work "rapport_revise.md"
$revDocx = Join-Path $work "rapport_revise.docx"
$csv     = Join-Path $work "commentaires.csv"

if (-not (Test-Path $revMd)) { throw "Manque work/rapport_revise.md (produit par l’agent)." }
if (-not (Test-Path $csv))   { Write-Host "Avertissement: work/commentaires.csv introuvable. Continuer sans commentaires." }

# Convertir MD -> DOCX
pandoc $revMd -o $revDocx

# Construire nom de sortie
$base = [System.IO.Path]::GetFileNameWithoutExtension($orig)
$outDocx = Join-Path $output ($base + "_AI_suivi+commentaires.docx")

# Compare + commentaires
pwsh -File (Join-Path $root "compare_and_comment.ps1") `
     -OriginalDocx $orig `
     -RevisedDocx  $revDocx `
     -CommentsCsv  $csv `
     -OutputDocx   $outDocx

Write-Host "OK -> $outDocx"


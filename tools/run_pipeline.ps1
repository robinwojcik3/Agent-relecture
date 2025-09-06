Param(
  [switch]$Cleanup
)

$ErrorActionPreference = "Stop"
$root   = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo   = Split-Path -Parent $root
$work   = Join-Path $repo "work"
$output = Join-Path $repo "output"

$sessionFile = Join-Path $work "session.json"
if (-not (Test-Path $sessionFile)) { throw "work/session.json manquant. Lancez Start.py d'abord." }
$session = Get-Content $sessionFile | ConvertFrom-Json
$origRel = $session.original_docx
$orig    = Join-Path $repo $origRel
$revMd   = Join-Path $work "rapport_revise.md"
$revDocx = Join-Path $work "rapport_revise.docx"
$csv     = Join-Path $work "commentaires.csv"

# Si la session fournit un dossier de sortie explicite, l'utiliser
if ($session.PSObject.Properties.Name -contains 'output_dir') {
  $userOut = [string]$session.output_dir
  if ($userOut) {
    if ([System.IO.Path]::IsPathRooted($userOut)) { $output = $userOut }
    else { $output = Join-Path $repo $userOut }
  }
}
if (-not (Test-Path $output)) { New-Item -ItemType Directory -Force -Path $output | Out-Null }

if (-not (Test-Path $revMd)) { throw "Manque work/rapport_revise.md (produit par l'agent)." }
if (-not (Test-Path $csv))   { Write-Host "Avertissement: work/commentaires.csv introuvable. Continuer sans commentaires." }

# Convertir MD -> DOCX
$pandocCmd = $null
try { $pandocCmd = (Get-Command pandoc -ErrorAction Stop).Source } catch {}
if (-not $pandocCmd) {
  $candidate = Join-Path $env:ProgramFiles "Pandoc/pandoc.exe"
  if (Test-Path $candidate) { $pandocCmd = $candidate }
}
if (-not $pandocCmd) {
  $candidate = Join-Path $env:LOCALAPPDATA "Pandoc/pandoc.exe"
  if (Test-Path $candidate) { $pandocCmd = $candidate }
}
if (-not $pandocCmd) { throw "Pandoc introuvable. Exécutez tools/setup_tools.ps1 puis relancez." }
& $pandocCmd $revMd -o $revDocx

# Construire nom de sortie
$base = [System.IO.Path]::GetFileNameWithoutExtension($orig)
$outDocx = Join-Path $output ($base + "_AI_suivi+commentaires.docx")

# Compare + commentaires
& (Join-Path $root "compare_and_comment.ps1") `
   -OriginalDocx $orig `
   -RevisedDocx  $revDocx `
   -CommentsCsv  $csv `
   -OutputDocx   $outDocx

if (-not (Test-Path $outDocx)) { throw "Echec de la génération du DOCX final: $outDocx introuvable." }
Write-Host "OK -> $outDocx"

# Nettoyage optionnel des fichiers intermédiaires
if ($Cleanup) {
  foreach ($p in @($revMd, $revDocx, $csv)) {
    try { if (Test-Path $p) { Remove-Item -Force -ErrorAction SilentlyContinue $p } } catch {}
  }
}

Param(
  [Parameter(Mandatory=$true)]
  [ValidateSet("offre","diagnostic","impacts","mesures")]
  [string]$Mode
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "..\modes\$Mode\refs"
$dst = Join-Path $root "..\..\work\refs_txt\$Mode"
New-Item -Force -ItemType Directory -Path $dst | Out-Null
If (-not (Test-Path $src)) { Write-Host "Pas de refs pour $Mode."; exit 0 }
Get-ChildItem $src -Filter *.pdf | ForEach-Object {
  $out = Join-Path $dst ($_.BaseName + ".txt")
  & pdftotext -layout $_.FullName $out
}
Write-Host "Références extraites vers $dst"


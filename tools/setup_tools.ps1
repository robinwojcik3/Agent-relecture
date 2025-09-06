Param([switch]$Force)
$ErrorActionPreference = "Stop"

function Ensure-App($name, $id) {
  if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
    Write-Host "Installing $name..."
    winget install -e --id $id --silent
  } else {
    Write-Host "$name found."
  }
}

Ensure-App -name "pandoc" -id "JohnMacFarlane.Pandoc"
if (-not (Get-Command "pdftotext" -ErrorAction SilentlyContinue)) {
  Write-Host "Installing Poppler..."
  winget install -e --id "oschwartz10612.Poppler" --silent
} else {
  Write-Host "Poppler found."
}
Write-Host "Done."


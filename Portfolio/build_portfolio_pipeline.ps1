param(
  [switch]$SkipPdf,
  [switch]$SkipPowerPoint
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

if (-not $SkipPdf) {
  foreach ($tex in @("portfolio_from_ppt_images.tex", "portfolio_from_ppt_images_a4.tex")) {
    lualatex -interaction=batchmode -halt-on-error -file-line-error $tex
    if ($LASTEXITCODE -ne 0) { throw "LaTeX failed for $tex" }
    lualatex -interaction=batchmode -halt-on-error -file-line-error $tex
    if ($LASTEXITCODE -ne 0) { throw "LaTeX failed for $tex" }
  }
}

if (-not $SkipPowerPoint) {
  powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "export_editable_powerpoint.ps1") -Format both
  if ($LASTEXITCODE -ne 0) { throw "PowerPoint export failed" }
}

Write-Host "Portfolio pipeline complete."

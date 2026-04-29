param(
  [switch]$SkipPdf,
  [switch]$SkipPowerPoint
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$pdfPairs = @(
  @{
    Local = Join-Path $Root "portfolio_from_ppt_images.pdf"
    Submission = Join-Path ([Environment]::GetFolderPath("Desktop")) "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent.pdf"
  },
  @{
    Local = Join-Path $Root "portfolio_from_ppt_images_a4.pdf"
    Submission = Join-Path ([Environment]::GetFolderPath("Desktop")) "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent_A4.pdf"
  }
)

$pdfPaths = @($pdfPairs | ForEach-Object { $_.Local; $_.Submission })
$viewerScript = Join-Path $Root "refresh_portfolio_pdf_viewers.ps1"

if (-not $SkipPdf) {
  powershell -NoProfile -ExecutionPolicy Bypass -File $viewerScript -Close -Paths $pdfPaths
  if ($LASTEXITCODE -ne 0) { throw "Could not close open portfolio PDF viewers" }
}

if (-not $SkipPdf) {
  foreach ($tex in @("portfolio_from_ppt_images.tex", "portfolio_from_ppt_images_a4.tex")) {
    lualatex -interaction=batchmode -halt-on-error -file-line-error $tex
    if ($LASTEXITCODE -ne 0) { throw "LaTeX failed for $tex" }
    lualatex -interaction=batchmode -halt-on-error -file-line-error $tex
    if ($LASTEXITCODE -ne 0) { throw "LaTeX failed for $tex" }
  }

  foreach ($pair in $pdfPairs) {
    if (Test-Path -LiteralPath $pair.Local) {
      Copy-Item -LiteralPath $pair.Local -Destination $pair.Submission -Force
    }
  }
}

if (-not $SkipPowerPoint) {
  powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "export_editable_powerpoint.ps1") -Format both
  if ($LASTEXITCODE -ne 0) { throw "PowerPoint export failed" }
}

Write-Host "Portfolio pipeline complete."

if (-not $SkipPdf) {
  $submissionPaths = @($pdfPairs | ForEach-Object { $_.Submission })
  powershell -NoProfile -ExecutionPolicy Bypass -File $viewerScript -Open -Paths $submissionPaths
}

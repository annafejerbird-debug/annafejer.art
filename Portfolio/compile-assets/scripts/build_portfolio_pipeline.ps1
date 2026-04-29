param(
  [ValidateSet("wide", "a4", "both")]
  [string]$Format = "both",
  [switch]$SkipPdf,
  [switch]$NoOpen
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = (Resolve-Path (Join-Path $ScriptDir "..\..")).Path
$OutputRoot = Join-Path $Root "Output"
$BuildDir = Join-Path $OutputRoot "build"
$PageRoot = Join-Path $OutputRoot "pages"
$ViewerScript = Join-Path $ScriptDir "refresh_portfolio_pdf_viewers.ps1"
$SplitScript = Join-Path $ScriptDir "split_portfolio_pages.ps1"
$AuditScript = Join-Path $ScriptDir "audit_portfolio_inclusion.py"

Set-Location $Root
New-Item -ItemType Directory -Force -Path $OutputRoot, $BuildDir, $PageRoot | Out-Null

$Targets = @(
  @{
    Key = "wide"
    Tex = "portfolio_from_ppt_images.tex"
    BuildPdf = Join-Path $BuildDir "portfolio_from_ppt_images.pdf"
    FinalPdf = Join-Path $OutputRoot "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent.pdf"
    PageDir = Join-Path $PageRoot "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent"
    PagePrefix = "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent"
    LegacyPaths = @(
      (Join-Path $Root "portfolio_from_ppt_images.pdf"),
      (Join-Path ([Environment]::GetFolderPath("Desktop")) "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent.pdf")
    )
  },
  @{
    Key = "a4"
    Tex = "portfolio_from_ppt_images_a4.tex"
    BuildPdf = Join-Path $BuildDir "portfolio_from_ppt_images_a4.pdf"
    FinalPdf = Join-Path $OutputRoot "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent_A4.pdf"
    PageDir = Join-Path $PageRoot "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent_A4"
    PagePrefix = "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent_A4"
    LegacyPaths = @(
      (Join-Path $Root "portfolio_from_ppt_images_a4.pdf"),
      (Join-Path ([Environment]::GetFolderPath("Desktop")) "Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent_A4.pdf")
    )
  }
)

if ($Format -ne "both") {
  $Targets = @($Targets | Where-Object { $_.Key -eq $Format })
}

if (-not $SkipPdf) {
  python $AuditScript --write --sync-tex
  if ($LASTEXITCODE -ne 0) { throw "Portfolio metadata sync failed" }

  $closePaths = @($Targets | ForEach-Object { $_.FinalPdf; $_.BuildPdf; $_.LegacyPaths })
  powershell -NoProfile -ExecutionPolicy Bypass -File $ViewerScript -Close -Paths $closePaths
  if ($LASTEXITCODE -ne 0) { throw "Could not close open portfolio PDF viewers" }

  foreach ($target in $Targets) {
    foreach ($pass in 1..2) {
      lualatex -interaction=batchmode -halt-on-error -file-line-error "-output-directory=$BuildDir" $target.Tex
      if ($LASTEXITCODE -ne 0) { throw "LaTeX failed for $($target.Tex) on pass $pass" }
    }

    Copy-Item -LiteralPath $target.BuildPdf -Destination $target.FinalPdf -Force

    powershell -NoProfile -ExecutionPolicy Bypass -File $SplitScript `
      -PdfPath $target.FinalPdf `
      -OutputDir $target.PageDir `
      -Prefix $target.PagePrefix
    if ($LASTEXITCODE -ne 0) { throw "PDF page splitting failed for $($target.FinalPdf)" }
  }

  python $AuditScript --write --require-output
  if ($LASTEXITCODE -ne 0) { throw "Portfolio inclusion audit failed" }
}

Write-Host "Portfolio pipeline complete."
Write-Host "Outputs: $OutputRoot"

if (-not $SkipPdf -and -not $NoOpen) {
  $openPaths = @($Targets | ForEach-Object { $_.FinalPdf })
  powershell -NoProfile -ExecutionPolicy Bypass -File $ViewerScript -Open -Paths $openPaths
}

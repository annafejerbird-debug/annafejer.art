param(
  [Parameter(Mandatory = $true)]
  [string]$PdfPath,
  [Parameter(Mandatory = $true)]
  [string]$OutputDir,
  [string]$Prefix
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $PdfPath)) {
  throw "PDF not found: $PdfPath"
}

$ResolvedPdf = (Resolve-Path -LiteralPath $PdfPath).Path
if ([string]::IsNullOrWhiteSpace($Prefix)) {
  $Prefix = [System.IO.Path]::GetFileNameWithoutExtension($ResolvedPdf)
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$ResolvedOutputDir = (Resolve-Path -LiteralPath $OutputDir).Path

Get-ChildItem -LiteralPath $ResolvedOutputDir -Filter "$Prefix-page-*.pdf" -File |
  Remove-Item -Force

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "split_portfolio_pages.py"
$PageCount = & python $PythonScript $ResolvedPdf $ResolvedOutputDir $Prefix
if ($LASTEXITCODE -ne 0) {
  throw "pypdf page splitting failed for $ResolvedPdf. Install with: python -m pip install pypdf"
}

Write-Host "Split $ResolvedPdf into $PageCount page PDFs at $ResolvedOutputDir"

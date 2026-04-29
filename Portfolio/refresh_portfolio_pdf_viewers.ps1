param(
  [switch]$Close,
  [switch]$Open,
  [string[]]$Paths = @()
)

$ErrorActionPreference = "Stop"

function Get-PortfolioPdfTitles {
  param([string[]]$InputPaths)

  $titles = New-Object System.Collections.Generic.List[string]
  foreach ($path in $InputPaths) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    $leaf = Split-Path -Leaf $path
    if (-not [string]::IsNullOrWhiteSpace($leaf)) {
      $titles.Add($leaf)
      $titles.Add([System.IO.Path]::GetFileNameWithoutExtension($leaf))
    }
  }
  $titles | Sort-Object -Unique
}

function Close-PortfolioPdfWindows {
  param([string[]]$InputPaths)

  $titles = @(Get-PortfolioPdfTitles -InputPaths $InputPaths)
  if ($titles.Count -eq 0) { return }

  $viewerNames = @(
    "Acrobat",
    "AcroRd32",
    "SumatraPDF",
    "FoxitPDFReader",
    "FoxitReader",
    "msedge",
    "chrome",
    "firefox"
  )
  $forceCloseNames = @(
    "Acrobat",
    "AcroRd32",
    "SumatraPDF",
    "FoxitPDFReader",
    "FoxitReader"
  )

  $matches = @(
    Get-Process -ErrorAction SilentlyContinue |
      Where-Object {
        $_.MainWindowHandle -ne 0 -and
        $viewerNames -contains $_.ProcessName -and
        ($title = $_.MainWindowTitle) -and
        ($titles | Where-Object { $title -like "*$_*" })
      }
  )

  foreach ($process in $matches) {
    try {
      [void]$process.CloseMainWindow()
    } catch {
      Write-Warning "Could not ask $($process.ProcessName) to close: $($_.Exception.Message)"
    }
  }

  Start-Sleep -Milliseconds 1200

  foreach ($process in $matches) {
    try {
      $fresh = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
      if ($fresh -and -not $fresh.HasExited -and $forceCloseNames -contains $fresh.ProcessName) {
        Stop-Process -Id $fresh.Id -Force -ErrorAction SilentlyContinue
      }
    } catch {
      Write-Warning "Could not force-close $($process.ProcessName): $($_.Exception.Message)"
    }
  }
}

function Open-PortfolioPdfs {
  param([string[]]$InputPaths)

  foreach ($path in $InputPaths) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    if (Test-Path -LiteralPath $path) {
      Start-Process -FilePath $path
    }
  }
}

if ($Close) {
  Close-PortfolioPdfWindows -InputPaths $Paths
}

if ($Open) {
  Open-PortfolioPdfs -InputPaths $Paths
}

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
    }
  }
  $titles | Sort-Object -Unique
}

function Close-PortfolioPdfWindows {
  param([string[]]$InputPaths)

  $titles = @(Get-PortfolioPdfTitles -InputPaths $InputPaths)
  if ($titles.Count -eq 0) { return }
  $resolvedPaths = @(
    $InputPaths |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      ForEach-Object {
        try {
          if (Test-Path -LiteralPath $_) { (Resolve-Path -LiteralPath $_).Path } else { $_ }
        } catch {
          $_
        }
      }
  )

  $viewerNames = @(
    "Acrobat",
    "AcroRd32",
    "SumatraPDF",
    "FoxitPDFReader",
    "FoxitReader",
    "okular",
    "msedge",
    "chrome",
    "firefox"
  )
  $forceCloseNames = @(
    "Acrobat",
    "AcroRd32",
    "SumatraPDF",
    "FoxitPDFReader",
    "FoxitReader",
    "okular"
  )

  function Get-MatchingViewerProcesses {
    @(
      Get-Process -ErrorAction SilentlyContinue |
        Where-Object {
          $_.MainWindowHandle -ne 0 -and
          $viewerNames -contains $_.ProcessName -and
          ($title = $_.MainWindowTitle) -and
          ($titles | Where-Object { $title -like "*$_*" })
        }
    )
  }

  function Get-MatchingViewerProcessIdsByCommandLine {
    @(
      Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
          $processName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
          $forceCloseNames -contains $processName -and
          ($commandLine = $_.CommandLine) -and
          (
            ($titles | Where-Object { $commandLine -like "*$_*" }) -or
            ($resolvedPaths | Where-Object { $commandLine -like "*$_*" })
          )
        } |
        ForEach-Object { [int]$_.ProcessId }
    )
  }

  $seenProcessIds = New-Object System.Collections.Generic.HashSet[int]
  foreach ($attempt in 1..3) {
    $matches = @(Get-MatchingViewerProcesses)
    if ($matches.Count -eq 0) { break }
    foreach ($process in $matches) {
      [void]$seenProcessIds.Add($process.Id)
      try {
        [void]$process.CloseMainWindow()
      } catch {
        Write-Warning "Could not ask $($process.ProcessName) to close: $($_.Exception.Message)"
      }
    }
    Start-Sleep -Milliseconds 800
  }

  foreach ($process in @(Get-MatchingViewerProcesses)) {
    [void]$seenProcessIds.Add($process.Id)
  }
  foreach ($processId in @(Get-MatchingViewerProcessIdsByCommandLine)) {
    [void]$seenProcessIds.Add($processId)
  }

  foreach ($processId in $seenProcessIds) {
    try {
      $fresh = Get-Process -Id $processId -ErrorAction SilentlyContinue
      if ($fresh -and -not $fresh.HasExited -and $forceCloseNames -contains $fresh.ProcessName) {
        Stop-Process -Id $fresh.Id -Force -ErrorAction SilentlyContinue
      }
    } catch {
      Write-Warning "Could not force-close viewer process ${processId}: $($_.Exception.Message)"
    }
  }
}

function Open-PortfolioPdfs {
  param([string[]]$InputPaths)

  $forceCloseNames = @(
    "Acrobat",
    "AcroRd32",
    "SumatraPDF",
    "FoxitPDFReader",
    "FoxitReader",
    "okular"
  )

  function Close-DuplicateViewerWindows {
    param([string]$InputPath)

    $leaf = Split-Path -Leaf $InputPath
    $resolved = try {
      if (Test-Path -LiteralPath $InputPath) { (Resolve-Path -LiteralPath $InputPath).Path } else { $InputPath }
    } catch {
      $InputPath
    }

    $matches = @(
      Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
          $processName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
          $forceCloseNames -contains $processName -and
          ($commandLine = $_.CommandLine) -and
          ($commandLine -like "*$leaf*" -or $commandLine -like "*$resolved*")
        } |
        Sort-Object CreationDate -Descending
    )

    if ($matches.Count -le 1) { return }
    foreach ($match in @($matches | Select-Object -Skip 1)) {
      Stop-Process -Id ([int]$match.ProcessId) -Force -ErrorAction SilentlyContinue
    }
  }

  foreach ($path in $InputPaths) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    if (Test-Path -LiteralPath $path) {
      Invoke-Item -LiteralPath $path
      Start-Sleep -Milliseconds 2500
      Close-DuplicateViewerWindows -InputPath $path
    }
  }
}

if ($Close) {
  Close-PortfolioPdfWindows -InputPaths $Paths
}

if ($Open) {
  Open-PortfolioPdfs -InputPaths $Paths
}

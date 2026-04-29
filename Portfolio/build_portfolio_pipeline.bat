@echo off
setlocal

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0compile-assets\scripts\build_portfolio_pipeline.ps1"
if errorlevel 1 (
  echo.
  echo Portfolio pipeline failed.
  pause
  exit /b 1
)

echo.
echo Portfolio pipeline complete.
pause

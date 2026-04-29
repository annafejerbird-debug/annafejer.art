@echo off
setlocal

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0compile-assets\scripts\build_portfolio_pipeline.ps1" -Format wide
if errorlevel 1 (
  echo.
  echo Portfolio PDF compilation failed.
  pause
  exit /b 1
)

echo.
echo Portfolio PDF compilation complete.
pause
exit /b 0

@echo off
setlocal

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0compile-assets\scripts\build_portfolio_pipeline.ps1"
if errorlevel 1 (
  echo.
  echo A4 portfolio PDF compilation failed.
  pause
  exit /b 1
)

echo.
echo A4 portfolio PDF compilation complete.
pause
exit /b 0

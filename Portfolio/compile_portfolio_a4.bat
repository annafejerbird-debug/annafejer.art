@echo off
setlocal

cd /d "%~dp0"

set "TEX_FILE=portfolio_from_ppt_images_a4.tex"
set "PDF_FILE=portfolio_from_ppt_images_a4.pdf"
set "SUBMISSION_PDF=%USERPROFILE%\Desktop\Fejer_Anna_88398_Mappe_BildendeKunst-Absolvent_A4.pdf"

where lualatex >nul 2>nul
if %ERRORLEVEL%==0 (
  lualatex -interaction=nonstopmode -halt-on-error -file-line-error "%TEX_FILE%"
  if errorlevel 1 goto failed
  lualatex -interaction=nonstopmode -halt-on-error -file-line-error "%TEX_FILE%"
  if errorlevel 1 goto failed
  goto done
)

where latexmk >nul 2>nul
if %ERRORLEVEL%==0 (
  latexmk -lualatex -interaction=nonstopmode -halt-on-error -file-line-error "%TEX_FILE%"
  if errorlevel 1 goto failed
  goto done
)

where xelatex >nul 2>nul
if %ERRORLEVEL%==0 (
  xelatex -interaction=nonstopmode -halt-on-error -file-line-error "%TEX_FILE%"
  if errorlevel 1 goto failed
  xelatex -interaction=nonstopmode -halt-on-error -file-line-error "%TEX_FILE%"
  if errorlevel 1 goto failed
  goto done
)

echo Could not find lualatex, latexmk, or xelatex on PATH.
echo Install MiKTeX or TeX Live, then run this file again.
pause
exit /b 1

:done
if exist "%PDF_FILE%" (
  copy /Y "%PDF_FILE%" "%SUBMISSION_PDF%" >nul
  echo.
  echo Done: %PDF_FILE%
  echo Submission copy: %SUBMISSION_PDF%
) else (
  echo.
  echo Done, but %PDF_FILE% was not found to rename.
)
pause
exit /b 0

:failed
echo.
echo LaTeX compilation failed. Check the .log file above for details.
pause
exit /b 1

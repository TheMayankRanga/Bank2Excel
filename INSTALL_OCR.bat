@echo off
setlocal
title Bank2Excel OCR Installer
echo.
echo ==========================================================
echo                 BANK2EXCEL OCR SETUP
echo ==========================================================
echo.

where winget >nul 2>&1
if errorlevel 1 (
  echo ERROR: Windows Package Manager (winget) is not available.
  echo.
  echo Use the official Windows installer instead:
  echo https://github.com/UB-Mannheim/tesseract/wiki
  echo.
  pause
  exit /b 1
)

echo Installing Tesseract OCR...
echo.
winget install --id UB-Mannheim.TesseractOCR -e --source winget --accept-source-agreements --accept-package-agreements

echo.
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
  echo SUCCESS: Tesseract was installed.
  "C:\Program Files\Tesseract-OCR\tesseract.exe" --version
  echo.
  echo Close this window, then run START_HERE.bat again.
  pause
  exit /b 0
)

if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" (
  echo SUCCESS: Tesseract was installed.
  "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" --version
  echo.
  echo Close this window, then run START_HERE.bat again.
  pause
  exit /b 0
)

echo.
echo Tesseract was not detected after the install command.
echo Run this command manually and send me the output:
echo.
echo winget search tesseract
echo.
pause

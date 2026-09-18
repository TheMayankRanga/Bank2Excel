@echo off
setlocal
cd /d "%~dp0"
title Bank2Excel - Bank Statement Converter
echo.
echo ==========================================================
echo                 BANK2EXCEL
 echo               BANK STATEMENT CONVERTER
echo ==========================================================
echo.
where py >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python 3.11+ was not found.
  echo Install Python, then run this file again.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  echo Creating local environment...
  py -3 -m venv .venv
  if errorlevel 1 (
    echo ERROR: Could not create the environment.
    pause
    exit /b 1
  )
)
echo Installing/updating required Python packages...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
  echo ERROR: Package installation failed. Check internet and try again.
  pause
  exit /b 1
)
echo.
echo Starting Bank2Excel...
echo Keep this window open.
start "" http://127.0.0.1:8000
.venv\Scripts\python.exe app.py
pause

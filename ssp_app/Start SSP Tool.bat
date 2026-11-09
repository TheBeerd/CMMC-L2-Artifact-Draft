@echo off
title CMMC L2 SSP Drafting Tool
cd /d "%~dp0"

set PY=
where py >nul 2>&1 && set PY=py
if "%PY%"=="" (
    where python >nul 2>&1 && set PY=python
)
if "%PY%"=="" (
    echo ERROR: Python was not found.
    echo Run "Setup (run once).bat" first, or install Python from https://python.org
    echo.
    pause
    exit /b 1
)

echo.
echo   Starting the CMMC L2 SSP Drafting Tool...
echo   Your browser will open automatically.
echo.
echo   KEEP THIS WINDOW OPEN while you work.
echo   To stop: close this window, or press Ctrl+C.
echo.

%PY% app.py

echo.
echo   The tool has stopped.
pause

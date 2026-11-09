@echo off
title CMMC L2 SSP Drafting Tool - Setup
echo ==========================================
echo   CMMC L2 SSP Drafting Tool
echo   One-time setup
echo ==========================================
echo.

cd /d "%~dp0"

REM Find a working Python. The py launcher is preferred on Windows.
set PY=
where py >nul 2>&1 && set PY=py
if "%PY%"=="" (
    where python >nul 2>&1 && set PY=python
)
if "%PY%"=="" (
    echo ERROR: Python was not found.
    echo.
    echo Install Python 3.10 or newer from https://python.org
    echo IMPORTANT: tick "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo Using: %PY%
%PY% --version
echo.
echo Installing required packages...
echo.
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo ==========================================
    echo   Setup FAILED
    echo ==========================================
    echo If the error mentions Pillow and you are on a very new
    echo Python version, install Python 3.12 and run this again.
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   Setup complete.
echo.
echo   You can now start the tool by
echo   double-clicking:  Start SSP Tool.bat
echo ==========================================
echo.
pause

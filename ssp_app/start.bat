@echo off
echo ==========================================
echo   CMMC L2 SSP Drafting Tool
echo ==========================================
echo.
echo Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download Python 3.10+ from https://python.org
    pause
    exit /b 1
)

echo Installing/updating dependencies...
pip install -r requirements.txt -q

echo.
echo Starting application...
echo Open your browser to: http://127.0.0.1:5000
echo.
echo Press Ctrl+C to stop the server.
echo.
python app.py
pause

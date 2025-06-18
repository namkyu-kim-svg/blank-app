@echo off
chcp 65001 >nul
title Business Trip Report System
color 0A

echo ====================================================
echo    Business Trip Report Automation System v1.0
echo ====================================================
echo.

REM Move to current directory
cd /d "%~dp0"

REM Check Python installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed!
    echo.
    echo Please install Python first:
    echo 1. Download from https://www.python.org/downloads/
    echo 2. Check "Add Python to PATH" during installation
    echo 3. Run this file again after installation
    echo.
    echo Or contact IT team for Python installation.
    echo.
    pause
    exit /b 1
)

echo [OK] Python found
echo.

REM Install required packages
echo [INFO] Checking and installing required packages...
pip install --only-binary=all -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Package installation failed.
    echo Please check your internet connection or contact IT team.
    pause
    exit /b 1
)

echo [OK] Packages ready
echo.

REM Get network IP
for /f "tokens=2 delims=:" %%i in ('ipconfig ^| findstr /c:"IPv4"') do set ip=%%i
set ip=%ip: =%

echo [INFO] Starting server...
echo.
echo ================================================
echo   Connection Information
echo ================================================
echo  Local:    http://localhost:8501
echo  Network:  http://%ip%:8501
echo.
echo  Browser will open automatically
echo  Press Ctrl+C to stop the server
echo  Other computers can access via network address
echo ================================================
echo.

REM Run Streamlit with network access
streamlit run main.py --server.address 0.0.0.0 --server.port 8501 --browser.serverAddress localhost

echo.
echo Business Trip Report System stopped.
echo You can run this file again anytime to restart.
pause

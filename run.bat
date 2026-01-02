@echo off
echo Marketing Dashboard Automator v2.0
echo ===================================
echo.

REM Set Python path (update this to your Python 3.11 installation)
set PYTHON311=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\python.exe

if not exist "%PYTHON311%" (
    echo Python 3.11 not found!
    echo Expected at: %PYTHON311%
    echo.
    echo Please install Python 3.11 first:
    echo 1. Download from https://www.python.org/downloads/release/python-3119/
    echo 2. Install with "Add Python to PATH" checked
    echo 3. Run this script again
    echo.
    pause
    exit /b 1
)

echo Using: %PYTHON311%
%PYTHON311% --version
echo.

REM Navigate to project directory
cd /d "%~dp0"

echo Checking and creating directories...
if not exist "output" mkdir output
if not exist "logs" mkdir logs
if not exist "cache" mkdir cache
if not exist "backups" mkdir backups
if not exist "config" mkdir config

echo Checking configuration...
if not exist "config\dashboard_config.yaml" (
    echo Creating default configuration...
    echo.
    echo Creating basic config file...
    (
        echo dashboard_config:
        echo   social_media:
        echo     keywords:
        echo       - "TF Value-Mart FB Page Wallposts Performance"
        echo       - "TF Value-Mart IG Page Wallposts Performance"
        echo       - "TF Value-Mart Tiktok Video Performance"
        echo     platforms:
        echo       Facebook:
        echo         metrics: ["Reach/Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
        echo       Instagram:
        echo         metrics: ["Reach/Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
        echo       TikTok:
        echo         metrics: ["Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
    ) > "config\dashboard_config.yaml"
)

echo Checking dependencies...
%PYTHON311% -c "import pandas, pptx, plotly, yaml, numpy" 2>nul
if errorlevel 1 (
    echo Installing required packages...
    %PYTHON311% -m pip install --upgrade pip
    %PYTHON311% -m pip install python-pptx pandas openpyxl plotly pyyaml numpy
)

echo.
echo Starting application...
echo.
echo Press Ctrl+C to stop the application
echo.

%PYTHON311% src\main.py

echo.
echo Application closed.
pause
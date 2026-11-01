@echo off
title C-PRO Case Suite Launcher
echo ===================================================
echo             C-PRO Case Suite Launcher
echo ===================================================
echo.

:: Check Python installation
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not found in your PATH.
    echo Please install Python 3.9 or newer from https://www.python.org/
    echo and ensure "Add Python to PATH" is checked during installation.
    pause
    exit /b 1
)

:: Run the suite
echo Starting C-PRO Case Suite...
python "%~dp0cpro_suite_app.py"

if %ERRORLEVEL% neq 0 (
    echo.
    echo [NOTE] If the application failed to start, make sure you have installed dependencies:
    echo        pip install -r requirements.txt
    echo        playwright install chromium
    echo.
    pause
)

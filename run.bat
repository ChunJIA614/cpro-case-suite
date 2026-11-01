@echo off
title C-PRO Case Suite Launcher
echo ===================================================
echo             C-PRO Case Suite Launcher
echo ===================================================
echo.

:: Detect Python launcher or executable
set "PY_CMD="
where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set "PY_CMD=py"
    goto :FOUND
)

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set "PY_CMD=python"
    goto :FOUND
)

where python3 >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set "PY_CMD=python3"
    goto :FOUND
)

echo [ERROR] Python is not found in your PATH.
echo Please install Python 3.9 or newer from https://www.python.org/
echo and ensure "Add Python to PATH" is checked during installation.
echo.
pause
exit /b 1

:FOUND
echo Starting C-PRO Case Suite using %PY_CMD%...
%PY_CMD% "%~dp0cpro_suite_app.py"

if %ERRORLEVEL% neq 0 (
    echo.
    echo [NOTE] If the application failed to start, verify your dependencies:
    echo        %PY_CMD% -m pip install -r requirements.txt
    echo        %PY_CMD% -m playwright install chromium
    echo.
    pause
)

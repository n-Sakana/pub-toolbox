@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    echo Drop Excel file(s) or folder to extract VBA code as text.
    echo Supported: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
echo Starting Extract...
echo Input: %*
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Extract.ps1" %* 2>&1
echo.
echo Exit code: %errorlevel%
pause

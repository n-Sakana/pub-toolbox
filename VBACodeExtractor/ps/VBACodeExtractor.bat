@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    echo Drop an Excel file onto this BAT to extract VBA code.
    echo Supported: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0VBACodeExtractor.ps1" "%~1"
pause

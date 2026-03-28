@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    echo Drop a file onto this BAT to remove VBA password.
    echo Supported: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0VBAPasswordRemover.ps1" "%~1"
pause

@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Analyze.ps1"
    pause
    exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Analyze.ps1" %*
pause

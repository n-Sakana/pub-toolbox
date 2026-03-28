@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    echo Drop a folder or Excel files to generate VBA inventory.
    echo Supported: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Inventory.ps1" %*
pause

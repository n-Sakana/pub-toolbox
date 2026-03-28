@echo off
chcp 65001 >nul
setlocal
echo Environment Probe - Tests EDR/compat patterns via Excel COM
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Probe.ps1"
pause

@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Sample.ps1" %*
pause

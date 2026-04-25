@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Addin.ps1" %*
pause

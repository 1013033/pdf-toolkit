@echo off
setlocal
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0word-to-pdf-native.ps1" %*
echo.
pause

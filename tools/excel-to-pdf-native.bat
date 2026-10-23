@echo off
setlocal
chcp 65001 >nul
title Toolkit AGY - Excel 本機 100%% 原生 PDF 轉檔工具

echo ======================================================================
echo    Toolkit AGY - Microsoft Excel 本機 100%% 原生 PDF 轉檔工具 (Windows)
echo ======================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0excel-to-pdf-native.ps1" %*

echo.
echo 處理完畢，請按任意鍵關閉視窗...
pause >nul
endlocal

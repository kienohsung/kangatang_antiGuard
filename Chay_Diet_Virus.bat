@echo off
:: ============================================================================
:: TRÌNH CHẠY TỰ ĐỘNG DIỆT VIRUS EXCEL KANGATANG (STANDALONE v3.8.1)
:: ============================================================================

title DIET VIRUS EXCEL KANGATANG - EXCEL CLEANER v3.8.1
chcp 65001 >nul

pushd "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Kangatang_Standalone_Scanner.ps1"
popd

echo.
pause

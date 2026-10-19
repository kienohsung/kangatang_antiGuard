@echo off
REM ============================================================================
REM KANGATANGGUARD v3.23.1 - STANDALONE 1-CLICK SCANNER
REM ============================================================================

title KANGATANGGUARD v3.23.1 - STANDALONE SCANNER
chcp 65001 >nul

pushd "%~dp0"

echo ======================================================================
echo    KANGATANGGUARD v3.23.1 - TRINH DIET VIRUS EXCEL NGOAI (STANDALONE)
echo    Khoi chay bo may diet virus doc lap...
echo ======================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Kangatang_Standalone_Scanner.ps1"

popd
echo.
echo ======================================================================
echo Da ket thuc chuong trinh. Nhan phim bat ky de thoat...
pause >nul

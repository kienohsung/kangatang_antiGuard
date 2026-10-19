@echo off
REM ============================================================================
REM KANGATANGGUARD v3.23.1 - STANDALONE SYSTEM & FOLDER CLEANER
REM ============================================================================

title KANGATANGGUARD v3.23.1 - EXCEL CLEANER
chcp 65001 >nul

pushd "%~dp0"
echo ======================================================================
echo    KANGATANGGUARD v3.23.1 - SYSTEM AND FOLDER CLEANER
echo ======================================================================
echo.

if exist "%~dp0addin\Kangatang_Standalone_Scanner.ps1" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0addin\Kangatang_Standalone_Scanner.ps1"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Kangatang_Standalone_Scanner.ps1"
)
popd

echo.
echo ======================================================================
echo Da ket thuc chuong trinh. Nhan phim bat ky de thoat...
pause >nul

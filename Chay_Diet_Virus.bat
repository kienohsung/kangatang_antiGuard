@echo off
REM ============================================================================
REM KANGATANGGUARD v3.17.1 - STANDALONE 1-CLICK SCANNER
REM ============================================================================

title KANGATANGGUARD v3.17.1 - STANDALONE SCANNER
chcp 65001 >nul

pushd "%~dp0"
if exist "%~dp0addin\Kangatang_Standalone_Scanner.ps1" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0addin\Kangatang_Standalone_Scanner.ps1"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Kangatang_Standalone_Scanner.ps1"
)
popd

echo.
pause

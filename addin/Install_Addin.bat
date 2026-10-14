@echo off
setlocal
title KANGATANG GUARD - CAI DAT ADD-IN EXCEL

echo ======================================================================
echo   KANGATANG GUARD - CAI DAT ADD-IN TU DONG QUET VIRUS EXCEL
echo ======================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install_KangatangGuard.ps1"

echo.
echo Nhan phim bat ky de thoat...
pause >nul
exit /b

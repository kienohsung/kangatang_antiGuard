@echo off
chcp 65001 >nul
title Cai dat Add-in KangatangGuard tu May chu LAN
echo ==============================================================================
echo    DANG CAI DAT ADD-IN KANGATANG GUARD TU MAY CHU LAN...
echo ==============================================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install_Client.ps1"

echo.
pause

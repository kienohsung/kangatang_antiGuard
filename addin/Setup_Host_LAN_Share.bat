@echo off
chcp 65001 >nul
title Thiet lap Thu muc Chia se LAN Hub - KangatangGuard
echo ==============================================================================
echo    DANG THIET LAP TRUNG TAM PHAN PHOI LAN HUB KANGATANG GUARD...
echo ==============================================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup_Host_LAN_Share.ps1"

echo.
pause

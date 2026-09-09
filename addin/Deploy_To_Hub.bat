@echo off
chcp 65001 >nul
title Phat hanh KangatangGuard len LAN Hub
echo ==============================================================================
echo    DANG BIEN DICH VA DONG BO PHIEN BAN MOI LEN TRUNG TAM LAN HUB...
echo ==============================================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Deploy_To_Hub.ps1"

echo.
pause

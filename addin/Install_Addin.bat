﻿﻿﻿@echo off
setlocal
chcp 65001 >nul
title KANGATANG GUARD - CAI DAT ADD-IN EXCEL

:: Kiem tra quyen Administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Dang yeu cau quyen Administrator...
    powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList '/c \"\"%~f0\"\"' -Verb RunAs"
    exit /b
)

echo ======================================================================
echo   KANGATANG GUARD - CAI DAT ADD-IN TU DONG QUET VIRUS EXCEL
echo ======================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install_KangatangGuard.ps1"

echo.
echo Nhan phim bat ky de thoat...
pause >nul
exit /b

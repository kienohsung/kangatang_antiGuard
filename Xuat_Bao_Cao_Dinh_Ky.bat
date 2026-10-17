@echo off
title KangatangGuard v3.20.0 - Xuat Bao Cao Dinh Ky & Mo Thu Muc Bao Cao
echo ======================================================================
echo    KANGATANG GUARD - TONG HOP & XUAT BAO CAO QUET DINH KY (v3.20.0)
echo ======================================================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0addin\Sync_ScanReports.ps1"
echo.
if exist "%~dp0Scan_Reports" (
    explorer.exe "%~dp0Scan_Reports"
)
pause

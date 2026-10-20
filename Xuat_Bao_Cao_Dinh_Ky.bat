@echo off
title KangatangGuard v3.24.1 - Xuat Bao Cao Quan Tri Excel Da Sheet (1 Trang A4 Dashboard)
echo ======================================================================
echo    KANGATANG GUARD - TONG HOP & XUAT BAO CAO QUAN TRI EXCEL (v3.24.1)
echo ======================================================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0addin\Sync_ScanReports.ps1"
echo.
if exist "%~dp0Scan_Reports" (
    explorer.exe "%~dp0Scan_Reports"
)
pause

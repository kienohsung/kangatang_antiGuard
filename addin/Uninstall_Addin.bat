﻿@echo off
setlocal
chcp 65001 >nul
title KANGATANG GUARD - GO BO ADD-IN

:: Kiem tra quyen Administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Dang yeu cau quyen Administrator...
    powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList '/c \"\"%~f0\"\"' -Verb RunAs"
    exit /b
)

echo ======================================================================
echo   KANGATANG GUARD - GO BO ADD-IN KHOI EXCEL
echo ======================================================================
echo.

:: Dong Excel truoc
echo [1/3] Dang dong cac tien trinh Excel...
taskkill /F /IM excel.exe >nul 2>&1
timeout /t 2 /nobreak >nul
echo    -> Hoan tat.

:: Xoa file .xlam khoi XLSTART
echo [2/4] Dang xoa KangatangGuard.xlam khoi XLSTART...
set "XLAM_PATH=%APPDATA%\Microsoft\Excel\XLSTART\KangatangGuard.xlam"
if exist "%XLAM_PATH%" (
    attrib -r -h -s "%XLAM_PATH%" >nul 2>&1
    del /f /q "%XLAM_PATH%"
    echo    -> Da xoa: %XLAM_PATH%
) else (
    echo    -> File khong ton tai trong XLSTART, bo qua.
)

:: Xoa file .xlam khoi AddIns
echo [3/4] Dang xoa KangatangGuard.xlam khoi AddIns...
set "ADDINS_PATH=%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam"
if exist "%ADDINS_PATH%" (
    attrib -r -h -s "%ADDINS_PATH%" >nul 2>&1
    del /f /q "%ADDINS_PATH%"
    echo    -> Da xoa: %ADDINS_PATH%
) else (
    echo    -> File khong ton tai trong AddIns, bo qua.
)

:: Go bo Registry OPEN* dang ky
echo [4/4] Dang go bo dang ky khoi Registry Office...
powershell -NoProfile -Command "foreach ($v in @('16.0','15.0','14.0')) { $k = 'HKCU:\Software\Microsoft\Office\' + $v + '\Excel\Options'; if (Test-Path $k) { $props = (Get-ItemProperty $k -ErrorAction SilentlyContinue).psobject.Properties; foreach ($p in $props) { if ($p.Name -match '^OPEN\d*$' -and [string]$p.Value -like '*KangatangGuard.xlam*') { Remove-ItemProperty -Path $k -Name $p.Name -Force -ErrorAction SilentlyContinue; Write-Host ('   -> Da go Registry ' + $p.Name) } } } }"
echo    -> Thu muc Log duoc giu lai tai: %APPDATA%\KangatangGuard
echo.
echo ======================================================================
echo   DA GO BO THANH CONG KANGATANG GUARD!
echo   Khoi dong lai Excel de ap dung thay doi.
echo ======================================================================
echo.

echo Nhan phim bat ky de thoat...
pause >nul
exit /b

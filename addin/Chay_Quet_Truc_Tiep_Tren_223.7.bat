@echo off
title KangatangGuard v3.21.4 - Direct Server Scanner on 223.7 (Log to 223.176)
pushd "%~dp0"

echo ============================================================================
echo    KANGATANG GUARD v3.21.4 - DIRECT SERVER FOLDER SCANNER (223.7 to 223.176)
echo    Developed by OhsungIT Team
echo ============================================================================
echo.
echo [1/4] Stopping any previous stuck scanner session and preparing engine...

wmic process where "name='powershell.exe' and commandline like '%%Kangatang_FolderScanner%%'" call terminate >nul 2>&1

set "ADDIN_DIR=%APPDATA%\Microsoft\AddIns"
set "GUARD_DIR=%APPDATA%\KangatangGuard"
set "SYNC_HUB_DIR=%~dp0Sync_To_223_176"

if not exist "%ADDIN_DIR%" mkdir "%ADDIN_DIR%" >nul 2>&1
if not exist "%GUARD_DIR%" mkdir "%GUARD_DIR%" >nul 2>&1
if not exist "%SYNC_HUB_DIR%" mkdir "%SYNC_HUB_DIR%" >nul 2>&1

if exist "%~dp0KangatangGuard.xlam" (
    copy /Y "%~dp0KangatangGuard.xlam" "%ADDIN_DIR%\KangatangGuard.xlam" >nul 2>&1
)
if exist "%~dp0Kangatang_FolderScanner.ps1" (
    copy /Y "%~dp0Kangatang_FolderScanner.ps1" "%ADDIN_DIR%\Kangatang_FolderScanner.ps1" >nul 2>&1
)
if exist "%~dp0version.json" (
    copy /Y "%~dp0version.json" "%ADDIN_DIR%\version.json" >nul 2>&1
)

if exist "%~dp0Shared_Resume_State\scanned_files.log" (
    copy /Y "%~dp0Shared_Resume_State\scanned_files.log" "%GUARD_DIR%\scanned_files.log" >nul 2>&1
)
if exist "%~dp0Shared_Resume_State\last_session.json" (
    copy /Y "%~dp0Shared_Resume_State\last_session.json" "%GUARD_DIR%\last_session.json" >nul 2>&1
)

for %%D in (D E F G C) do (
    if exist "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang" (
        set "SYNC_HUB_DIR=%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176"
        if not exist "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176" mkdir "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176" >nul 2>&1
        copy /Y "%~dp0Chay_Quet_Truc_Tiep_Tren_223.7.bat" "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Chay_Quet_Truc_Tiep_Tren_223.7.bat" >nul 2>&1
        copy /Y "%~dp0Kangatang_FolderScanner.ps1" "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Kangatang_FolderScanner.ps1" >nul 2>&1
        copy /Y "%~dp0version.json" "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\version.json" >nul 2>&1
    )
)

echo [2/4] Collecting Hardware and Software Profile of 223.7 for auto-tuning...
echo [3/4] Opening Folder Selection Dialog...
echo       (Tip: Select the local drive folder on 223.7, e.g. D:\file_shared\vietnam)
echo [4/4] All scan reports, logs, and backups will be synced to 192.168.223.176!
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -Path '%ADDIN_DIR%\Kangatang_FolderScanner.ps1' -ErrorAction SilentlyContinue; Unblock-File -Path '%ADDIN_DIR%\KangatangGuard.xlam' -ErrorAction SilentlyContinue" >nul 2>&1
powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -File "%ADDIN_DIR%\Kangatang_FolderScanner.ps1" -InteractivePrompt

popd

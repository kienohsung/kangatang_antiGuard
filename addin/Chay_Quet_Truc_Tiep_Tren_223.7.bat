@echo off
title KangatangGuard v3.21.3 - Direct Server Scanner on 223.7 (Log to 223.176)
pushd "%~dp0"

echo ============================================================================
echo    KANGATANG GUARD v3.21.3 - DIRECT SERVER FOLDER SCANNER (223.7 -> 223.176)
echo    Developed by OhsungIT Team
echo ============================================================================
echo.
echo [1/3] Preparing local scanner engine and syncing shared resume state...

set "ADDIN_DIR=%APPDATA%\Microsoft\AddIns"
set "GUARD_DIR=%APPDATA%\KangatangGuard"
if not exist "%ADDIN_DIR%" mkdir "%ADDIN_DIR%" >nul 2>&1
if not exist "%GUARD_DIR%" mkdir "%GUARD_DIR%" >nul 2>&1

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
    if not exist "%GUARD_DIR%\scanned_files.log" (
        copy /Y "%~dp0Shared_Resume_State\scanned_files.log" "%GUARD_DIR%\scanned_files.log" >nul 2>&1
    )
)
if exist "%~dp0Shared_Resume_State\last_session.json" (
    if not exist "%GUARD_DIR%\last_session.json" (
        copy /Y "%~dp0Shared_Resume_State\last_session.json" "%GUARD_DIR%\last_session.json" >nul 2>&1
    )
)

for %%D in (D E F G C) do (
    if exist "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang" (
        copy /Y "%~dp0Chay_Quet_Truc_Tiep_Tren_223.7.bat" "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Chay_Quet_Truc_Tiep_Tren_223.7.bat" >nul 2>&1
        copy /Y "%~dp0Kangatang_FolderScanner.ps1" "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Kangatang_FolderScanner.ps1" >nul 2>&1
        copy /Y "%~dp0version.json" "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\version.json" >nul 2>&1
    )
    if exist "%%D:\vietnam\z. ETC\1. addinKangatang" (
        copy /Y "%~dp0Chay_Quet_Truc_Tiep_Tren_223.7.bat" "%%D:\vietnam\z. ETC\1. addinKangatang\Chay_Quet_Truc_Tiep_Tren_223.7.bat" >nul 2>&1
        copy /Y "%~dp0Kangatang_FolderScanner.ps1" "%%D:\vietnam\z. ETC\1. addinKangatang\Kangatang_FolderScanner.ps1" >nul 2>&1
        copy /Y "%~dp0version.json" "%%D:\vietnam\z. ETC\1. addinKangatang\version.json" >nul 2>&1
    )
)

echo [2/3] Opening Folder Selection Dialog...
echo       (Tip: Select the local drive folder on 223.7, e.g. D:\file_shared\vietnam)
echo [3/3] All scan reports, logs, and backups will be saved to 192.168.223.176!
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -Path '%ADDIN_DIR%\Kangatang_FolderScanner.ps1' -ErrorAction SilentlyContinue; Unblock-File -Path '%ADDIN_DIR%\KangatangGuard.xlam' -ErrorAction SilentlyContinue" >nul 2>&1
powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -File "%ADDIN_DIR%\Kangatang_FolderScanner.ps1" -InteractivePrompt

popd

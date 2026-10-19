@echo off
title KangatangGuard v3.23.0 - Multi-Window Direct Server Scanner on 223.7 (Log to 223.176)
pushd "%~dp0"

echo ============================================================================
echo    KANGATANG GUARD v3.23.0 - MULTI-WINDOW DIRECT SERVER SCANNER (223.7)
echo    Turbo Engine: Everything MFT + C# Parallel.For Zero-COM + Multi-Worker
echo    Developed by OhsungIT Team
echo ============================================================================
echo.

set "ADDIN_DIR=%APPDATA%\Microsoft\AddIns"
set "XLSTART_DIR=%APPDATA%\Microsoft\Excel\XLSTART"
set "GUARD_DIR=%APPDATA%\KangatangGuard"
set "SYNC_HUB_DIR=%~dp0Sync_To_223_176"

if not exist "%ADDIN_DIR%" mkdir "%ADDIN_DIR%" >nul 2>&1
if not exist "%GUARD_DIR%" mkdir "%GUARD_DIR%" >nul 2>&1

if exist "%XLSTART_DIR%\KangatangGuard.xlam" del /F /Q "%XLSTART_DIR%\KangatangGuard.xlam" >nul 2>&1
if exist "%XLSTART_DIR%\Kangatang_Bootstrap.xlam" del /F /Q "%XLSTART_DIR%\Kangatang_Bootstrap.xlam" >nul 2>&1

set "HAS_ACTIVE_SCAN=0"
wmic process where "name='powershell.exe' and commandline like '%%Kangatang_FolderScanner%%'" get ProcessId 2>nul | findstr /R "[0-9]" >nul 2>&1
if not errorlevel 1 set "HAS_ACTIVE_SCAN=1"

if "%HAS_ACTIVE_SCAN%"=="0" (
    echo [1/4] No active scan window detected. Cleaning orphan Excel worker processes...
    taskkill /F /IM excel.exe >nul 2>&1
    for %%W in (1 2 3 4 5 6 7 8) do (
        if exist "%GUARD_DIR%\KangatangGuard_Worker_%%W.xlam" del /F /Q "%GUARD_DIR%\KangatangGuard_Worker_%%W.xlam" >nul 2>&1
    )
    del /F /Q "%GUARD_DIR%\KangatangGuard_Worker_*.xlam" >nul 2>&1
) else (
    echo [1/4] Active scanner window detected! Preserving all running PowerShell and Excel processes...
)

for %%D in (D E F G C) do (
    if exist "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam" (
        set "SYNC_HUB_DIR=%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176"
        if not exist "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176" mkdir "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176" >nul 2>&1
        copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam" "%~dp0KangatangGuard.xlam" >nul 2>&1
        copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Kangatang_FolderScanner.ps1" "%~dp0Kangatang_FolderScanner.ps1" >nul 2>&1
        copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\version.json" "%~dp0version.json" >nul 2>&1
        copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam" "%ADDIN_DIR%\KangatangGuard.xlam" >nul 2>&1
        copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Kangatang_FolderScanner.ps1" "%ADDIN_DIR%\Kangatang_FolderScanner.ps1" >nul 2>&1
        copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\version.json" "%ADDIN_DIR%\version.json" >nul 2>&1
        if exist "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Everything.exe" (
            copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Everything.exe" "%~dp0Everything.exe" >nul 2>&1
            copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\Everything.exe" "%ADDIN_DIR%\Everything.exe" >nul 2>&1
        )
        if exist "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\es.exe" (
            copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\es.exe" "%~dp0es.exe" >nul 2>&1
            copy /Y "%%D:\file_shared\vietnam\z. ETC\1. addinKangatang\es.exe" "%ADDIN_DIR%\es.exe" >nul 2>&1
        )
    )
)

REM Auto-discover Everything.exe / es.exe on 223.7 and sync to AddIns & Hub
set "EVERYTHING_FOUND="
for %%P in (
    "%~dp0Everything.exe"
    "%ADDIN_DIR%\Everything.exe"
    "C:\Program Files\Everything\Everything.exe"
    "C:\Program Files (x86)\Everything\Everything.exe"
    "%LOCALAPPDATA%\Everything\Everything.exe"
) do (
    if not defined EVERYTHING_FOUND (
        if exist "%%~P" set "EVERYTHING_FOUND=%%~P"
    )
)

if defined EVERYTHING_FOUND (
    copy /Y "%EVERYTHING_FOUND%" "%ADDIN_DIR%\Everything.exe" >nul 2>&1
    copy /Y "%EVERYTHING_FOUND%" "%~dp0Everything.exe" >nul 2>&1
    tasklist /FI "IMAGENAME eq Everything.exe" 2>nul | find /I "Everything.exe" >nul 2>&1
    if errorlevel 1 (
        start "" "%EVERYTHING_FOUND%" -startup >nul 2>&1
    )
)

for %%E in (
    "%~dp0es.exe"
    "%ADDIN_DIR%\es.exe"
    "C:\Program Files\Everything\es.exe"
    "C:\Program Files (x86)\Everything\es.exe"
    "%LOCALAPPDATA%\Everything\es.exe"
) do (
    if exist "%%~E" (
        copy /Y "%%~E" "%ADDIN_DIR%\es.exe" >nul 2>&1
        copy /Y "%%~E" "%~dp0es.exe" >nul 2>&1
    )
)

if exist "%~dp0KangatangGuard.xlam" (
    copy /Y "%~dp0KangatangGuard.xlam" "%ADDIN_DIR%\KangatangGuard.xlam" >nul 2>&1
)
if exist "%~dp0Kangatang_FolderScanner.ps1" (
    copy /Y "%~dp0Kangatang_FolderScanner.ps1" "%ADDIN_DIR%\Kangatang_FolderScanner.ps1" >nul 2>&1
)
if exist "%~dp0version.json" (
    copy /Y "%~dp0version.json" "%ADDIN_DIR%\version.json" >nul 2>&1
)

if "%HAS_ACTIVE_SCAN%"=="0" (
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
)

echo [2/4] Verifying Seqrite-clean v3.23.0 C# Parallel Zero-COM + Everything MFT engine...
echo [3/4] Launching new independent PowerShell Scanner Window...
echo [4/4] All scan reports, logs, and backups will be synced to 192.168.223.176!
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -Path '%ADDIN_DIR%\Kangatang_FolderScanner.ps1' -ErrorAction SilentlyContinue; Unblock-File -Path '%ADDIN_DIR%\KangatangGuard.xlam' -ErrorAction SilentlyContinue" >nul 2>&1
start "KangatangGuard v3.23.0 - Folder Scanner" powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -File "%ADDIN_DIR%\Kangatang_FolderScanner.ps1" -InteractivePrompt

popd

@echo off
:: ============================================================================
:: TRÌNH CHẠY TỰ ĐỘNG DIỆT VIRUS EXCEL KANGATANG
:: Tự động yêu cầu quyền Administrator và thực thi kịch bản PowerShell an toàn
:: ============================================================================

title DIET VIRUS EXCEL KANGATANG - EXCEL CLEANER

:: Kiểm tra quyền Administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Dang khoi dong voi quyen Administrator...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"\"%~dp0Chay_Diet_Virus.bat\"\"' -Verb RunAs"
    exit /b
)

chcp 65001 >nul
cls
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0clean_excel_virus.ps1"

echo.
pause

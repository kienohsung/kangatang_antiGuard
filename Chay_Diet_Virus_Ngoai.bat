@echo off
:: ============================================================================
:: KANGATANGGUARD v3.8.1 - TRÌNH DIỆT VIRUS EXCEL ĐỘC LẬP (STANDALONE 1-CLICK)
:: Dành cho mọi máy tính trong mạng LAN (Không cần cài đặt Add-in)
:: ============================================================================

title KANGATANGGUARD v3.8.1 - DIET VIRUS EXCEL NGOAI (STANDALONE)
chcp 65001 >nul

:: Chuyển đổi ngữ cảnh đường dẫn an toàn (Hỗ trợ đường dẫn mạng UNC)
pushd "%~dp0"

echo ======================================================================
echo    KANGATANGGUARD v3.8.1 - TRINH DIET VIRUS EXCEL NGOAI (STANDALONE)
echo    Khoi chay bo may diet virus doc lap chong treo...
echo ======================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Kangatang_Standalone_Scanner.ps1"

popd
echo.
echo ======================================================================
echo Da ket thuc chuong trinh. Nhan phim bat ky de thoat...
pause >nul

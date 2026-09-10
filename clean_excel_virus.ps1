# ==============================================================================
# clean_excel_virus.ps1 -> Chuyển tiếp tương thích ngược sang Kangatang_Standalone_Scanner.ps1
# Phiên bản: v3.8.1
# ==============================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$standaloneScanner = Join-Path $ScriptDir "Kangatang_Standalone_Scanner.ps1"

if (Test-Path $standaloneScanner) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $standaloneScanner @args
} else {
    Write-Host "[LỖI] Không tìm thấy Kangatang_Standalone_Scanner.ps1 tại $ScriptDir" -ForegroundColor Red
}

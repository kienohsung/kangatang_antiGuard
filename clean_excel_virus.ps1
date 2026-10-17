# ==============================================================================
# clean_excel_virus.ps1 -> Chuyen tiep tuong thich nguoc sang Kangatang_Standalone_Scanner.ps1
# Phien ban: v3.19.0
# ==============================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$candidates = @(
    (Join-Path $ScriptDir "Kangatang_Standalone_Scanner.ps1"),
    (Join-Path $ScriptDir "addin\Kangatang_Standalone_Scanner.ps1")
)
$standaloneScanner = $null
foreach ($c in $candidates) {
    if (Test-Path $c) { $standaloneScanner = $c; break }
}

if ($standaloneScanner) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $standaloneScanner @args
} else {
    Write-Host "[LOI] Khong tim thay Kangatang_Standalone_Scanner.ps1 tai $ScriptDir" -ForegroundColor Red
}

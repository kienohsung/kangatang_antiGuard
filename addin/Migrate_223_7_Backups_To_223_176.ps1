# ==============================================================================
# Migrate_223_7_Backups_To_223_176.ps1
# Script di chuyen toan bo 8,470+ file backup cach ly tu 223.7 ve may ca nhan 223.176
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$SourceDir = "\\192.168.223.7\file_shared\vietnam\z. ETC\2. Virus backupfile - DO NOT OPEN IT"
$TargetDir = "D:\7. AI tools\kangatang\Quarantine_Backup"
$LogFile   = "D:\7. AI tools\kangatang\Quarantine_Backup_Migration.log"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - DONG BO FILE CACH LY VE MAY 223.176               " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "Nguon (223.7)  : $SourceDir"
Write-Host "Dich (223.176) : $TargetDir"
Write-Host "Nhat ky (Log)  : $LogFile"

if (-not (Test-Path $SourceDir)) {
    Write-Host "[LOI] Khong the ket noi den may chu 223.7: $SourceDir" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
}

Write-Host "`nDang chay Robocopy da luong (8 threads) de sao chep toan bo file..." -ForegroundColor Yellow
$sw = [System.Diagnostics.Stopwatch]::StartNew()

robocopy "$SourceDir" "$TargetDir" /E /MT:8 /R:1 /W:1 /NP /LOG:"$LogFile"
$rc = $LASTEXITCODE

$sw.Stop()
Write-Host "`nThoi gian thuc hien: $([math]::Round($sw.Elapsed.TotalMinutes, 2)) phut."

if ($rc -le 3) {
    Write-Host "[OK] Da dong bo thanh cong toan bo file ve $TargetDir (Robocopy ExitCode: $rc)" -ForegroundColor Green
    $count = (Get-ChildItem -Path $TargetDir -File).Count
    Write-Host "Tong so file hien co tai kho 223.176: $count file" -ForegroundColor Cyan
} else {
    Write-Host "[CANH BAO] Robocopy ket thuc voi ma loi: $rc. Vui long xem log tai $LogFile" -ForegroundColor Yellow
}

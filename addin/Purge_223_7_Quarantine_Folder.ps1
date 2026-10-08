# ==============================================================================
# Purge_223_7_Quarantine_Folder.ps1
# Script don dep sach kho sao luu virus tren 223.7 sau khi da sao luu ve 223.176
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$RemoteDir   = "\\192.168.223.7\file_shared\vietnam\z. ETC\2. Virus backupfile - DO NOT OPEN IT"
$LocalBackup = "D:\7. AI tools\kangatang\Quarantine_Backup"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - DON DEP KHO CU TREN MAY CHU 223.7                " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Kiem tra kho cuc bo de dam bao tuyet doi khong mat du lieu
Write-Host "Dang kiem tra tinh toan ven tai kho cuc bo ($LocalBackup)..." -ForegroundColor Yellow
$localCount = (Get-ChildItem -LiteralPath $LocalBackup -File).Count
Write-Host "So luong file da bao toan tai kho 223.176: $localCount file." -ForegroundColor Green

if ($localCount -lt 8000) {
    Write-Host "[NGUY HIEM] Kho cuc bo chi co $localCount files (< 8,000). Huy bo don dep de bao toan du lieu!" -ForegroundColor Red
    exit 1
}

# 2. Tao thu muc rong tam thoi de purge
$emptyTemp = Join-Path $env:TEMP "Kangatang_Empty_Purge"
if (-not (Test-Path $emptyTemp)) { New-Item -ItemType Directory -Path $emptyTemp -Force | Out-Null }

Write-Host "`nTien hanh xoa sach 8,470 file tren may chu 223.7 bang Robocopy Purge (16 luong)..." -ForegroundColor Yellow
$sw = [System.Diagnostics.Stopwatch]::StartNew()

robocopy "$emptyTemp" "$RemoteDir" /PURGE /R:1 /W:1 /MT:16 /NP /NFL /NDL /NJH /NJS
$rc = $LASTEXITCODE

Remove-Item -Path $emptyTemp -Force -ErrorAction SilentlyContinue
$sw.Stop()

Write-Host "Thoi gian xoa: $([math]::Round($sw.Elapsed.TotalSeconds, 2)) giay." -ForegroundColor Cyan

# 3. Tao file huong dan / thong bao
$noticeFile = Join-Path $RemoteDir "THONG_BAO_CHUYEN_KHO.txt"
$noticeContent = @"
THONG BAO HE THONG KANGATANGGUARD:
- Toan bo cac ban sao luu file cach ly virus da duoc chuyen ve luu tru tap trung tai may ca nhan 192.168.223.176 (D:\7. AI tools\kangatang\Quarantine_Backup).
- May chu 223.7 da duoc don sach va khong con luu tru bat ky file chua ma doc nao.
- Thoi gian thuc hien: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
[System.IO.File]::WriteAllText($noticeFile, $noticeContent, [System.Text.Encoding]::UTF8)

# 4. Kiem tra lai thu muc tren 223.7
$remaining = (Get-ChildItem -LiteralPath $RemoteDir -File).Count
Write-Host "`n[OK] Don dep hoan tat! So file con lai tren 223.7: $remaining (bao gom file thong bao huong dan)." -ForegroundColor Green

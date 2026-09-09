# ==============================================================================
# Setup_Host_LAN_Share.ps1
# Thiet lap Thu muc Chia se (SMB Share) KangatangGuard_Hub tren May chu
# Phien ban: v3.7.0 (Production-grade LAN Distribution Hub)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - THIET LAP TRUNG TAM PHAN PHOI LAN HUB v3.7.0     " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$RootDir   = Split-Path -Parent $ScriptDir
$DistDir   = Join-Path $RootDir "distribution"

if (-not (Test-Path $DistDir)) {
    New-Item -ItemType Directory -Path $DistDir -Force | Out-Null
    Write-Host "[OK] Da tao thu muc distribution: $DistDir" -ForegroundColor Green
}

$ShareName = "KangatangGuard_Hub"
$ExistingShare = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue

if ($null -eq $ExistingShare) {
    Write-Host "`n[1/2] Dang tao SMB Share '$ShareName'..." -ForegroundColor Yellow
    try {
        New-SmbShare -Name $ShareName -Path $DistDir -ReadAccess "Everyone" -Description "KangatangGuard LAN Distribution Hub" -ErrorAction Stop | Out-Null
        Write-Host "   -> [OK] Da tao thanh cong SMB Share '$ShareName'" -ForegroundColor Green
    } catch {
        Write-Host "   -> [CANH BAO] Can quyen Administrator de tao SMB Share." -ForegroundColor Yellow
        Write-Host "   -> Dang thu khoi chay voi quyen Administrator..." -ForegroundColor Yellow
        $arg = "-NoProfile -ExecutionPolicy Bypass -Command `"New-SmbShare -Name '$ShareName' -Path '$DistDir' -ReadAccess 'Everyone' -Description 'KangatangGuard LAN Distribution Hub'`""
        Start-Process powershell -Verb RunAs -ArgumentList $arg -Wait
    }
} else {
    Write-Host "`n[1/2] SMB Share '$ShareName' da ton tai tren he thong:" -ForegroundColor Green
    Write-Host "   -> Duong dan: $($ExistingShare.Path)" -ForegroundColor Gray
}

# Cau hinh NTFS Permissions: Everyone Read & Execute
Write-Host "`n[2/2] Dang phan quyen NTFS Read-Only cho Everyone..." -ForegroundColor Yellow
try {
    icacls $DistDir /grant "Everyone:(OI)(CI)RX" /T /C | Out-Null
    Write-Host "   -> [OK] Da phan quyen Read & Execute thanh cong cho Everyone." -ForegroundColor Green
} catch {
    Write-Host "   -> [CANH BAO] Khong the phan quyen icacls: $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# Lay danh sach dia chi IP LAN cua May chu
$ipList = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { 
    $_.InterfaceAlias -notmatch 'Loopback|vEthernet' -and $_.IPAddress -notlike '169.254*' 
}).IPAddress

$primaryIP = if ($ipList) { $ipList[0] } else { "127.0.0.1" }

Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   THIET LAP TRUNG TAM PHAN PHOI LAN THANH CONG!                     " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   Ten may chu (ComputerName) : $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "   Dia chi IP May chu         : $primaryIP" -ForegroundColor Cyan
Write-Host ""
Write-Host "   Duong dan mang (UNC Paths) de cac may khac truy cap:" -ForegroundColor Yellow
Write-Host "   1. \\$env:COMPUTERNAME\$ShareName" -ForegroundColor White
Write-Host "   2. \\$primaryIP\$ShareName" -ForegroundColor White
Write-Host ""
Write-Host "   Huong dan cho cac may Client:" -ForegroundColor Yellow
Write-Host "   - Tren may khac trong LAN, mo Run (Win + R) va go: \\$env:COMPUTERNAME\$ShareName" -ForegroundColor White
Write-Host "   - Chay file 'Install_Client_Kangatang.bat' de cai dat 1-Click!" -ForegroundColor White
Write-Host ""

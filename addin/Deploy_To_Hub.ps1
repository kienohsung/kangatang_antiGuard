# ==============================================================================
# Deploy_To_Hub.ps1
# Lenh Phat hanh va Dong bo Phien ban moi len Trung tam LAN Hub
# Phien ban: v3.7.0 (Production-grade Release Orchestrator)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$RootDir   = if (Test-Path (Join-Path $ScriptDir "KangatangGuard_Code.vba")) { Split-Path -Parent $ScriptDir } else { $ScriptDir }
$AddinDir  = Join-Path $RootDir "addin"
$DistDir   = Join-Path $RootDir "distribution"

$ReleaseVersion = "3.7.0"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - PHAT HANH PHIEN BAN MOI LEN LAN HUB (v$ReleaseVersion)  " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Bien dich Add-in .xlam tren May chu
Write-Host "`n[1/4] Dang bien dich KangatangGuard.xlam moi nhat tu ma nguon VBA..." -ForegroundColor Yellow
$installerScript = Join-Path $AddinDir "Install_KangatangGuard.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $installerScript

$compiledXlam = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART\KangatangGuard.xlam"
if (-not (Test-Path $compiledXlam)) {
    Write-Host "`n[LOI] Khong tim thay KangatangGuard.xlam da bien dich tai $compiledXlam!" -ForegroundColor Red
    exit 1
}
Write-Host "   -> [OK] Bien dich thanh cong KangatangGuard.xlam" -ForegroundColor Green

# 2. Dong bo file sang thu muc distribution
Write-Host "`n[2/4] Dang dong bo tap tin sang thu muc phan phoi (distribution)..." -ForegroundColor Yellow
if (-not (Test-Path $DistDir)) {
    New-Item -ItemType Directory -Path $DistDir -Force | Out-Null
}

$targetXlam = Join-Path $DistDir "KangatangGuard.xlam"
if (Test-Path $targetXlam) {
    try { Set-ItemProperty -Path $targetXlam -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
}
Copy-Item -Path $compiledXlam -Destination $targetXlam -Force
try { Set-ItemProperty -Path $targetXlam -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}
Write-Host "   -> [OK] Da sao chep KangatangGuard.xlam vao distribution" -ForegroundColor Green

$srcScanner = Join-Path $AddinDir "Kangatang_FolderScanner.ps1"
if (Test-Path $srcScanner) {
    $targetScanner = Join-Path $DistDir "Kangatang_FolderScanner.ps1"
    Copy-Item -Path $srcScanner -Destination $targetScanner -Force
    Write-Host "   -> [OK] Da sao chep Kangatang_FolderScanner.ps1 vao distribution" -ForegroundColor Green
}

# 3. Tao file version.json
Write-Host "`n[3/4] Dang tao tệp sieu du lieu version.json..." -ForegroundColor Yellow
$ipList = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { 
    $_.InterfaceAlias -notmatch 'Loopback|vEthernet' -and $_.IPAddress -notlike '169.254*' 
}).IPAddress
$primaryIP = if ($ipList) { $ipList[0] } else { "192.168.223.176" }

$versionObj = @{
    version           = $ReleaseVersion
    release_date      = (Get-Date -Format "yyyy-MM-dd")
    release_timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    hub_host          = $env:COMPUTERNAME
    hub_ip            = $primaryIP
    share_name        = "KangatangGuard_Hub"
    unc_path          = "\\$env:COMPUTERNAME\KangatangGuard_Hub"
    files             = @{
        xlam    = "KangatangGuard.xlam"
        scanner = "Kangatang_FolderScanner.ps1"
    }
    changelog         = "v${ReleaseVersion} - Trung tam Phan phoi Mang LAN Hub va Tu dong Cap nhat Dong bo (Offline-First Architecture)"
}

$versionJsonPath = Join-Path $DistDir "version.json"
$versionJsonContent = $versionObj | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText($versionJsonPath, $versionJsonContent, [System.Text.Encoding]::UTF8)
Write-Host "   -> [OK] Da tao version.json: $versionJsonPath" -ForegroundColor Green

# 4. Kiem tra tinh san sang cua SMB Share tren mang
Write-Host "`n[4/4] Kiem tra ket noi mang LAN den Hub..." -ForegroundColor Yellow
$uncTestHost = "\\$env:COMPUTERNAME\KangatangGuard_Hub\version.json"
$uncTestIP   = "\\$primaryIP\KangatangGuard_Hub\version.json"

$hostOk = Test-Path $uncTestHost
$ipOk   = Test-Path $uncTestIP

if ($hostOk -or $ipOk) {
    Write-Host "   -> [OK] Mang LAN da san sang:" -ForegroundColor Green
    if ($hostOk) { Write-Host "      + UNC Host: $uncTestHost" -ForegroundColor Green }
    if ($ipOk)   { Write-Host "      + UNC IP  : $uncTestIP" -ForegroundColor Green }
} else {
    Write-Host "   -> [CANH BAO] Chua phat hien SMB Share hoat dong." -ForegroundColor DarkYellow
    Write-Host "   -> Dang tu dong kich hoat Setup_Host_LAN_Share.ps1..." -ForegroundColor Yellow
    $setupShare = Join-Path $AddinDir "Setup_Host_LAN_Share.ps1"
    if (Test-Path $setupShare) {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $setupShare
    }
}

Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   PHAT HANH PHIEN BAN v$ReleaseVersion LEN LAN HUB THANH CONG!        " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   Tat ca cac may client trong mang LAN khi mo Excel se tu dong" -ForegroundColor Cyan
Write-Host "   nhan dien va cap nhat len phien ban v$ReleaseVersion!" -ForegroundColor Cyan
Write-Host ""

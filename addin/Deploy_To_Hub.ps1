# ==============================================================================
# Deploy_To_Hub.ps1
# Lenh Phat hanh va Dong bo Phien ban moi len May chu LAN & Tao Clone Cuc bo
# Phien ban: v3.8.0 (Production-grade Dual-Mirror Release Orchestrator)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$RootDir   = if (Test-Path (Join-Path $ScriptDir "KangatangGuard_Code.vba")) { Split-Path -Parent $ScriptDir } else { $ScriptDir }
$AddinDir  = Join-Path $RootDir "addin"
$LocalDist = Join-Path $RootDir "distribution"
$DesktopDir = "C:\Users\mrKienIT\Desktop\python\coding\AI tools\kangatang"

# Dia chi May chu Tep LAN chuyen dung (Online 24/7)
$PrimaryRemoteHub = "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang"

$ReleaseVersion = "3.8.0"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - PHAT HANH PHIEN BAN MOI (v$ReleaseVersion - DUAL-MIRROR) " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Bien dich Add-in .xlam tren May chu
Write-Host "`n[1/5] Dang bien dich KangatangGuard.xlam v$ReleaseVersion tu ma nguon VBA..." -ForegroundColor Yellow
$installerScript = Join-Path $AddinDir "Install_KangatangGuard.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $installerScript

$compiledXlam = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART\KangatangGuard.xlam"
if (-not (Test-Path $compiledXlam)) {
    Write-Host "`n[LOI] Khong tim thay KangatangGuard.xlam da bien dich tai $compiledXlam!" -ForegroundColor Red
    exit 1
}
Write-Host "   -> [OK] Bien dich thanh cong KangatangGuard.xlam" -ForegroundColor Green

# 2. Tao file sieu du lieu version.json
Write-Host "`n[2/5] Dang tao tệp sieu du lieu version.json..." -ForegroundColor Yellow
$versionObj = @{
    version           = $ReleaseVersion
    release_date      = (Get-Date -Format "yyyy-MM-dd")
    release_timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    hub_unc           = $PrimaryRemoteHub
    hub_ip            = "192.168.223.7"
    local_clone_host  = $env:COMPUTERNAME
    local_clone_ip    = "192.168.223.176"
    files             = @{
        xlam    = "KangatangGuard.xlam"
        scanner = "Kangatang_FolderScanner.ps1"
    }
    changelog         = "v${ReleaseVersion} - Chuyen sang May chu Tep Chuyen dung LAN (\\192.168.223.7) & Co che Clone Kep tren 223.176"
}

$versionJsonContent = $versionObj | ConvertTo-Json -Depth 4
$localJsonPath = Join-Path $LocalDist "version.json"
[System.IO.File]::WriteAllText($localJsonPath, $versionJsonContent, [System.Text.Encoding]::UTF8)
Write-Host "   -> [OK] Da tao version.json: $localJsonPath" -ForegroundColor Green

# 3. Dong bo tap tin vao Ban Clone Cuc bo (Local Clone tren may 223.176)
Write-Host "`n[3/5] Dang tao ban CLONE cuc bo tai may 223.176 ($LocalDist)..." -ForegroundColor Yellow
if (-not (Test-Path $LocalDist)) {
    New-Item -ItemType Directory -Path $LocalDist -Force | Out-Null
}

$targetLocalXlam = Join-Path $LocalDist "KangatangGuard.xlam"
if (Test-Path $targetLocalXlam) {
    try { Set-ItemProperty -Path $targetLocalXlam -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
}
Copy-Item -Path $compiledXlam -Destination $targetLocalXlam -Force
try { Set-ItemProperty -Path $targetLocalXlam -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}

$srcScanner = Join-Path $AddinDir "Kangatang_FolderScanner.ps1"
if (Test-Path $srcScanner) {
    Copy-Item -Path $srcScanner -Destination (Join-Path $LocalDist "Kangatang_FolderScanner.ps1") -Force
}
Write-Host "   -> [OK] Ban CLONE cuc bo tren may 223.176 da duoc dong bo day du." -ForegroundColor Green

# 4. Phat hanh len May chu Tep LAN Chuyen dung (\\192.168.223.7)
Write-Host "`n[4/5] Dang phat hanh len May chu Tep LAN ($PrimaryRemoteHub)..." -ForegroundColor Yellow
if (-not (Test-Path $PrimaryRemoteHub)) {
    Write-Host "   -> [CANH BAO] Khong the ket noi den $PrimaryRemoteHub!" -ForegroundColor Red
    Write-Host "   -> Vui long kiem tra may 192.168.223.7 co dang ket noi mang khong." -ForegroundColor Red
} else {
    $remoteXlam = Join-Path $PrimaryRemoteHub "KangatangGuard.xlam"
    if (Test-Path $remoteXlam) {
        try { Set-ItemProperty -Path $remoteXlam -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    }
    Copy-Item -Path $compiledXlam -Destination $remoteXlam -Force
    try { Set-ItemProperty -Path $remoteXlam -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}
    Write-Host "   -> [OK] Da chep KangatangGuard.xlam len May chu 223.7" -ForegroundColor Green

    # Chep cac tep ho tro
    Copy-Item -Path (Join-Path $LocalDist "Kangatang_FolderScanner.ps1") -Destination (Join-Path $PrimaryRemoteHub "Kangatang_FolderScanner.ps1") -Force
    Copy-Item -Path $localJsonPath -Destination (Join-Path $PrimaryRemoteHub "version.json") -Force
    Copy-Item -Path (Join-Path $LocalDist "Install_Client_Kangatang.bat") -Destination (Join-Path $PrimaryRemoteHub "Install_Client_Kangatang.bat") -Force
    Copy-Item -Path (Join-Path $LocalDist "Install_Client.ps1") -Destination (Join-Path $PrimaryRemoteHub "Install_Client.ps1") -Force
    Copy-Item -Path (Join-Path $LocalDist "README_HUONG_DAN_CLIENT.txt") -Destination (Join-Path $PrimaryRemoteHub "README_HUONG_DAN_CLIENT.txt") -Force

    # Dong bo file zip phat hanh
    $zipFileVer = Join-Path $AddinDir "addin_kangatang_v${ReleaseVersion}.zip"
    if (Test-Path $zipFileVer) {
        Copy-Item -Path $zipFileVer -Destination (Join-Path $PrimaryRemoteHub "addin_kangatang_v${ReleaseVersion}.zip") -Force
        Copy-Item -Path $zipFileVer -Destination (Join-Path $PrimaryRemoteHub "addin_kangatang.zip") -Force
    }
    Write-Host "   -> [OK] Da dong bo toan bo goi cai dat & file zip len May chu 223.7 thanh cong!" -ForegroundColor Green
}

# 5. Dong bo sang Desktop Mirror
Write-Host "`n[5/5] Dang dong bo sang thu muc Desktop Mirror..." -ForegroundColor Yellow
if (Test-Path $DesktopDir) {
    Copy-Item (Join-Path $LocalDist "*") -Destination (Join-Path $DesktopDir "distribution") -Recurse -Force
    Copy-Item (Join-Path $AddinDir "*") -Destination (Join-Path $DesktopDir "addin") -Recurse -Force -Exclude "*.zip"
    Write-Host "   -> [OK] Desktop mirror da duoc cap nhat." -ForegroundColor Green
} else {
    Write-Host "   -> [Bo qua] Khong tim thay thu muc Desktop mirror." -ForegroundColor Gray
}

Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   PHAT HANH PHIEN BAN v$ReleaseVersion THANH CONG! (DUAL-MIRROR)       " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   1. May chu phan phoi LAN chính : $PrimaryRemoteHub" -ForegroundColor Cyan
Write-Host "   2. Ban Clone cuc bo may 223.176 : $LocalDist" -ForegroundColor Cyan
Write-Host ""
Write-Host "   Tat ca cac may client trong mang LAN khi mo Excel se tu dong" -ForegroundColor White
Write-Host "   nhan dien va cap nhat len phien ban v$ReleaseVersion tu server 223.7!" -ForegroundColor White
Write-Host ""

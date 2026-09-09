# ==============================================================================
# Install_Client.ps1
# PowerShell Installer cho May Client trong Mang LAN
# Phien ban: v3.7.0 (Production-grade LAN Client Auto-Sync)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - CAI DAT ADD-IN TU MAY CHU LAN (CLIENT v3.7.0)    " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

# Xac dinh Nguon Cap nhat (UpdateSource UNC Path)
$DefaultServerName = "CM-GA-MRKIENIT1"
$DefaultServerIP   = "192.168.223.176"
$ShareName         = "KangatangGuard_Hub"

$uncSource = ""
if ($ScriptDir.StartsWith("\\")) {
    $uncSource = $ScriptDir
} else {
    $uncHost = "\\$DefaultServerName\$ShareName"
    $uncIP   = "\\$DefaultServerIP\$ShareName"
    if (Test-Path $uncHost) {
        $uncSource = $uncHost
    } elseif (Test-Path $uncIP) {
        $uncSource = $uncIP
    } else {
        $uncSource = $ScriptDir
    }
}

Write-Host "`n[1/5] May chu phan phoi (Hub): $uncSource" -ForegroundColor Yellow

# Doc thong tin phien ban tu version.json neu co
$versionInfo = "3.7.0"
$versionJsonPath = Join-Path $uncSource "version.json"
if (Test-Path $versionJsonPath) {
    try {
        $jsonContent = Get-Content -Path $versionJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($jsonContent.version) {
            $versionInfo = $jsonContent.version
        }
    } catch {}
}
Write-Host "   -> Phien ban phat hanh: v$versionInfo" -ForegroundColor Green

# Kiem tra file nguon
$srcXlam    = Join-Path $uncSource "KangatangGuard.xlam"
$srcScanner = Join-Path $uncSource "Kangatang_FolderScanner.ps1"

if (-not (Test-Path $srcXlam)) {
    # Fallback to current script directory if different
    $srcXlam = Join-Path $ScriptDir "KangatangGuard.xlam"
}
if (-not (Test-Path $srcScanner)) {
    $srcScanner = Join-Path $ScriptDir "Kangatang_FolderScanner.ps1"
}

if (-not (Test-Path $srcXlam)) {
    Write-Host "`n[LOI] Khong tim thay file KangatangGuard.xlam tai $srcXlam!" -ForegroundColor Red
    Write-Host "Vui long kiem tra ket noi den May chu LAN: $uncSource" -ForegroundColor Red
    exit 1
}

# Dong cac tien trinh Excel dang chay
Write-Host "`n[2/5] Dang dong tien trinh Excel dang chay..." -ForegroundColor Yellow
$excelProcs = Get-Process -Name excel -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host "   -> Dang dong $($excelProcs.Count) cua so Excel de cap nhat tep an toan..." -ForegroundColor Gray
    $excelProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
} else {
    Write-Host "   -> Khong co Excel nao dang chay." -ForegroundColor Gray
}

# Cau hinh Office Trust Center
Write-Host "`n[3/5] Dang cau hinh Office Trust Center (Cho phep Macro & Mang noi bo)..." -ForegroundColor Yellow
$officeVersions = @("14.0", "15.0", "16.0")

foreach ($ver in $officeVersions) {
    $secKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $secKey) {
        try {
            # Cho phep macro tu vi tri mang (Trusted Locations tren mang)
            Set-ItemProperty -Path $secKey -Name "AllowNetworkLocations" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $secKey -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        } catch {}

        # Them Trusted Locations
        $trustedRoot = "$secKey\Trusted Locations"
        if (-not (Test-Path $trustedRoot)) {
            New-Item -Path $trustedRoot -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # Trust LAN Hub
        $locHub = "$trustedRoot\KangatangLANHub"
        if (-not (Test-Path $locHub)) { New-Item -Path $locHub -Force -ErrorAction SilentlyContinue | Out-Null }
        Set-ItemProperty -Path $locHub -Name "Path" -Value $uncSource -Type String -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locHub -Name "AllowSubfolders" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locHub -Name "Description" -Value "KangatangGuard LAN Hub" -Type String -Force -ErrorAction SilentlyContinue

        # Trust local APPDATA KangatangGuard
        $locLocal = "$trustedRoot\KangatangLocal"
        if (-not (Test-Path $locLocal)) { New-Item -Path $locLocal -Force -ErrorAction SilentlyContinue | Out-Null }
        $appDataGuard = Join-Path $env:APPDATA "KangatangGuard"
        Set-ItemProperty -Path $locLocal -Name "Path" -Value $appDataGuard -Type String -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locLocal -Name "AllowSubfolders" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locLocal -Name "Description" -Value "KangatangGuard Local Cache" -Type String -Force -ErrorAction SilentlyContinue
    }
}
Write-Host "   -> [OK] Da cau hinh Trusted Locations thanh cong." -ForegroundColor Green

# Sao chep tep vao bo nho dem cuc bo (Local Cache: XLSTART va APPDATA)
Write-Host "`n[4/5] Dang cai dat Add-in vao bo nho dem cuc bo (Offline-First Cache)..." -ForegroundColor Yellow

$xlStartPath = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (-not (Test-Path $xlStartPath)) { New-Item -ItemType Directory -Path $xlStartPath -Force | Out-Null }

$addInsPath = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path $addInsPath)) { New-Item -ItemType Directory -Path $addInsPath -Force | Out-Null }

$localAppData = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $localAppData)) { New-Item -ItemType Directory -Path $localAppData -Force | Out-Null }

$sessionsDir = Join-Path $localAppData "Sessions"
if (-not (Test-Path $sessionsDir)) { New-Item -ItemType Directory -Path $sessionsDir -Force | Out-Null }

$updatesDir = Join-Path $localAppData "staged_update"
if (-not (Test-Path $updatesDir)) { New-Item -ItemType Directory -Path $updatesDir -Force | Out-Null }

# Sao chep KangatangGuard.xlam vao XLSTART
$dstXlamXLSTART = Join-Path $xlStartPath "KangatangGuard.xlam"
if (Test-Path $dstXlamXLSTART) {
    try { Set-ItemProperty -Path $dstXlamXLSTART -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    Remove-Item -Path $dstXlamXLSTART -Force -ErrorAction SilentlyContinue
}
Copy-Item -Path $srcXlam -Destination $dstXlamXLSTART -Force
try { Set-ItemProperty -Path $dstXlamXLSTART -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}
Write-Host "   -> [OK] Da cai dat vao XLSTART: $dstXlamXLSTART" -ForegroundColor Green

# Sao chep sang AddIns (Dual-registration)
$dstXlamAddIns = Join-Path $addInsPath "KangatangGuard.xlam"
if (Test-Path $dstXlamAddIns) {
    try { Set-ItemProperty -Path $dstXlamAddIns -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    Remove-Item -Path $dstXlamAddIns -Force -ErrorAction SilentlyContinue
}
Copy-Item -Path $srcXlam -Destination $dstXlamAddIns -Force
try { Set-ItemProperty -Path $dstXlamAddIns -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}
Write-Host "   -> [OK] Da cai dat vao AddIns: $dstXlamAddIns" -ForegroundColor Green

# Sao chep Background Worker
if (Test-Path $srcScanner) {
    $dstScanner = Join-Path $localAppData "Kangatang_FolderScanner.ps1"
    Copy-Item -Path $srcScanner -Destination $dstScanner -Force
    Write-Host "   -> [OK] Da cai dat Background Worker: $dstScanner" -ForegroundColor Green
}

# Dang ky vao Registry KangatangGuard
Write-Host "`n[5/5] Dang dang ky Nguon Cap nhat LAN vao Registry..." -ForegroundColor Yellow
$regKey = "HKCU:\Software\KangatangGuard"
if (-not (Test-Path $regKey)) { New-Item -Path $regKey -Force | Out-Null }

Set-ItemProperty -Path $regKey -Name "UpdateSource" -Value $uncSource -Force
Set-ItemProperty -Path $regKey -Name "InstalledVersion" -Value $versionInfo -Force
Set-ItemProperty -Path $regKey -Name "ScannerScript" -Value (Join-Path $localAppData "Kangatang_FolderScanner.ps1") -Force
Set-ItemProperty -Path $regKey -Name "LastUpdateCheck" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Force

Write-Host "   -> [OK] UpdateSource    : $uncSource" -ForegroundColor Green
Write-Host "   -> [OK] InstalledVersion: v$versionInfo" -ForegroundColor Green

# Dual Registration trong Excel Options OPEN keys
foreach ($ver in @("16.0", "15.0", "14.0")) {
    $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path $optKey) {
        $props = (Get-ItemProperty -Path $optKey -ErrorAction SilentlyContinue).psobject.Properties
        $alreadyRegistered = $false
        $maxOpenIndex = -1
        
        foreach ($p in $props) {
            if ($p.Name -eq "OPEN" -or $p.Name -match "^OPEN(\d+)$") {
                $val = [string]$p.Value
                if ($val -like "*KangatangGuard.xlam*") {
                    $alreadyRegistered = $true
                    break
                }
                if ($p.Name -eq "OPEN") {
                    if ($maxOpenIndex -lt 0) { $maxOpenIndex = 0 }
                } elseif ($p.Name -match "^OPEN(\d+)$") {
                    $idxNum = [int]$matches[1]
                    if ($idxNum -gt $maxOpenIndex) { $maxOpenIndex = $idxNum }
                }
            }
        }
        
        if (-not $alreadyRegistered) {
            $targetSlot = if ($maxOpenIndex -lt 0) { "OPEN" } else { "OPEN$($maxOpenIndex + 1)" }
            $regValue = "/R `"$dstXlamAddIns`""
            Set-ItemProperty -Path $optKey -Name $targetSlot -Value $regValue -Type String -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   CAI DAT KANGATANG GUARD v$versionInfo HOAN TAT!                    " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   May chu cap nhat  : $uncSource" -ForegroundColor Cyan
Write-Host "   Phien ban da cai  : v$versionInfo" -ForegroundColor Cyan
Write-Host ""
Write-Host "   Dac diem he thong:" -ForegroundColor Yellow
Write-Host "   1. Mo Excel binh thuong, Add-in da duoc nap vao thanh cong cu." -ForegroundColor White
Write-Host "   2. Khong phu thuoc mang: Excel luon mo tuc thi trong 0.05s (Offline-First)." -ForegroundColor White
Write-Host "   3. Khi may chu phat hanh ban moi, may cua ban se TU DONG duoc cap nhat!" -ForegroundColor White
Write-Host "   4. Co the bam nut 'Kiem tra cap nhat tu May chu...' trong menu bat ky luc nao." -ForegroundColor White
Write-Host ""

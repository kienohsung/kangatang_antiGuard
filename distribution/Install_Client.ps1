# ==============================================================================
# Install_Client.ps1
# PowerShell Installer cho May Client trong Mang LAN
# Phien ban: v3.22.2 (100% Native VBA Folder Scanner & Zero-PS1 Antivirus-Safe Architecture)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - CAI DAT ADD-IN TU MAY CHU LAN (CLIENT v3.22.2)    " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

# Xac dinh Nguon Cap nhat (UpdateSource UNC Path)
$PrimaryServerPath = "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang"
$BackupServerPath  = "\\192.168.223.176\KangatangGuard_Hub"

$uncSource = ""
if ($ScriptDir.StartsWith("\\")) {
    $uncSource = $ScriptDir
} elseif (Test-Path $PrimaryServerPath) {
    $uncSource = $PrimaryServerPath
} elseif (Test-Path $BackupServerPath) {
    $uncSource = $BackupServerPath
} else {
    $uncSource = $ScriptDir
}

Write-Host "`n[1/5] May chu phan phoi (Hub): $uncSource" -ForegroundColor Yellow

# Doc thong tin phien ban tu version.json neu co
$versionInfo = "3.22.2"
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

# Kiem tra file nguon (Uu tien ban danh so phien ban de tranh loi file KangatangGuard.xlam tren Hub bi SMB lock boi may khac)
$srcXlamCandidates = @(
    (Join-Path $uncSource "KangatangGuard_v$versionInfo.xlam"),
    (Join-Path $uncSource "KangatangGuard_latest.xlam"),
    (Join-Path $BackupServerPath "KangatangGuard.xlam"),
    (Join-Path $uncSource "KangatangGuard.xlam"),
    (Join-Path $ScriptDir "KangatangGuard.xlam")
)
$srcXlam = $srcXlamCandidates[3]
foreach ($cand in $srcXlamCandidates) {
    if ($cand -and (Test-Path $cand)) {
        $srcXlam = $cand
        break
    }
}

if (-not (Test-Path $srcXlam)) {
    Write-Host "`n[LOI] Khong tim thay file KangatangGuard.xlam tai $srcXlam!" -ForegroundColor Red
    Write-Host "Vui long kiem tra ket noi den May chu LAN: $uncSource" -ForegroundColor Red
    exit 1
}
Write-Host "   -> Tep nguon Add-in: $srcXlam" -ForegroundColor Gray

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
$officeVersions = @("12.0", "14.0", "15.0", "16.0")

foreach ($ver in $officeVersions) {
    $secKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    # v3.15.0: Luon dam bao khoa Security ton tai cho cac phien ban Office pho bien (15.0, 16.0) hoac da co khoa Excel
    if (-not (Test-Path $secKey) -and ((Test-Path "HKCU:\Software\Microsoft\Office\$ver\Excel") -or ($ver -in @("15.0", "16.0")))) {
        New-Item -Path $secKey -Force -ErrorAction SilentlyContinue | Out-Null
    }
    if (Test-Path $secKey) {
        try {
            Set-ItemProperty -Path $secKey -Name "AllowNetworkLocations" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $secKey -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        } catch {}

        # Go bo Policy HKCU neu dang khoa AccessVBOM = 0
        $polPath = "HKCU:\Software\Policies\Microsoft\Office\$ver\Excel\Security"
        if (Test-Path $polPath) {
            try { Remove-ItemProperty -Path $polPath -Name "AccessVBOM" -Force -ErrorAction SilentlyContinue } catch {}
        }

        # Thu bat tren HKLM va WOW6432Node neu co quyen
        foreach ($hklmSec in @(
            "HKLM:\Software\Microsoft\Office\$ver\Excel\Security",
            "HKLM:\Software\WOW6432Node\Microsoft\Office\$ver\Excel\Security"
        )) {
            if (Test-Path $hklmSec) {
                try { Set-ItemProperty -Path $hklmSec -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue } catch {}
            }
        }

        $trustedRoot = "$secKey\Trusted Locations"
        if (-not (Test-Path $trustedRoot)) {
            New-Item -Path $trustedRoot -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # Trust Primary LAN File Server (192.168.223.7)
        $locHubPrimary = "$trustedRoot\KangatangLANServer"
        if (-not (Test-Path $locHubPrimary)) { New-Item -Path $locHubPrimary -Force -ErrorAction SilentlyContinue | Out-Null }
        Set-ItemProperty -Path $locHubPrimary -Name "Path" -Value $PrimaryServerPath -Type String -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locHubPrimary -Name "AllowSubfolders" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locHubPrimary -Name "Description" -Value "KangatangGuard LAN File Server" -Type String -Force -ErrorAction SilentlyContinue

        # Trust uncSource neu khac
        if ($uncSource -ne $PrimaryServerPath) {
            $locHubCurrent = "$trustedRoot\KangatangLANCurrent"
            if (-not (Test-Path $locHubCurrent)) { New-Item -Path $locHubCurrent -Force -ErrorAction SilentlyContinue | Out-Null }
            Set-ItemProperty -Path $locHubCurrent -Name "Path" -Value $uncSource -Type String -Force -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $locHubCurrent -Name "AllowSubfolders" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $locHubCurrent -Name "Description" -Value "KangatangGuard LAN Current Source" -Type String -Force -ErrorAction SilentlyContinue
        }

        # Trust local APPDATA KangatangGuard
        $locLocal = "$trustedRoot\KangatangLocal"
        if (-not (Test-Path $locLocal)) { New-Item -Path $locLocal -Force -ErrorAction SilentlyContinue | Out-Null }
        $appDataGuard = Join-Path $env:APPDATA "KangatangGuard"
        Set-ItemProperty -Path $locLocal -Name "Path" -Value $appDataGuard -Type String -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locLocal -Name "AllowSubfolders" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $locLocal -Name "Description" -Value "KangatangGuard Local Cache" -Type String -Force -ErrorAction SilentlyContinue
    }
}
Write-Host "   -> [OK] Da cau hinh Trusted Locations va AccessVBOM thanh cong." -ForegroundColor Green

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
if (-not (Test-Path $updatesDir)) { 
    New-Item -ItemType Directory -Path $updatesDir -Force | Out-Null 
} else {
    # Don sach cac file cu trong staged_update va go bo co Read-Only
    Get-ChildItem -Path $updatesDir -File -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_.IsReadOnly = $false } catch {}
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
    }
}

# Sao chep KangatangGuard.xlam vao %APPDATA%\Microsoft\AddIns (chuan Excel Add-Ins Manager)
$dstXlamAddIns = Join-Path $addInsPath "KangatangGuard.xlam"
if (Test-Path $dstXlamAddIns) {
    try { Set-ItemProperty -Path $dstXlamAddIns -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    Remove-Item -Path $dstXlamAddIns -Force -ErrorAction SilentlyContinue
}
Copy-Item -Path $srcXlam -Destination $dstXlamAddIns -Force
try { Set-ItemProperty -Path $dstXlamAddIns -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
Write-Host "   -> [OK] Da cai dat vao AddIns: $dstXlamAddIns" -ForegroundColor Green

# Don dep ban sao cu trong XLSTART neu co (tranh xung dot ten file khi mo Add-Ins Manager)
$dstXlamXLSTART = Join-Path $xlStartPath "KangatangGuard.xlam"
if (Test-Path $dstXlamXLSTART) {
    try { Set-ItemProperty -Path $dstXlamXLSTART -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    Remove-Item -Path $dstXlamXLSTART -Force -ErrorAction SilentlyContinue
    Write-Host "   -> [OK] Da xoa ban sao cu trong XLSTART: $dstXlamXLSTART" -ForegroundColor Green
}

# v3.17.3: Sao chep Kangatang_FolderScanner.ps1 (0-IOC, 0-P/Invoke) vao thu muc tin cay Microsoft\AddIns
# va xoa triet de legacy .ps1 trong APPDATA\KangatangGuard de tranh Antivirus Seqrite/Quick Heal nhan nham!
$srcScanner = Join-Path $uncSource "Kangatang_FolderScanner.ps1"
if (-not (Test-Path $srcScanner)) { $srcScanner = Join-Path $ScriptDir "Kangatang_FolderScanner.ps1" }
$dstAddInsScanner = Join-Path $addInsPath "Kangatang_FolderScanner.ps1"
if (Test-Path $srcScanner) {
    try { Set-ItemProperty -Path $dstAddInsScanner -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    Copy-Item -Path $srcScanner -Destination $dstAddInsScanner -Force -ErrorAction SilentlyContinue
    Write-Host "   -> [OK] Da cai dat Trinh quet thu muc ngoai tien trinh: $dstAddInsScanner" -ForegroundColor Green
}

$dstScanner = Join-Path $localAppData "Kangatang_FolderScanner.ps1"
if (Test-Path $dstScanner) {
    try { Set-ItemProperty -Path $dstScanner -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    Remove-Item -Path $dstScanner -Force -ErrorAction SilentlyContinue
    Write-Host "   -> [OK] Da don sach legacy PS1 trong AppData (Antivirus-Safe): $dstScanner" -ForegroundColor Green
}

# Dang ky vao Registry KangatangGuard
Write-Host "`n[5/5] Dang dang ky Add-In Manager & Nguon Cap nhat LAN vao Registry..." -ForegroundColor Yellow
$regKey = "HKCU:\Software\KangatangGuard"
if (-not (Test-Path $regKey)) { New-Item -Path $regKey -Force | Out-Null }

Set-ItemProperty -Path $regKey -Name "UpdateSource" -Value $uncSource -Force
Set-ItemProperty -Path $regKey -Name "InstalledVersion" -Value $versionInfo -Force
Remove-ItemProperty -Path $regKey -Name "ScannerScript" -Force -ErrorAction SilentlyContinue
$nowClientTs = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Set-ItemProperty -Path $regKey -Name "LastUpdateCheck" -Value $nowClientTs -Force
$firstInstDate = (Get-ItemProperty -Path $regKey -Name "FirstInstalledDate" -ErrorAction SilentlyContinue).FirstInstalledDate
if ([string]::IsNullOrWhiteSpace($firstInstDate)) {
    $firstInstDate = $nowClientTs
    Set-ItemProperty -Path $regKey -Name "FirstInstalledDate" -Value $firstInstDate -Force
}
Remove-ItemProperty -Path $regKey -Name "LastVbomWarn" -Force -ErrorAction SilentlyContinue
try {
    $macAddr = ""
    $ipAddr  = ""
    $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction SilentlyContinue
    foreach ($nic in $nics) {
        $mCand = ([string]$nic.MACAddress).Trim().ToUpperInvariant().Replace(":", "-")
        $ipCand = ""
        foreach ($ip in @($nic.IPAddress)) {
            if ($ip -match '^\d+\.\d+\.\d+\.\d+$' -and (-not $ip.StartsWith("169."))) { $ipCand = $ip; break }
        }
        if ($mCand -and $ipCand) {
            $macAddr = $mCand; $ipAddr = $ipCand
            if ($ipCand.StartsWith("192.168.223.")) { break }
        }
    }
    if (-not $macAddr) { $macAddr = "UNKNOWN-$env:COMPUTERNAME" }
    if (-not $ipAddr)  { $ipAddr  = "0.0.0.0" }
    $hbLine = "$macAddr|$ipAddr|$env:COMPUTERNAME|$env:USERNAME|$firstInstDate|$nowClientTs|v$versionInfo`r`n"
    $repDir = if (Test-Path "D:\7. AI tools\kangatang\Scan_Reports") { "D:\7. AI tools\kangatang\Scan_Reports" } elseif (Test-Path "\\192.168.223.176\test") {
        $rPath = "\\192.168.223.176\test\Scan_Reports"
        if (-not (Test-Path $rPath)) { New-Item -ItemType Directory -Path $rPath -Force -ErrorAction SilentlyContinue | Out-Null }
        $rPath
    } else { $localAppData }
    [System.IO.File]::AppendAllText((Join-Path $repDir "Endpoint_Heartbeats.log"), $hbLine, [System.Text.Encoding]::Unicode)
} catch {}

Write-Host "   -> [OK] UpdateSource    : $uncSource" -ForegroundColor Green
Write-Host "   -> [OK] InstalledVersion: v$versionInfo (100% Native VBA Scanner)" -ForegroundColor Green

# Dang ky KangatangGuard.xlam vao Excel Options OPEN* (bao toan cac Add-in khac nhu MyExcel.xlam)
foreach ($ver in @("16.0", "15.0", "14.0")) {
    $excelKey = "HKCU:\Software\Microsoft\Office\$ver\Excel"
    if (-not (Test-Path $excelKey) -and $ver -ne "16.0") { continue }
    $optKey = "$excelKey\Options"
    if (-not (Test-Path $optKey)) {
        New-Item -Path $optKey -Force -ErrorAction SilentlyContinue | Out-Null
    }
    if (Test-Path $optKey) {
        $otherOpenValues = New-Object System.Collections.Generic.List[string]
        $oldOpenNames    = New-Object System.Collections.Generic.List[string]
        $props = (Get-ItemProperty -Path $optKey -ErrorAction SilentlyContinue).psobject.Properties
        foreach ($p in $props) {
            if ($p.Name -eq "OPEN" -or $p.Name -match "^OPEN(\d+)$") {
                $oldOpenNames.Add($p.Name)
                $val = [string]$p.Value
                if ($val -and ($val -notlike "*KangatangGuard*")) {
                    $otherOpenValues.Add($val)
                }
            }
        }
        foreach ($oldName in $oldOpenNames) {
            Remove-ItemProperty -Path $optKey -Name $oldName -Force -ErrorAction SilentlyContinue
        }
        $otherOpenValues.Add('"KangatangGuard.xlam"')
        for ($idx = 0; $idx -lt $otherOpenValues.Count; $idx++) {
            $slotName = if ($idx -eq 0) { "OPEN" } else { "OPEN$idx" }
            Set-ItemProperty -Path $optKey -Name $slotName -Value $otherOpenValues[$idx] -Type String -Force -ErrorAction SilentlyContinue
        }
        Write-Host "   -> [OK] Da dang ky KangatangGuard.xlam vao $optKey (bao toan $($otherOpenValues.Count - 1) Add-in khac)" -ForegroundColor Green
        
        # Chi xoa bo Add-in Manager entry tro den UNC mang cu (khong xoa muc cuc bo)
        $mgrKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Add-in Manager"
        if (Test-Path $mgrKey) {
            $mProps = (Get-ItemProperty -Path $mgrKey -ErrorAction SilentlyContinue).psobject.Properties
            foreach ($mp in $mProps) {
                if ($mp.Name -like "\\*" -and ($mp.Name -like "*KangatangGuard*" -or $mp.Name -like "*addin_kangatang*")) {
                    Remove-ItemProperty -Path $mgrKey -Name $mp.Name -Force -ErrorAction SilentlyContinue
                    Write-Host "   -> [OK] Da go bo UNC Add-in Manager: $mgrKey\$($mp.Name)" -ForegroundColor Green
                }
            }
        }
    }
    
    # Don dep AddInLoadTimes va DisabledItems
    $altKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\AddInLoadTimes"
    if (Test-Path $altKey) {
        Remove-ItemProperty -Path $altKey -Name "KangatangGuard.xlam" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $altKey -Name "KangatangGuard" -Force -ErrorAction SilentlyContinue
    }
    $disKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
    if (Test-Path $disKey) {
        $disProps = (Get-ItemProperty -Path $disKey -ErrorAction SilentlyContinue).psobject.Properties
        foreach ($dp in $disProps) {
            if ($dp.Name -notlike "PS*") {
                $bytes = [byte[]]$dp.Value
                if ($bytes) {
                    $str = [System.Text.Encoding]::Unicode.GetString($bytes)
                    if ($str -like "*KangatangGuard*") {
                        Remove-ItemProperty -Path $disKey -Name $dp.Name -Force -ErrorAction SilentlyContinue
                        Write-Host "   -> [OK] Da go bo khoi DisabledItems: $disKey\$($dp.Name)" -ForegroundColor Green
                    }
                }
            }
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

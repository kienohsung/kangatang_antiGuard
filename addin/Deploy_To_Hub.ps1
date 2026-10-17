# ==============================================================================
# Deploy_To_Hub.ps1
# Lenh Phat hanh va Dong bo Phien ban moi len May chu LAN & Tao Clone Cuc bo
# Phien ban: v3.8.1 (Ho tro Bo cong cu Standalone & Dual-Mirror Release)
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

$ReleaseVersion = "3.17.0"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - PHAT HANH PHIEN BAN MOI (v$ReleaseVersion - DUAL-MIRROR) " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Bien dich Add-in .xlam tren May chu
Write-Host "`n[1/6] Dang bien dich KangatangGuard.xlam v$ReleaseVersion tu ma nguon VBA..." -ForegroundColor Yellow
$installerScript = Join-Path $AddinDir "Install_KangatangGuard.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $installerScript

$compiledXlam = Join-Path $env:APPDATA "Microsoft\AddIns\KangatangGuard.xlam"
if (-not (Test-Path $compiledXlam)) {
    Write-Host "`n[LOI] Khong tim thay KangatangGuard.xlam da bien dich tai $compiledXlam!" -ForegroundColor Red
    exit 1
}
Write-Host "   -> [OK] Bien dich thanh cong KangatangGuard.xlam" -ForegroundColor Green

# 2. Tao file sieu du lieu version.json
Write-Host "`n[2/6] Dang tao tệp sieu du lieu version.json..." -ForegroundColor Yellow
$versionObj = @{
    version           = $ReleaseVersion
    release_date      = (Get-Date -Format "yyyy-MM-dd")
    release_timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    hub_unc           = $PrimaryRemoteHub
    hub_ip            = "192.168.223.7"
    local_clone_host  = $env:COMPUTERNAME
    local_clone_ip    = "192.168.223.176"
    files             = @{
        xlam           = "KangatangGuard.xlam"
        scanner        = "Kangatang_FolderScanner.ps1"
        standalone     = "Kangatang_Standalone_Scanner.ps1"
        standalone_bat = "Chay_Diet_Virus_Ngoai.bat"
    }
    # v3.17.0: 100% Native VBA Folder Scanner & Zero-PS1 Antivirus-Safe Architecture
    mandatory         = $true
    changelog         = "v${ReleaseVersion} - 100% Native VBA Folder Scanner & Zero-PS1 Antivirus-Safe Architecture: Khac phuc triet de loi Antivirus (Seqrite Endpoint Security / Quick Heal) nhan nham va cach ly Kangatang_FolderScanner.ps1 trong %APPDATA%\KangatangGuard; tich hop trinh quet thu muc RunNativeFolderScan 100% Native VBA ben trong KangatangGuard.xlam, tu dong xoa bo file .ps1 cu trong %APPDATA%\KangatangGuard va loai bo toan bo ma C# P/Invoke trong cac script ho tro."
}

$versionJsonContent = $versionObj | ConvertTo-Json -Depth 4
$localJsonPath = Join-Path $LocalDist "version.json"
[System.IO.File]::WriteAllText($localJsonPath, $versionJsonContent, [System.Text.Encoding]::UTF8)
Write-Host "   -> [OK] Da tao version.json: $localJsonPath" -ForegroundColor Green

# 3. Dong bo tap tin vao Ban Clone Cuc bo (Local Clone tren may 223.176)
Write-Host "`n[3/6] Dang tao ban CLONE cuc bo tai may 223.176 ($LocalDist)..." -ForegroundColor Yellow
if (-not (Test-Path $LocalDist)) {
    New-Item -ItemType Directory -Path $LocalDist -Force | Out-Null
}

$targetLocalXlam = Join-Path $LocalDist "KangatangGuard.xlam"
if (Test-Path $targetLocalXlam) {
    try { Set-ItemProperty -Path $targetLocalXlam -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
}
Copy-Item -Path $compiledXlam -Destination $targetLocalXlam -Force
Copy-Item -Path $compiledXlam -Destination (Join-Path $LocalDist "KangatangGuard_latest.xlam") -Force
Copy-Item -Path $compiledXlam -Destination (Join-Path $LocalDist "KangatangGuard_v${ReleaseVersion}.xlam") -Force
try { Set-ItemProperty -Path $targetLocalXlam -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}

# Chep cac script vao distribution
Copy-Item -Path (Join-Path $AddinDir "Kangatang_Standalone_Scanner.ps1") -Destination (Join-Path $LocalDist "Kangatang_Standalone_Scanner.ps1") -Force
Copy-Item -Path (Join-Path $AddinDir "Kangatang_FolderScanner.ps1") -Destination (Join-Path $LocalDist "Kangatang_FolderScanner.ps1") -Force
if (Test-Path (Join-Path $AddinDir "Chay_Diet_Virus_Ngoai.bat")) {
    Copy-Item -Path (Join-Path $AddinDir "Chay_Diet_Virus_Ngoai.bat") -Destination (Join-Path $LocalDist "Chay_Diet_Virus_Ngoai.bat") -Force
} else {
    Copy-Item -Path (Join-Path $RootDir "Chay_Diet_Virus_Ngoai.bat") -Destination (Join-Path $LocalDist "Chay_Diet_Virus_Ngoai.bat") -Force
}

Write-Host "   -> [OK] Ban CLONE cuc bo tren may 223.176 da duoc dong bo day du." -ForegroundColor Green

# 4. Dong goi file zip phat hanh (Bao gom ca Add-in va Trinh chay ngoai doc lap)
Write-Host "`n[4/6] Dang dong goi file zip phat hanh..." -ForegroundColor Yellow
$zipFileVer = Join-Path $AddinDir "addin_kangatang_v${ReleaseVersion}.zip"
$zipFileLatest = Join-Path $AddinDir "addin_kangatang.zip"

if (Test-Path $zipFileVer) { Remove-Item $zipFileVer -Force }
if (Test-Path $zipFileLatest) { Remove-Item $zipFileLatest -Force }

# Thu muc tam dong goi
$stageZipDir = Join-Path $env:TEMP "KangatangZipStage"
if (Test-Path $stageZipDir) { Remove-Item $stageZipDir -Recurse -Force }
New-Item -ItemType Directory -Path $stageZipDir -Force | Out-Null

Copy-Item -Path (Join-Path $LocalDist "*") -Destination $stageZipDir -Recurse -Force
Compress-Archive -Path (Join-Path $stageZipDir "*") -DestinationPath $zipFileVer -Force
Copy-Item -Path $zipFileVer -Destination $zipFileLatest -Force
Remove-Item $stageZipDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "   -> [OK] Da tao file zip phat hanh: $([System.IO.Path]::GetFileName($zipFileVer))" -ForegroundColor Green

# 5. Phat hanh len May chu Tep LAN Chuyen dung (\\192.168.223.7)
Write-Host "`n[5/6] Dang phat hanh len May chu Tep LAN ($PrimaryRemoteHub)..." -ForegroundColor Yellow
if (-not (Test-Path $PrimaryRemoteHub)) {
    Write-Host "   -> [CANH BAO] Khong the ket noi den $PrimaryRemoteHub!" -ForegroundColor Red
    Write-Host "   -> Vui long kiem tra may 192.168.223.7 co dang ket noi mang khong." -ForegroundColor Red
} else {
    # Luon chep ban danh so phien ban (KHONG BAO GIO bi SMB lock boi Excel client)
    $remoteXlamVer    = Join-Path $PrimaryRemoteHub "KangatangGuard_v${ReleaseVersion}.xlam"
    $remoteXlamLatest = Join-Path $PrimaryRemoteHub "KangatangGuard_latest.xlam"
    Copy-Item -Path $compiledXlam -Destination $remoteXlamVer -Force
    Write-Host "   -> [OK] Da chep KangatangGuard_v${ReleaseVersion}.xlam len May chu 223.7" -ForegroundColor Green

    foreach ($remoteTarget in @($remoteXlamLatest, (Join-Path $PrimaryRemoteHub "KangatangGuard.xlam"))) {
        $targetName = [System.IO.Path]::GetFileName($remoteTarget)
        try {
            if (Test-Path $remoteTarget) {
                Set-ItemProperty -Path $remoteTarget -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
            }
            Copy-Item -Path $compiledXlam -Destination $remoteTarget -Force -ErrorAction Stop
            Set-ItemProperty -Path $remoteTarget -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
            Write-Host "   -> [OK] Da chep $targetName len May chu 223.7" -ForegroundColor Green
        } catch {
            Write-Host "   -> [LUU Y] $targetName tren 223.7 dang duoc mo boi 1 may tram LAN (SMB lock)." -ForegroundColor Yellow
            $bgSyncCmd = "for (`$i=0; `$i -lt 720; `$i++) { try { Copy-Item -LiteralPath '$targetLocalXlam' -Destination '$remoteTarget' -Force -ErrorAction Stop; break } catch { Start-Sleep -Seconds 30 } }"
            Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$bgSyncCmd`""
            Write-Host "   -> [OK] Da kich hoat Background Sync Watcher cho $targetName." -ForegroundColor Green
        }
    }

    # Chep cac tep ho tro
    Copy-Item -Path (Join-Path $LocalDist "Kangatang_FolderScanner.ps1") -Destination (Join-Path $PrimaryRemoteHub "Kangatang_FolderScanner.ps1") -Force
    Copy-Item -Path (Join-Path $LocalDist "Kangatang_Standalone_Scanner.ps1") -Destination (Join-Path $PrimaryRemoteHub "Kangatang_Standalone_Scanner.ps1") -Force
    Copy-Item -Path (Join-Path $LocalDist "Chay_Diet_Virus_Ngoai.bat") -Destination (Join-Path $PrimaryRemoteHub "Chay_Diet_Virus_Ngoai.bat") -Force
    Copy-Item -Path $localJsonPath -Destination (Join-Path $PrimaryRemoteHub "version.json") -Force
    Copy-Item -Path (Join-Path $LocalDist "Install_Client_Kangatang.bat") -Destination (Join-Path $PrimaryRemoteHub "Install_Client_Kangatang.bat") -Force
    Copy-Item -Path (Join-Path $LocalDist "Install_Client.ps1") -Destination (Join-Path $PrimaryRemoteHub "Install_Client.ps1") -Force
    Copy-Item -Path (Join-Path $LocalDist "README_HUONG_DAN_CLIENT.txt") -Destination (Join-Path $PrimaryRemoteHub "README_HUONG_DAN_CLIENT.txt") -Force

    # Dong bo file zip phat hanh
    Copy-Item -Path $zipFileVer -Destination (Join-Path $PrimaryRemoteHub "addin_kangatang_v${ReleaseVersion}.zip") -Force
    Copy-Item -Path $zipFileLatest -Destination (Join-Path $PrimaryRemoteHub "addin_kangatang.zip") -Force

    Write-Host "   -> [OK] Da dong bo toan bo goi cai dat & file zip len May chu 223.7 thanh cong!" -ForegroundColor Green
}

# 6. Dong bo sang Desktop Mirror
Write-Host "`n[6/6] Dang dong bo sang thu muc Desktop Mirror..." -ForegroundColor Yellow
if (Test-Path $DesktopDir) {
    Copy-Item (Join-Path $LocalDist "*") -Destination (Join-Path $DesktopDir "distribution") -Recurse -Force
    Copy-Item (Join-Path $AddinDir "*") -Destination (Join-Path $DesktopDir "addin") -Recurse -Force -Exclude "*.zip"
    Copy-Item -Path $zipFileVer -Destination (Join-Path $DesktopDir "addin") -Force
    Copy-Item -Path $zipFileLatest -Destination (Join-Path $DesktopDir "addin") -Force
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
Write-Host "   Goi phat hanh da bao gom ca Add-in va Trinh diet ngoai Standalone 1-Click!" -ForegroundColor White
Write-Host ""

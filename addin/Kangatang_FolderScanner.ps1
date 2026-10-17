# ==============================================================================
# Kangatang_FolderScanner.ps1
# Trinh Quet Thu Muc Ngoai Tien Trinh (Out-of-Process Streaming Console + Native VBA Worker)
# Phien ban: v3.17.3 (Real-Time File Streaming, Fast Resume & 100% Antivirus-Safe)
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = "",
    [Parameter(Mandatory=$false)]
    [switch]$Resume
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# 1. Dong bo phien ban dong tu version.json hoac Registry InstalledVersion
$AppVersion = "3.17.3"
$verJsonPath = Join-Path $ScriptDir "version.json"
if (Test-Path $verJsonPath) {
    try {
        $vMeta = Get-Content -Path $verJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($vMeta.version) { $AppVersion = $vMeta.version }
    } catch {}
} elseif (Test-Path "HKCU:\Software\KangatangGuard") {
    try {
        $regVer = (Get-ItemProperty -Path "HKCU:\Software\KangatangGuard" -Name "InstalledVersion" -ErrorAction SilentlyContinue).InstalledVersion
        if ($regVer) { $AppVersion = $regVer }
    } catch {}
}

$Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - Trinh Quet Thu Muc Ngoai Tien Trinh"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD v$AppVersion - OUT-OF-PROCESS FOLDER SCANNER       " -ForegroundColor Cyan
Write-Host "   Day la addin phong chong virus kangatang - Phat trien boi OhsungIT team !" -ForegroundColor White
Write-Host "======================================================================" -ForegroundColor Cyan

# 2. Dam bao AccessVBOM = 1 trong Registry HKCU
foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
    $secPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $secPath) {
        Set-ItemProperty -Path $secPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    }
}

# 3. Doc yeu cau quet tu scan_request.json (khi duoc goi tu Excel LaunchBackgroundScanner)
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$scanReqFile     = Join-Path $logDir "scan_request.json"
$lastSessionFile = Join-Path $logDir "last_session.json"
$scannedLogFile  = Join-Path $logDir "scanned_files.log"

$isResume = [bool]$Resume
if ((-not $TargetFolder) -and (Test-Path $scanReqFile)) {
    try {
        $reqMeta = Get-Content -Path $scanReqFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($reqMeta.TargetFolder) { $TargetFolder = [string]$reqMeta.TargetFolder }
        if ($reqMeta.Resume -eq $true) { $isResume = $true }
    } catch {}
    Remove-Item -Path $scanReqFile -Force -ErrorAction SilentlyContinue
}

if ($isResume -and (-not $TargetFolder) -and (Test-Path $lastSessionFile)) {
    try {
        $meta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($meta.TargetFolder) { $TargetFolder = [string]$meta.TargetFolder }
    } catch {}
}

if (-not $TargetFolder) {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "KangatangGuard v$AppVersion - Chon thu muc can quet:"
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TargetFolder = $dlg.SelectedPath
    } else {
        Write-Host "`n[HUY] Nguoi dung da huy chon thu muc." -ForegroundColor Yellow
        exit 0
    }
}

$TargetFolder = $TargetFolder.Trim()
while ($TargetFolder.Length -gt 3 -and $TargetFolder.EndsWith("\")) {
    $TargetFolder = $TargetFolder.Substring(0, $TargetFolder.Length - 1)
}

if (-not (Test-Path -LiteralPath $TargetFolder)) {
    Write-Host "`n[LOI] Khong tim thay thu muc can quet: $TargetFolder" -ForegroundColor Red
    exit 1
}

# 4. Xac dinh duong dan KangatangGuard.xlam
$xlamCandidates = @(
    (Join-Path $env:APPDATA "Microsoft\AddIns\KangatangGuard.xlam"),
    (Join-Path $ScriptDir "KangatangGuard.xlam"),
    "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam"
)
$xlamPath = $null
foreach ($c in $xlamCandidates) {
    if ($c -and (Test-Path -LiteralPath $c)) {
        $xlamPath = $c
        break
    }
}

if (-not $xlamPath) {
    Write-Host "`n[LOI] Khong tim thay file KangatangGuard.xlam!" -ForegroundColor Red
    exit 1
}

$modeLabel = if ($isResume) { "Tiep tuc phien truoc (Fast Resume)" } else { "Quet moi toan bo thu muc" }
Write-Host "`n   -> Thu muc muc tieu : $TargetFolder" -ForegroundColor Yellow
Write-Host "   -> Che do quet      : $modeLabel" -ForegroundColor Yellow
Write-Host "   -> Engine Add-in    : $xlamPath (v$AppVersion)" -ForegroundColor Gray

# 5. Nap du lieu Fast Resume tu scanned_files.log
$resumeMap = @{}
if ($isResume -and (Test-Path -LiteralPath $scannedLogFile)) {
    try {
        foreach ($line in [System.IO.File]::ReadAllLines($scannedLogFile)) {
            $trimmed = $line.Trim()
            if ($trimmed.Length -gt 0) {
                $idxSep = $trimmed.IndexOf("|")
                if ($idxSep -gt 0) {
                    $k = $trimmed.Substring(0, $idxSep).ToLowerInvariant()
                    $v = $trimmed.Substring($idxSep + 1)
                    $resumeMap[$k] = $v
                }
            }
        }
        Write-Host "   -> Fast Resume      : Da nap $($resumeMap.Count) tep da kiem tra an toan tu truoc." -ForegroundColor Cyan
    } catch {}
} else {
    try { [System.IO.File]::WriteAllText($scannedLogFile, "", [System.Text.Encoding]::Unicode) } catch {}
}

function Write-SessionProgress([string]$status, [int]$totalFiles, [int]$scanned, [int]$cleaned, [string]$startTime) {
    try {
        $obj = [ordered]@{
            TargetFolder = $TargetFolder
            StartTime    = $startTime
            LastUpdate   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Status       = $status
            TotalFiles   = $totalFiles
            TotalScanned = $scanned
            TotalCleaned = $cleaned
        }
        $json = $obj | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($lastSessionFile, $json, $utf8NoBom)
    } catch {}
}

function Test-SkippedFolderName([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return $true }
    $low = $name.Trim().ToLowerInvariant()
    if ($low.StartsWith(".") -or $low.StartsWith("$")) { return $true }
    if ($low -in @("_backup_kangatang", "quarantine_backup", "threat_samples", "node_modules", "system volume information")) { return $true }
    if ($low.StartsWith("_backup_source_")) { return $true }
    return $false
}

function Get-ScannableFolders([string]$rootPath) {
    $list = New-Object System.Collections.Generic.List[string]
    $queue = New-Object System.Collections.Generic.Queue[string]
    $queue.Enqueue($rootPath)
    while ($queue.Count -gt 0) {
        $cur = $queue.Dequeue()
        $list.Add($cur)
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($cur)
            foreach ($sd in $subDirs) {
                $dirName = [System.IO.Path]::GetFileName($sd)
                if (-not (Test-SkippedFolderName $dirName)) {
                    $queue.Enqueue($sd)
                }
            }
        } catch {}
    }
    return $list
}

# 6. Khoi tao Background Excel Worker
$script:excelWorker = $null
$script:wbAddin     = $null

function Start-ExcelWorker {
    Stop-ExcelWorker
    $script:excelWorker = New-Object -ComObject Excel.Application
    $script:excelWorker.Visible = $false
    $script:excelWorker.DisplayAlerts = $false
    $script:excelWorker.EnableEvents = $false
    $script:wbAddin = $script:excelWorker.Workbooks.Open($xlamPath, 0, $true)
    $script:excelWorker.Run("'$($script:wbAddin.Name)'!InitWorkerMode") | Out-Null
}

function Stop-ExcelWorker {
    if ($null -ne $script:wbAddin) {
        try { $script:wbAddin.Close($false) } catch {}
        $script:wbAddin = $null
    }
    if ($null -ne $script:excelWorker) {
        try { $script:excelWorker.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:excelWorker) | Out-Null } catch {}
        $script:excelWorker = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "`n[1/2] Dang khoi tao Background Excel Worker (v$AppVersion)..." -ForegroundColor Yellow
try {
    Start-ExcelWorker
    Write-Host "   -> [OK] Background Worker da san sang (Excel chinh cua ban 100% khong bi anh huong!)." -ForegroundColor Green
} catch {
    Write-Host "`n[LOI] Khong the khoi tao Excel Worker: $($_.Exception.Message)" -ForegroundColor Red
    Stop-ExcelWorker
    exit 1
}

Write-Host "`n[2/2] Dang duyet va quet truc tiep tung thu muc..." -ForegroundColor Yellow
$startTimeStr  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$folders       = Get-ScannableFolders $TargetFolder
$totalScanned  = 0
$totalSkipped  = 0
$totalInfected = 0
$totalCleaned  = 0
$totalErrors   = 0
$folderIdx     = 0
$validExts     = @(".xls", ".xlsx", ".xlsm", ".xlsb")

Write-SessionProgress "IN_PROGRESS" 0 0 0 $startTimeStr

try {
    foreach ($folder in $folders) {
        $excelFiles = @()
        try {
            $allFiles = [System.IO.Directory]::GetFiles($folder)
            foreach ($fPath in $allFiles) {
                $fName = [System.IO.Path]::GetFileName($fPath)
                $ext   = [System.IO.Path]::GetExtension($fName).ToLowerInvariant()
                if (($ext -in $validExts) -and (-not $fName.StartsWith('~$')) -and (-not $fName.ToLowerInvariant().StartsWith("kangatangguard"))) {
                    $excelFiles += $fPath
                }
            }
        } catch {
            continue
        }

        if ($excelFiles.Count -eq 0) { continue }
        $folderIdx++
        Write-Host "`n----------------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "Folder [$folderIdx]: $folder ($($excelFiles.Count) files)" -ForegroundColor Cyan

        foreach ($filePath in $excelFiles) {
            $fileName = [System.IO.Path]::GetFileName($filePath)
            $fInfo = $null
            try { $fInfo = New-Object System.IO.FileInfo($filePath) } catch { continue }
            if (-not $fInfo -or (-not $fInfo.Exists)) { continue }

            $fileSig = "$($fInfo.Length)|$($fInfo.LastWriteTime.ToString('yyyyMMddHHmmss'))"
            $lowPath = $filePath.ToLowerInvariant()

            if ($isResume -and $resumeMap.ContainsKey($lowPath) -and ($resumeMap[$lowPath] -eq $fileSig)) {
                $totalSkipped++
                $totalScanned++
                $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion | Da quet: $totalScanned (Bo qua: $totalSkipped) | Da diet: $totalCleaned"
                Write-Host "   [BO QUA] [$totalScanned] Da an toan tu truoc: $fileName" -ForegroundColor DarkGray
                continue
            }

            $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion | Da quet: $totalScanned | Da diet: $totalCleaned | Dang mo: $fileName"

            $res = ""
            try {
                $res = [string]$script:excelWorker.Run("'$($script:wbAddin.Name)'!ScanSingleFileForWorker", [string]$filePath)
            } catch {
                Write-Host "   [KHOI PHUC WORKER] Tep gay gian doan COM ($fileName), dang khoi dong lai Worker..." -ForegroundColor Magenta
                try { Start-ExcelWorker } catch {}
                $totalErrors++
                continue
            }

            if ($res -eq "SAFE") {
                $totalScanned++
                Write-Host "   [OK] [$totalScanned] An toan: $fileName" -ForegroundColor Green
                try {
                    [System.IO.File]::AppendAllText($scannedLogFile, "$lowPath|$fileSig`r`n", [System.Text.Encoding]::Unicode)
                } catch {}
            } elseif ($res.StartsWith("CLEANED|")) {
                $totalScanned++
                $totalInfected++
                $totalCleaned++
                $detail = $res.Substring(8)
                Write-Host "   [DIET VIRUS] [$totalScanned] Da lam sach & sao luu: $fileName ($detail)" -ForegroundColor Yellow
                try {
                    $fAfter = New-Object System.IO.FileInfo($filePath)
                    $newSig = "$($fAfter.Length)|$($fAfter.LastWriteTime.ToString('yyyyMMddHHmmss'))"
                    [System.IO.File]::AppendAllText($scannedLogFile, "$lowPath|$newSig`r`n", [System.Text.Encoding]::Unicode)
                } catch {}
            } elseif ($res.StartsWith("PURGE_NEEDED|")) {
                $totalScanned++
                $totalInfected++
                $detail = $res.Substring(13)
                $purgeRes = "MORE_PURGE"
                for ($attempt = 1; $attempt -le 5; $attempt++) {
                    try {
                        $script:excelWorker.Run("'$($script:wbAddin.Name)'!WorkerDeleteOneModuleSheet", [string]$filePath) | Out-Null
                    } catch {}
                    try {
                        $purgeRes = [string]$script:excelWorker.Run("'$($script:wbAddin.Name)'!WorkerFinishCleanAndClose", [string]$filePath)
                    } catch {
                        $purgeRes = "INFECTED_ERR|RPC error"
                        try { Start-ExcelWorker } catch {}
                        break
                    }
                    if ($purgeRes -ne "MORE_PURGE") { break }
                }
                if ($purgeRes.StartsWith("CLEANED|")) {
                    $totalCleaned++
                    Write-Host "   [DIET VIRUS] [$totalScanned] Da go Module-Sheet Excel 5.0 & lam sach: $fileName ($detail)" -ForegroundColor Yellow
                    try {
                        $fAfter = New-Object System.IO.FileInfo($filePath)
                        $newSig = "$($fAfter.Length)|$($fAfter.LastWriteTime.ToString('yyyyMMddHHmmss'))"
                        [System.IO.File]::AppendAllText($scannedLogFile, "$lowPath|$newSig`r`n", [System.Text.Encoding]::Unicode)
                    } catch {}
                } else {
                    $totalErrors++
                    Write-Host "   [CANH BAO] [$totalScanned] Khong the go Module-Sheet: $fileName ($purgeRes)" -ForegroundColor Red
                }
            } elseif ($res.StartsWith("INFECTED_")) {
                $totalScanned++
                $totalInfected++
                $totalErrors++
                Write-Host "   [CANH BAO] [$totalScanned] Phat hien virus nhung tep dang khoa/ReadOnly: $fileName ($res)" -ForegroundColor Red
            } else {
                $totalErrors++
                Write-Host "   [BO QUA] Khong the mo tep (co mat khau hoac dang khoa): $fileName" -ForegroundColor DarkYellow
            }

            if (($totalScanned % 5) -eq 0) {
                Write-SessionProgress "IN_PROGRESS" $totalScanned $totalScanned $totalCleaned $startTimeStr
            }
        }

        Write-SessionProgress "IN_PROGRESS" $totalScanned $totalScanned $totalCleaned $startTimeStr
    }

    Write-SessionProgress "COMPLETED" $totalScanned $totalScanned $totalCleaned $startTimeStr
    $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - HOAN TAT | Da quet: $totalScanned | Da diet: $totalCleaned"

    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host "   HOAN TAT QUET THU MUC (KANGATANG GUARD v$AppVersion)               " -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host "   Thu muc da quet       : $TargetFolder" -ForegroundColor White
    Write-Host "   Tong so tep kiem tra  : $totalScanned (Bo qua nhanh Fast Resume: $totalSkipped)" -ForegroundColor Cyan
    Write-Host "   Phat hien nhiem virus : $totalInfected" -ForegroundColor $(if ($totalInfected -gt 0) { "Yellow" } else { "Green" })
    Write-Host "   Da diet & sao luu     : $totalCleaned" -ForegroundColor Green
    Write-Host "   Tep bo qua (Password) : $totalErrors" -ForegroundColor Gray
    Write-Host "======================================================================" -ForegroundColor Green
} finally {
    Stop-ExcelWorker
}

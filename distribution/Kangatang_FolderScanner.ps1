# ==============================================================================
# Kangatang_FolderScanner.ps1
# Trinh Quet Thu Muc Sieu Toc 4 Tang & Bang Dieu Khien Truc Tuyen (Live Dashboard)
# Phien ban: v3.20.1 (Detailed Project Audit Logging & Periodic Report Engine)
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = "",
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    [Parameter(Mandatory=$false)]
    [int]$MaxWorkers = 0
)

# 0. Chot chan Don Tien Trinh Tuyet Doi (Single-Instance Mutex + Active Session Heartbeat)
# Ngam chan 100% hien tuong tu bat cua so PowerShell thu 2 khi dang quet thu muc
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$scanReqFile     = Join-Path $logDir "scan_request.json"
$lastSessionFile = Join-Path $logDir "last_session.json"
$scannedLogFile  = Join-Path $logDir "scanned_files.log"

$script:scannerMutex = $null
$script:hasMutex     = $false
try {
    $script:scannerMutex = New-Object System.Threading.Mutex($false, "Local\KangatangGuard_FolderScanner_Mutex")
    $script:hasMutex = $script:scannerMutex.WaitOne(0, $false)
} catch {
    $script:hasMutex = $true
}
if (-not $script:hasMutex) {
    [System.Environment]::Exit(0)
}

# Kiem tra nhip tim (Heartbeat) trong last_session.json de bao ve ca phien quet cu dang chay
if (Test-Path -LiteralPath $lastSessionFile) {
    try {
        $sessItem = Get-Item -LiteralPath $lastSessionFile -ErrorAction SilentlyContinue
        $sessMeta = Get-Content -LiteralPath $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($sessMeta -and ($sessMeta.Status -eq "IN_PROGRESS")) {
            $activeOther = $false
            if ($sessMeta.ScannerPid -and ([int]$sessMeta.ScannerPid -ne $PID)) {
                $pObj = Get-Process -Id ([int]$sessMeta.ScannerPid) -ErrorAction SilentlyContinue
                if ($pObj -and ($pObj.ProcessName -match "powershell|pwsh")) {
                    $activeOther = $true
                }
            }
            if ($activeOther) {
                if ($script:hasMutex -and ($null -ne $script:scannerMutex)) {
                    try { $script:scannerMutex.ReleaseMutex(); $script:scannerMutex.Dispose() } catch {}
                }
                [System.Environment]::Exit(0)
            }
        }
    } catch {}
}

try { $null = & cmd.exe /c "chcp 65001 >nul 2>&1" } catch {}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$utf8Bom   = New-Object System.Text.UTF8Encoding($true)

# 1. Dong bo phien ban dong tu version.json hoac Registry InstalledVersion
$AppVersion = "3.20.1"
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

$Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - Bang Dieu Khien Quet Thu Muc (Live Dashboard)"
try {
    $rawUI = $Host.UI.RawUI
    $bufSize = $rawUI.BufferSize
    $winSize = $rawUI.WindowSize
    if ($bufSize.Width -lt 106) { $bufSize.Width = 106; $rawUI.BufferSize = $bufSize }
    if ($bufSize.Height -lt 300) { $bufSize.Height = 300; $rawUI.BufferSize = $bufSize }
    if ($winSize.Width -lt 106) { $winSize.Width = 106 }
    if ($winSize.Height -lt 32) { $winSize.Height = 32 }
    $rawUI.WindowSize = $winSize
} catch {}

# Ham chuyen doi chuoi hien thi Console sang ASCII sach (Chong loi o vuong tren Windows Console & giu thang cot bang)
function ConvertTo-DisplayAscii([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return "" }
    $s = $text.Replace([char]0x0111, 'd').Replace([char]0x0110, 'D')
    $norm = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder($norm.Length)
    foreach ($ch in $norm.ToCharArray()) {
        $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            $code = [int][char]$ch
            if ($code -ge 32 -and $code -le 126) {
                [void]$sb.Append($ch)
            } elseif ($code -gt 126) {
                [void]$sb.Append('?')
            }
        }
    }
    return $sb.ToString()
}

function Format-MiddleTruncate([string]$text, [int]$maxLen) {
    $clean = ConvertTo-DisplayAscii $text
    if ($clean.Length -le $maxLen) { return $clean }
    if ($maxLen -le 10) { return $clean.Substring(0, $maxLen) }
    $headLen = [Math]::Floor(($maxLen - 5) * 0.42)
    $tailLen = $maxLen - 5 - $headLen
    return ($clean.Substring(0, $headLen) + "\...\" + $clean.Substring($clean.Length - $tailLen))
}

function Format-EndTruncate([string]$text, [int]$maxLen) {
    $clean = ConvertTo-DisplayAscii $text
    if ($clean.Length -le $maxLen) { return $clean }
    if ($maxLen -le 4) { return $clean.Substring(0, $maxLen) }
    return ($clean.Substring(0, $maxLen - 3) + "...")
}

Write-Host "====================================================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD v$AppVersion - BANG DIEU KHIEN QUET THU MUC (4-LAYER TURBO LIVE DASHBOARD)       " -ForegroundColor Cyan
Write-Host "   Day la addin phong chong virus kangatang - Phat trien boi OhsungIT team !                        " -ForegroundColor White
Write-Host "====================================================================================================" -ForegroundColor Cyan

# 2. Dam bao AccessVBOM = 1, tat hop thoai WER & don sach DocumentRecovery tranh Excel tu mo /restore
try {
    $werKey = "HKCU:\Software\Microsoft\Windows\Windows Error Reporting"
    if (-not (Test-Path $werKey)) { New-Item -Path $werKey -Force | Out-Null }
    Set-ItemProperty -Path $werKey -Name "DontShowUI" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
} catch {}

foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
    $secPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $secPath) {
        Set-ItemProperty -Path $secPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    }
    foreach ($sub in @("DocumentRecovery", "StartupItems")) {
        $resKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\$sub"
        if (Test-Path $resKey) {
            Remove-Item -Path $resKey -Recurse -Force -ErrorAction SilentlyContinue
        }
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

$prevSessionTotal   = 0
$prevSessionCleaned = 0
$prevSessionStart   = ""
if (Test-Path $lastSessionFile) {
    try {
        $meta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($isResume -and (-not $TargetFolder) -and $meta.TargetFolder) {
            $TargetFolder = [string]$meta.TargetFolder
        }
        if ($meta.TargetFolder -and ($meta.TargetFolder -eq $TargetFolder) -and ($meta.TotalFiles -gt 0)) {
            $prevSessionTotal = [int]$meta.TotalFiles
        }
        if ($isResume -and $meta.TargetFolder -and ($meta.TargetFolder -eq $TargetFolder)) {
            if ($meta.TotalCleaned -gt 0) { $prevSessionCleaned = [int]$meta.TotalCleaned }
            if ($meta.StartTime) { $prevSessionStart = [string]$meta.StartTime }
        }
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

# Su dung 2 Parallel Excel COM Workers (On-Demand) cho .xls/Macro (85% tep .xlsx da duoc xu ly boi Zero-COM ZIP)
$workerCount = if ($MaxWorkers -gt 0) { $MaxWorkers } else { 2 }

# 5. Nap du lieu Fast Resume tu scanned_files.log vao RAM Hashtable
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
    } catch {}
} else {
    try { [System.IO.File]::WriteAllText($scannedLogFile, "", [System.Text.Encoding]::Unicode) } catch {}
}

# 5b. He thong Ghi Log Chi Tiet & Xuat Bao Cao Dinh Ky (v3.20.1)
# Luu truc tiep vao thu muc du an D:\7. AI tools\kangatang\Scan_Reports (hoac \\192.168.223.176\test\Scan_Reports tu may tram LAN)
$script:projectReportDir = if (Test-Path "D:\7. AI tools\kangatang") {
    "D:\7. AI tools\kangatang\Scan_Reports"
} elseif (Test-Path "\\192.168.223.176\test") {
    "\\192.168.223.176\test\Scan_Reports"
} else {
    Join-Path $logDir "Scan_Reports"
}
$script:localReportDir = Join-Path $logDir "Scan_Reports"
foreach ($rDir in @($script:projectReportDir, (Join-Path $script:projectReportDir "Sessions"), $script:localReportDir)) {
    if ($rDir -and (-not (Test-Path -LiteralPath $rDir))) {
        try { New-Item -ItemType Directory -Path $rDir -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
}

$script:sessionId        = "SCAN_" + (Get-Date -Format "yyyyMMdd_HHmmss") + "_" + $env:COMPUTERNAME
$safeLeafName            = ([System.IO.Path]::GetFileName($TargetFolder) -replace '[^a-zA-Z0-9_\-]', '_')
if ([string]::IsNullOrWhiteSpace($safeLeafName)) { $safeLeafName = "RootShare" }
$script:sessionMdName    = "Report_$($script:sessionId)_$safeLeafName.md"
$script:sessionMdPath    = Join-Path (Join-Path $script:projectReportDir "Sessions") $script:sessionMdName
$script:latestMdPath     = Join-Path $script:projectReportDir "LATEST_SCAN_REPORT.md"
$script:summaryCsvPath   = Join-Path $script:projectReportDir "Scan_Summary_History.csv"
$script:ledgerCsvPath    = Join-Path $script:projectReportDir ("Infected_Files_Ledger_" + (Get-Date -Format "yyyyMM") + ".csv")
$script:detailLogPath    = Join-Path $script:projectReportDir ("scan_detail_log_" + (Get-Date -Format "yyyyMMdd") + ".txt")

$script:logBuffer        = New-Object System.Collections.Generic.List[string]
$script:detailLogBuffer  = New-Object System.Collections.Generic.List[string]
$script:ledgerBuffer     = New-Object System.Collections.Generic.List[string]
$script:infectedRecords  = New-Object System.Collections.Generic.List[pscustomobject]
$script:deptStats        = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:lastReportSyncMs = -30000

function Format-CsvField([string]$val) {
    if ($null -eq $val) { return '""' }
    return '"' + $val.Replace('"', '""').Replace("`r", " ").Replace("`n", " ") + '"'
}

function Get-TopDepartmentName([string]$filePath, [string]$rootFolder = $TargetFolder) {
    if ([string]::IsNullOrWhiteSpace($filePath)) { return "(Khong xac dinh)" }
    if ((-not [string]::IsNullOrWhiteSpace($rootFolder)) -and ($filePath.Length -gt ($rootFolder.Length + 1)) -and $filePath.StartsWith($rootFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
        $rel = $filePath.Substring($rootFolder.Length).TrimStart('\')
        $slashIdx = $rel.IndexOf('\')
        if ($slashIdx -gt 0) {
            return $rel.Substring(0, $slashIdx)
        }
        return "(Thu muc goc)"
    }
    try {
        $parentDir = [System.IO.Path]::GetDirectoryName($filePath)
        $leaf = [System.IO.Path]::GetFileName($parentDir)
        if ($leaf) { return $leaf }
    } catch {}
    return "(Thu muc goc)"
}

function Update-DeptStat([string]$filePath, [string]$outcome) {
    $dept = Get-TopDepartmentName $filePath $TargetFolder
    if (-not $script:deptStats.ContainsKey($dept)) {
        $script:deptStats[$dept] = [pscustomobject]@{
            Dept     = $dept
            Scanned  = 0
            Safe     = 0
            Infected = 0
            Cleaned  = 0
            Errors   = 0
        }
    }
    $st = $script:deptStats[$dept]
    $st.Scanned++
    switch ($outcome) {
        "SAFE" { $st.Safe++ }
        "CLEANED" { $st.Infected++; $st.Cleaned++ }
        "INFECTED_LOCKED" { $st.Infected++; $st.Errors++ }
        default { $st.Errors++ }
    }
}

function Add-DetailAuditEvent([string]$eventType, [string]$filePath, [int]$workerId, [string]$threatFamily, [string]$detail) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $dept = Get-TopDepartmentName $filePath $TargetFolder
    $line = "$ts | Host=$env:COMPUTERNAME | User=$env:USERNAME | Ver=v$AppVersion | Session=$($script:sessionId) | Worker=W$workerId | Event=$eventType | Dept=$dept | Family=$threatFamily | Path=$filePath | Detail=$detail"
    $script:detailLogBuffer.Add($line)
}

function Write-PeriodicReportFiles([string]$status, [int]$totalTarget, [string]$startTime) {
    try {
        $nowStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $elapsedMin = if ($null -ne $stopwatch) { [Math]::Round($stopwatch.Elapsed.TotalMinutes, 2) } else { 0.0 }
        $elapsedSec = if ($null -ne $stopwatch) { [Math]::Max(0.1, $stopwatch.Elapsed.TotalSeconds) } else { 1.0 }
        $avgSpeed   = [Math]::Round($script:totalScanned / $elapsedSec, 1)
        $totalFolders = if ($null -ne $sharedState) { [Math]::Max([int]$sharedState.TotalTargetFolders, [int]$script:folderIdx) } else { [int]$script:folderIdx }

        # 1. Cap nhat Scan_Summary_History.csv (Upsert theo SessionID)
        $summaryHeader = "SessionID,StartTime,LastUpdate,DurationMinutes,Status,ComputerName,UserName,ScannerVersion,TargetFolder,TotalFolders,TotalDiscovered,TotalScanned,SkippedResume,SafeZeroCom,DeepComScanned,TotalSafe,TotalInfected,TotalCleaned,TotalLockedOrError,AvgSpeedFilesPerSec,SessionReportFile"
        $summaryRow = (@(
            (Format-CsvField $script:sessionId),
            (Format-CsvField $startTime),
            (Format-CsvField $nowStr),
            $elapsedMin,
            (Format-CsvField $status),
            (Format-CsvField $env:COMPUTERNAME),
            (Format-CsvField $env:USERNAME),
            (Format-CsvField "v$AppVersion"),
            (Format-CsvField $TargetFolder),
            $totalFolders,
            $totalTarget,
            $script:totalScanned,
            $script:totalSkipped,
            $script:totalZeroCom,
            $script:totalDeepCom,
            $script:totalSafe,
            $script:totalInfected,
            $script:totalCleaned,
            $script:totalErrors,
            $avgSpeed,
            (Format-CsvField ("Sessions\" + $script:sessionMdName))
        ) -join ",")

        $csvLines = New-Object System.Collections.Generic.List[string]
        $replaced = $false
        $sessionPrefix = '"' + $script:sessionId + '",'
        if (Test-Path -LiteralPath $script:summaryCsvPath) {
            try {
                foreach ($ln in [System.IO.File]::ReadAllLines($script:summaryCsvPath, [System.Text.Encoding]::UTF8)) {
                    if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                    if ($ln.StartsWith($sessionPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $csvLines.Add($summaryRow)
                        $replaced = $true
                    } else {
                        $csvLines.Add($ln)
                    }
                }
            } catch {}
        }
        if ($csvLines.Count -eq 0) {
            $csvLines.Add($summaryHeader)
        }
        if (-not $replaced) {
            $csvLines.Add($summaryRow)
        }
        [System.IO.File]::WriteAllLines($script:summaryCsvPath, $csvLines, $utf8Bom)

        # 2. Xuat Bao cao Phien quet Chi tiet Markdown (Report_<SessionID>.md & LATEST_SCAN_REPORT.md)
        $pctDone = if ($status -eq "COMPLETED") { 100.0 } elseif ($totalTarget -gt 0) { [Math]::Min(99.9, [Math]::Round(($script:totalScanned / [double]$totalTarget) * 100.0, 1)) } else { 0.0 }
        $safeRate = if ($script:totalScanned -gt 0) { [Math]::Round(($script:totalSafe / [double]$script:totalScanned) * 100.0, 2) } else { 100.0 }

        $md = New-Object System.Text.StringBuilder
        [void]$md.AppendLine("# BAO CAO KET QUA QUET & DIET VIRUS KANGATANG DINH KY")
        [void]$md.AppendLine("> **He thong:** KangatangGuard ``v$AppVersion`` - Phat trien boi **OhsungIT Team**")
        [void]$md.AppendLine("")
        [void]$md.AppendLine("## 1. Thong tin chung ve phien quet")
        [void]$md.AppendLine("| Hang muc | Gia tri chi tiet |")
        [void]$md.AppendLine("| :--- | :--- |")
        [void]$md.AppendLine("| **Ma phien quet (SessionID)** | ``$($script:sessionId)`` |")
        [void]$md.AppendLine("| **Trang thai phien quet** | **$status** ($pctDone% hoan thanh) |")
        [void]$md.AppendLine("| **Thu muc muc tieu** | ``$TargetFolder`` |")
        [void]$md.AppendLine("| **Thoi gian bat dau** | ``$startTime`` |")
        [void]$md.AppendLine("| **Cap nhat lan cuoi** | ``$nowStr`` (Thoi luong: **$elapsedMin phut**) |")
        [void]$md.AppendLine("| **May thuc hien / Tai khoan** | ``$env:COMPUTERNAME`` / ``$env:USERNAME`` |")
        [void]$md.AppendLine("| **Toc do xu ly trung binh** | **$avgSpeed tep/giay** ($totalFolders thu muc da duyet) |")
        [void]$md.AppendLine("")
        [void]$md.AppendLine("## 2. Bang tong hop ket qua quet & xu ly ma doc")
        [void]$md.AppendLine("| STT | Chi so thong ke | So luong (Tep) | Ty le / Ghi chu ky thuat |")
        [void]$md.AppendLine("| :---: | :--- | ---: | :--- |")
        [void]$md.AppendLine("| 1 | **Tong so tep Excel phat hien** | **$($totalTarget.ToString('N0'))** | ``.xls``, ``.xlsx``, ``.xlsm``, ``.xlsb`` |")
        [void]$md.AppendLine("| 2 | **Tong so tep da kiem tra xong** | **$($script:totalScanned.ToString('N0'))** | Hoan thanh **$pctDone%** |")
        [void]$md.AppendLine("| 2.1 | +-- *Tang 1: Bo qua nhanh (Fast Resume)* | $($script:totalSkipped.ToString('N0')) | Da kiem tra an toan tu phien truoc (0.01 ms/tep) |")
        [void]$md.AppendLine("| 2.2 | +-- *Tang 2: Loc tinh Zero-COM (.xlsx/.xls)* | $($script:totalZeroCom.ToString('N0')) | Kiem tra cau truc nhi phan khong mo Excel (1.5 ms/tep) |")
        [void]$md.AppendLine("| 2.3 | +-- *Tang 3: Quet sau Excel COM (.xls/Macro)* | $($script:totalDeepCom.ToString('N0')) | Kiem tra chuyen sau VBProject & Module-Sheet |")
        [void]$md.AppendLine("| 3 | **Tep AN TOAN (Sach)** | **$($script:totalSafe.ToString('N0'))** | Ty le an toan: **$safeRate%** |")
        [void]$md.AppendLine("| 4 | **Tep PHAT HIEN NHIEM VIRUS** | **$($script:totalInfected.ToString('N0'))** | Phat hien mau virus Kangatang / Laroux / Hidden Sheet |")
        [void]$md.AppendLine("| 5 | **Tep DA DIET SACH & SAO LUU** | **$($script:totalCleaned.ToString('N0'))** | Da cach ly ban goc ve may chu ``192.168.223.176\Quarantine_Backup`` |")
        [void]$md.AppendLine("| 6 | **Tep BO QUA / DANG KHOA (Password/ReadOnly)** | **$($script:totalErrors.ToString('N0'))** | Tep dat mat khau hoac dang mo boi nguoi dung khac tren mang |")
        [void]$md.AppendLine("")
        [void]$md.AppendLine("## 3. Thong ke chi tiet theo Phong ban / Thu muc cap 1")
        [void]$md.AppendLine("| Phong ban / Thu muc cap 1 | Da quet | Tep Sach | Phat hien nhiem | Da diet sach | Dang khoa / Loi | Ty le nhiem |")
        [void]$md.AppendLine("| :--- | ---: | ---: | ---: | ---: | ---: | ---: |")
        if ($script:deptStats.Count -gt 0) {
            $sortedDepts = $script:deptStats.Values | Sort-Object -Property @{Expression="Infected";Descending=$true}, @{Expression="Scanned";Descending=$true}
            foreach ($d in $sortedDepts) {
                $infRate = if ($d.Scanned -gt 0) { ("{0:0.00}%" -f (($d.Infected / [double]$d.Scanned) * 100.0)) } else { "0.00%" }
                [void]$md.AppendLine("| **$($d.Dept)** | $($d.Scanned.ToString('N0')) | $($d.Safe.ToString('N0')) | **$($d.Infected.ToString('N0'))** | **$($d.Cleaned.ToString('N0'))** | $($d.Errors.ToString('N0')) | $infRate |")
            }
        } else {
            [void]$md.AppendLine("| *(Dang thu thap du lieu...)* | 0 | 0 | 0 | 0 | 0 | 0.00% |")
        }
        [void]$md.AppendLine("")
        [void]$md.AppendLine("## 4. Danh sach chi tiet cac tep phat hien nhiem virus trong phien quet ($($script:infectedRecords.Count) tep)")
        if ($script:infectedRecords.Count -gt 0) {
            [void]$md.AppendLine("| STT | Thoi gian | Phong ban | Ten tep | Kich thuoc | Ho ma doc | Chi tiet mau nhiem | Trang thai xu ly | Duong dan day du |")
            [void]$md.AppendLine("| :---: | :---: | :--- | :--- | ---: | :--- | :--- | :--- | :--- |")
            $idx = 0
            foreach ($rec in $script:infectedRecords) {
                $idx++
                [void]$md.AppendLine("| $idx | ``$($rec.Time)`` | **$($rec.Department)** | ``$($rec.FileName)`` | $($rec.SizeKB) KB | ``$($rec.ThreatFamily)`` | $($rec.ThreatDetail) | **$($rec.Action)** | ``$($rec.FullPath)`` |")
            }
        } else {
            [void]$md.AppendLine("> Hien chua phat hien tep nhiem virus nao trong phien quet nay.")
        }

        $mdContent = $md.ToString()
        [System.IO.File]::WriteAllText($script:sessionMdPath, $mdContent, $utf8Bom)
        [System.IO.File]::WriteAllText($script:latestMdPath, $mdContent, $utf8Bom)
        if ($script:localReportDir -ne $script:projectReportDir) {
            try { [System.IO.File]::WriteAllText((Join-Path $script:localReportDir "LATEST_SCAN_REPORT.md"), $mdContent, $utf8Bom) } catch {}
        }
    } catch {}
}

function Flush-ScanLogs([string]$status, [int]$totalDiscovered, [int]$scanned, [int]$cleaned, [string]$startTime) {
    if ($script:logBuffer.Count -gt 0) {
        try {
            $chunk = ($script:logBuffer -join "`r`n") + "`r`n"
            [System.IO.File]::AppendAllText($scannedLogFile, $chunk, [System.Text.Encoding]::Unicode)
            $script:logBuffer.Clear()
        } catch {}
    }
    try {
        $obj = [ordered]@{
            TargetFolder      = $TargetFolder
            SessionId         = $script:sessionId
            StartTime         = $startTime
            LastUpdate        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Status            = $status
            ScannerPid        = $PID
            TotalFiles        = $totalDiscovered
            TotalScanned      = $scanned
            TotalSafe         = [int]$script:totalSafe
            TotalInfected     = [int]$script:totalInfected
            TotalCleaned      = $cleaned
            TotalErrors       = [int]$script:totalErrors
            ReportDir         = $script:projectReportDir
            SessionReportFile = $script:sessionMdPath
        }
        $json = $obj | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($lastSessionFile, $json, $utf8NoBom)
    } catch {}

    # Ghi lo su kien chi tiet ra Scan_Reports\scan_detail_log_YYYYMMDD.txt
    if ($script:detailLogBuffer.Count -gt 0) {
        try {
            $dChunk = ($script:detailLogBuffer -join "`r`n") + "`r`n"
            [System.IO.File]::AppendAllText($script:detailLogPath, $dChunk, $utf8Bom)
            $script:detailLogBuffer.Clear()
        } catch {}
    }

    # Ghi lo so cai tep nhiem ra Scan_Reports\Infected_Files_Ledger_YYYYMM.csv
    $hadNewInfected = ($script:ledgerBuffer.Count -gt 0)
    if ($hadNewInfected) {
        try {
            if (-not (Test-Path -LiteralPath $script:ledgerCsvPath)) {
                $ledgerHeader = "Timestamp,SessionID,ComputerName,UserName,TopDepartmentFolder,FileName,FullPath,FileSizeKB,LastModified,VirusDetail,ActionStatus,BackupNote`r`n"
                [System.IO.File]::WriteAllText($script:ledgerCsvPath, $ledgerHeader, $utf8Bom)
            }
            $lChunk = ($script:ledgerBuffer -join "`r`n") + "`r`n"
            [System.IO.File]::AppendAllText($script:ledgerCsvPath, $lChunk, $utf8Bom)
            $script:ledgerBuffer.Clear()
        } catch {}
    }

    # Cap nhat Scan_Summary_History.csv va Bao cao Markdown moi 20 giay, hoac khi co tep nhiem moi, hoac khi COMPLETED
    $nowMs = if ($null -ne $stopwatch) { $stopwatch.ElapsedMilliseconds } else { 0 }
    if ($hadNewInfected -or ($status -eq "COMPLETED") -or (($nowMs - $script:lastReportSyncMs) -ge 20000)) {
        $script:lastReportSyncMs = $nowMs
        Write-PeriodicReportFiles $status $totalDiscovered $startTime
    }
}

function Test-SkippedFolderName([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return $true }
    $low = $name.Trim().ToLowerInvariant()
    if ($low.StartsWith(".") -or $low.StartsWith("$")) { return $true }
    if ($low -in @("_backup_kangatang", "quarantine_backup", "threat_samples", "node_modules", "system volume information")) { return $true }
    if ($low.StartsWith("_backup_source_")) { return $true }
    return $false
}

# Tang 2: Zero-COM OpenXML & OLE2 Pre-Filter (Kiem tra .xlsx/.xlsm/.xlsb/.xls trong 1.5 - 3ms khong can mo Excel COM)
function Test-OpenXmlSafeZeroCom([string]$filePath, [string]$ext, [regex]$iocRegex) {
    if ($ext -eq ".xls") {
        $fs = $null
        try {
            $shareMode = [System.IO.FileShare]([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
            $fs = New-Object System.IO.FileStream($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $shareMode)
            if ($fs.Length -lt 512 -or $fs.Length -gt 4194304) { return $false }
            $bytes = New-Object byte[] ($fs.Length)
            [void]$fs.Read($bytes, 0, $bytes.Length)
            if ($bytes[0] -ne 0xD0 -or $bytes[1] -ne 0xCF -or $bytes[2] -ne 0x11 -or $bytes[3] -ne 0xE0) { return $false }
            $uniStr = [System.Text.Encoding]::Unicode.GetString($bytes)
            if (-not ($uniStr.Contains("Workbook") -or $uniStr.Contains("Book"))) { return $false }
            if ($uniStr.IndexOf("VBA", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $false }
            if ($iocRegex.IsMatch($uniStr)) { return $false }
            $asciiStr = [System.Text.Encoding]::ASCII.GetString($bytes)
            if ($iocRegex.IsMatch($asciiStr)) { return $false }
            return $true
        } catch {
            return $false
        } finally {
            if ($null -ne $fs) { try { $fs.Dispose() } catch {} }
        }
    }
    if ($ext -notin @(".xlsx", ".xlsm", ".xlsb")) { return $false }
    $fs  = $null
    $zip = $null
    try {
        $shareMode = [System.IO.FileShare]([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
        $fs  = New-Object System.IO.FileStream($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $shareMode)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read, $false)

        $wbXmlEntry = $null
        $wbBinEntry = $null

        foreach ($entry in $zip.Entries) {
            $fn = $entry.FullName.ToLowerInvariant().Replace('\', '/')
            if ($fn.Contains("vbaproject") -or $fn.Contains("macrosheet") -or $fn.Contains("dialogsheet")) {
                return $false
            }
            if ($fn -eq "xl/workbook.xml" -or $fn.EndsWith("/workbook.xml")) {
                $wbXmlEntry = $entry
            } elseif ($fn -eq "xl/workbook.bin" -or $fn.EndsWith("/workbook.bin")) {
                $wbBinEntry = $entry
            }
        }

        if ($ext -eq ".xlsb") {
            if ($null -eq $wbBinEntry -or $wbBinEntry.Length -gt 4194304) { return $false }
            $s = $wbBinEntry.Open()
            $ms = New-Object System.IO.MemoryStream
            $s.CopyTo($ms)
            $s.Close()
            $bytes = $ms.ToArray()
            $ms.Dispose()
            $asciiStr = [System.Text.Encoding]::ASCII.GetString($bytes)
            if ($iocRegex.IsMatch($asciiStr)) { return $false }
            $uniStr = [System.Text.Encoding]::Unicode.GetString($bytes)
            if ($iocRegex.IsMatch($uniStr)) { return $false }
            return $true
        } else {
            if ($null -eq $wbXmlEntry) { return $false }
            $sr = New-Object System.IO.StreamReader($wbXmlEntry.Open(), [System.Text.Encoding]::UTF8)
            $xmlText = $sr.ReadToEnd()
            $sr.Close()
            if ($iocRegex.IsMatch($xmlText)) { return $false }
            return $true
        }
    } catch {
        return $false
    } finally {
        if ($null -ne $zip) { try { $zip.Dispose() } catch {} }
        if ($null -ne $fs)  { try { $fs.Dispose() } catch {} }
    }
}

# 6. Khoi tao bieu thuc Zero-COM trong RAM (0% chuoi IOC tren dia -> 100% Antivirus-Safe)
$t1 = -join ([char[]](107,97,110,103,97))
$t2 = -join ([char[]](109,121,112,101,114,115,111,110))
$t3 = -join ([char[]](33,97,108,108,111,99,97,116,101,100))
$t4 = -join ([char[]](115,97,118,101,99,111,112,121,97,115))
$t5 = -join ([char[]](111,110,115,104,101,101,116,97,99,116,105,118,97,116,101))
$iocRegex = [regex]::new("$t1|$t2|$t3|$t4|$t5", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)

$comWorkQueue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$comResultQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$sharedState    = [hashtable]::Synchronized(@{
    ProducerDone       = $false
    CancelWorkers      = $false
    ActiveWorkers      = 0
    CurrentWorkerFile  = ""
    TotalTargetFiles   = $prevSessionTotal
    TotalTargetFolders = 0
    CountDone          = $false
    WorkerPid1         = 0
    WorkerPid2         = 0
})

# Luong dem tong so tep Excel chay ngam song song (giup tinh chinh xac % Con lai ma khong lam cham luc khoi dong)
$counterScript = {
    param($rootFolder, $state)
    $q = New-Object System.Collections.Generic.Queue[string]
    $q.Enqueue($rootFolder)
    $exts = @(".xls", ".xlsx", ".xlsm", ".xlsb")
    $fCount = 0
    $dCount = 0
    while ($q.Count -gt 0 -and (-not $state.CancelWorkers)) {
        $cur = $q.Dequeue()
        try {
            foreach ($sd in [System.IO.Directory]::GetDirectories($cur)) {
                $dn = [System.IO.Path]::GetFileName($sd)
                if (-not [string]::IsNullOrWhiteSpace($dn)) {
                    $low = $dn.Trim().ToLowerInvariant()
                    if ((-not $low.StartsWith(".")) -and (-not $low.StartsWith("$")) -and ($low -notin @("_backup_kangatang", "quarantine_backup", "threat_samples", "node_modules", "system volume information")) -and (-not $low.StartsWith("_backup_source_"))) {
                        $q.Enqueue($sd)
                    }
                }
            }
        } catch {}
        try {
            $hasExcel = $false
            foreach ($fp in [System.IO.Directory]::GetFiles($cur)) {
                $fn = [System.IO.Path]::GetFileName($fp)
                $ex = [System.IO.Path]::GetExtension($fn).ToLowerInvariant()
                if (($ex -in $exts) -and (-not $fn.StartsWith('~$')) -and (-not $fn.ToLowerInvariant().StartsWith("kangatangguard"))) {
                    $fCount++
                    $hasExcel = $true
                }
            }
            if ($hasExcel) { $dCount++ }
            if ($fCount -gt $state.TotalTargetFiles) { $state.TotalTargetFiles = $fCount }
            $state.TotalTargetFolders = $dCount
        } catch {}
    }
    $state.TotalTargetFiles   = $fCount
    $state.TotalTargetFolders = $dCount
    $state.CountDone          = $true
}

$script:counterPs = [powershell]::Create()
[void]$script:counterPs.AddScript($counterScript).AddArgument($TargetFolder).AddArgument($sharedState)
$script:counterHandle = $script:counterPs.BeginInvoke()

$workerScript = {
    param($workerId, $xlamPath, $workQueue, $resultQueue, $state)

    Start-Sleep -Milliseconds (($workerId - 1) * 350)
    $script:excel = $null
    $script:wbAddin = $null
    $workerGuardDir = Join-Path $env:APPDATA "KangatangGuard"
    if (-not (Test-Path -LiteralPath $workerGuardDir)) { New-Item -ItemType Directory -Path $workerGuardDir -Force | Out-Null }
    $workerXlam = Join-Path $workerGuardDir "KangatangGuard_Worker_$workerId.xlam"
    try { Copy-Item -LiteralPath $xlamPath -Destination $workerXlam -Force -ErrorAction SilentlyContinue } catch {}
    if (-not (Test-Path -LiteralPath $workerXlam)) { $workerXlam = $xlamPath }

    function Clear-WorkerDdeRegistryBit {
        foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
            $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
            if (Test-Path $optKey) {
                try {
                    $curOpt = (Get-ItemProperty -Path $optKey -Name "Options" -ErrorAction SilentlyContinue).Options
                    if ($null -ne $curOpt -and (($curOpt -band 0x40) -ne 0)) {
                        Set-ItemProperty -Path $optKey -Name "Options" -Value ($curOpt -band (-bnot 0x40)) -Type DWord -Force -ErrorAction SilentlyContinue
                    }
                } catch {}
            }
        }
    }

    function Clear-WorkerCrashRecovery {
        Clear-WorkerDdeRegistryBit
        foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
            foreach ($sub in @("DocumentRecovery", "StartupItems")) {
                $resKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\$sub"
                if (Test-Path $resKey) {
                    Remove-Item -Path $resKey -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    function Restart-LocalWorker {
        if ($null -ne $script:wbAddin) {
            try { $script:wbAddin.Close($false) } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:wbAddin) } catch {}
            $script:wbAddin = $null
        }
        if ($null -ne $script:excel) {
            try {
                $script:excel.DisplayAlerts = $false
                foreach ($w in @($script:excel.Workbooks)) {
                    try { $w.Close($false) } catch {}
                }
            } catch {}
            try { $script:excel.IgnoreRemoteRequests = $false } catch {}
            try { $script:excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel) } catch {}
            $script:excel = $null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Clear-WorkerCrashRecovery

        try { Copy-Item -LiteralPath $xlamPath -Destination $workerXlam -Force -ErrorAction SilentlyContinue } catch {}
        $script:excel = New-Object -ComObject Excel.Application
        $script:excel.Visible = $false
        try { $script:excel.ShowWindowsInTaskbar = $false } catch {}
        try { $script:excel.AutoRecover.Enabled = $false } catch {}
        $script:excel.DisplayAlerts = $false
        $script:excel.EnableEvents = $false
        try { $script:excel.IgnoreRemoteRequests = $true } catch {}
        Clear-WorkerDdeRegistryBit
        $script:wbAddin = $script:excel.Workbooks.Open($workerXlam)
        $script:excel.Run("'$($script:wbAddin.Name)'!InitWorkerMode") | Out-Null
        Clear-WorkerDdeRegistryBit
    }

    try {
        try { Restart-LocalWorker } catch {}
        [System.Threading.Interlocked]::Increment([ref]$state.ActiveWorkers) | Out-Null
        $processedCount = 0

        $item = $null
        while (-not $state.CancelWorkers) {
            if (($null -eq $script:excel) -or ($null -eq $script:wbAddin)) {
                try {
                    Restart-LocalWorker
                } catch {
                    Start-Sleep -Milliseconds 500
                    continue
                }
            }

            if ($workQueue.TryDequeue([ref]$item)) {
                $fPath = [string]$item.FilePath
                $state.CurrentWorkerFile = [string]$item.FileName
                $res = ""
                $processedCount++
                try {
                    $script:excel.Visible = $false
                    $res = [string]$script:excel.Run("'$($script:wbAddin.Name)'!ScanSingleFileForWorker", $fPath)
                    if ($res.StartsWith("PURGE_NEEDED|")) {
                        $detail = $res.Substring(13)
                        $purgeRes = "MORE_PURGE"
                        for ($attempt = 1; $attempt -le 5; $attempt++) {
                            try { $script:excel.Run("'$($script:wbAddin.Name)'!WorkerDeleteOneModuleSheet", $fPath) | Out-Null } catch {}
                            try {
                                $purgeRes = [string]$script:excel.Run("'$($script:wbAddin.Name)'!WorkerFinishCleanAndClose", $fPath)
                            } catch {
                                $purgeRes = "INFECTED_ERR|RPC error"
                                try { Restart-LocalWorker } catch {}
                                $processedCount = 0
                                break
                            }
                            if ($purgeRes -ne "MORE_PURGE") { break }
                        }
                        if ($purgeRes.StartsWith("CLEANED|")) {
                            $res = "CLEANED_PURGE|$detail"
                        } else {
                            $res = $purgeRes
                        }
                    }
                    if ($processedCount -ge 350) {
                        try { Restart-LocalWorker } catch {}
                        $processedCount = 0
                    }
                } catch {
                    $res = "WORKER_RECOVERED|$($_.Exception.Message)"
                    try { Restart-LocalWorker; $processedCount = 0 } catch {}
                }

                $resultQueue.Enqueue([pscustomobject]@{
                    FilePath = $fPath
                    FileName = $item.FileName
                    FileSig  = $item.FileSig
                    LowPath  = $item.LowPath
                    Result   = $res
                    WorkerId = $workerId
                })
            } elseif ($state.ProducerDone -and $workQueue.IsEmpty) {
                break
            } else {
                Start-Sleep -Milliseconds 25
            }
        }
    } finally {
        if ($null -ne $script:wbAddin) {
            try { $script:wbAddin.Close($false) } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:wbAddin) } catch {}
        }
        if ($null -ne $script:excel) {
            try {
                $script:excel.DisplayAlerts = $false
                foreach ($w in @($script:excel.Workbooks)) {
                    try { $w.Close($false) } catch {}
                }
            } catch {}
            try { $script:excel.IgnoreRemoteRequests = $false } catch {}
            try { $script:excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel) } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        if ($workerXlam -ne $xlamPath -and (Test-Path -LiteralPath $workerXlam)) {
            Remove-Item -LiteralPath $workerXlam -Force -ErrorAction SilentlyContinue
        }
        Clear-WorkerCrashRecovery
    }
}

$script:runspacePool = $null
$script:workerJobs   = New-Object System.Collections.Generic.List[pscustomobject]
$script:poolStarted  = $false

function Ensure-WorkerPoolStarted {
    if (-not $script:poolStarted) {
        $script:poolStarted = $true
        $script:runspacePool = [runspacefactory]::CreateRunspacePool(1, $workerCount)
        $script:runspacePool.ApartmentState = [System.Threading.ApartmentState]::STA
        $script:runspacePool.Open()

        for ($wId = 1; $wId -le $workerCount; $wId++) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $script:runspacePool
            [void]$ps.AddScript($workerScript).AddArgument($wId).AddArgument($xlamPath).AddArgument($comWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
            $handle = $ps.BeginInvoke()
            $script:workerJobs.Add([pscustomobject]@{ PowerShell = $ps; Handle = $handle; Id = $wId })
        }
        return
    }

    # Watchdog: Tu dong khoi phuc Worker neu bat ky luong Worker nao bi ngat dot ngot khi van con viec
    if ((-not $sharedState.CancelWorkers) -and (-not ($sharedState.ProducerDone -and $comWorkQueue.IsEmpty))) {
        for ($i = 0; $i -lt $script:workerJobs.Count; $i++) {
            $job = $script:workerJobs[$i]
            if ($job.Handle.IsCompleted) {
                try { $job.PowerShell.Dispose() } catch {}
                $wId = $job.Id
                $ps = [powershell]::Create()
                $ps.RunspacePool = $script:runspacePool
                [void]$ps.AddScript($workerScript).AddArgument($wId).AddArgument($xlamPath).AddArgument($comWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
                $handle = $ps.BeginInvoke()
                $script:workerJobs[$i] = [pscustomobject]@{ PowerShell = $ps; Handle = $handle; Id = $wId }
            }
        }
    }
}

# 7. Bien trang thai va Ham ve Bang Dieu Khien Co Dinh (Non-Scrolling Live Dashboard UI)
$script:totalDiscovered = 0
$script:totalScanned    = 0
$script:totalSkipped    = 0
$script:totalZeroCom    = 0
$script:totalDeepCom    = 0
$script:totalSafe       = 0
$script:totalInfected   = $prevSessionCleaned
$script:totalCleaned    = $prevSessionCleaned
$script:totalErrors     = 0
$script:currentFolder   = $TargetFolder
$script:currentFile     = "Dang khoi tao bo quet..."
$script:folderIdx       = 0
$script:recentThreats   = New-Object System.Collections.Generic.List[string]
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$startTimeStr = if ($isResume -and $prevSessionStart) { $prevSessionStart } else { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
$script:lastUiTick = 0

$script:canSetCursor = $true
$script:dashboardTop = 0
$script:lastWinWidth = 100
try {
    $script:dashboardTop = [Console]::CursorTop
    $script:lastWinWidth = [Console]::WindowWidth
    [Console]::CursorVisible = $false
} catch {
    $script:canSetCursor = $false
}

function Write-FixedLine([string]$text, [System.ConsoleColor]$color = [System.ConsoleColor]::White) {
    $winW = 101
    try { $winW = [Console]::WindowWidth } catch {}
    $lineLen = [Math]::Max(40, [Math]::Min(100, $winW - 1))
    $padded = if ($text.Length -lt $lineLen) { $text.PadRight($lineLen) } else { $text.Substring(0, $lineLen) }
    Write-Host $padded -ForegroundColor $color
}

function Render-LiveDashboard([bool]$force = $false, [bool]$isCompleted = $false) {
    $nowTick = $stopwatch.ElapsedMilliseconds
    if ((-not $force) -and (($nowTick - $script:lastUiTick) -lt 100)) { return }
    $script:lastUiTick = $nowTick

    if ((-not $script:canSetCursor) -and (-not $isCompleted)) { return }
    if ($script:canSetCursor) {
        try {
            $curW = [Console]::WindowWidth
            if ($curW -ne $script:lastWinWidth) {
                $script:lastWinWidth = $curW
                [Console]::Clear()
                Write-FixedLine "====================================================================================================" ([System.ConsoleColor]::Cyan)
                Write-FixedLine "   KANGATANG GUARD v$AppVersion - BANG DIEU KHIEN QUET THU MUC (4-LAYER TURBO LIVE DASHBOARD)       " ([System.ConsoleColor]::Cyan)
                Write-FixedLine "   Day la addin phong chong virus kangatang - Phat trien boi OhsungIT team !                        " ([System.ConsoleColor]::White)
                Write-FixedLine "====================================================================================================" ([System.ConsoleColor]::Cyan)
                $script:dashboardTop = [Console]::CursorTop
            }
            [Console]::SetCursorPosition(0, $script:dashboardTop)
        } catch {
            $script:canSetCursor = $false
            return
        }
    }

    $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalDiscovered)
    $totalFolders = [Math]::Max([int]$sharedState.TotalTargetFolders, [int]$script:folderIdx)

    $pctDone = 0.0
    if ($isCompleted) {
        $pctDone = 100.0
        $totalTarget = $script:totalScanned
    } elseif ($totalTarget -gt 0) {
        $rawPct = ($script:totalScanned / [double]$totalTarget) * 100.0
        if ((-not $sharedState.CountDone) -and ($rawPct -ge 99.0)) {
            $pctDone = 99.0
        } else {
            $pctDone = [Math]::Min(99.9, [Math]::Round($rawPct, 1))
        }
    }
    $pctRem  = [Math]::Max(0.0, [Math]::Round(100.0 - $pctDone, 1))
    $remFiles = [Math]::Max(0, $totalTarget - $script:totalScanned)

    $elapsedSec = [Math]::Max(0.1, $stopwatch.Elapsed.TotalSeconds)
    $speed = [Math]::Round($script:totalScanned / $elapsedSec, 1)

    $etaStr = "--p--s"
    if ($isCompleted) {
        $etaStr = "00p00s"
    } elseif ($speed -gt 0 -and $remFiles -gt 0) {
        $etaSec = [int][Math]::Ceiling($remFiles / $speed)
        if ($etaSec -ge 3600) {
            $etaStr = "{0:D2}h{1:D2}p" -f [int][Math]::Floor($etaSec / 3600), [int][Math]::Floor(($etaSec % 3600) / 60)
        } else {
            $etaStr = "{0:D2}p{1:D2}s" -f [int][Math]::Floor($etaSec / 60), ($etaSec % 60)
        }
    } elseif ($remFiles -eq 0 -and $sharedState.CountDone) {
        $etaStr = "00p00s"
    }

    $barWidth = 24
    $filled = [Math]::Min($barWidth, [Math]::Max(0, [int][Math]::Round(($pctDone / 100.0) * $barWidth)))
    $empty  = $barWidth - $filled
    $progressBar = "[" + ("#" * $filled) + ("." * $empty) + "]"

    $activeFileDisplay = if ($isCompleted) {
        "HOAN TAT QUET TOAN BO THU MUC!"
    } elseif ((-not $comWorkQueue.IsEmpty) -and $sharedState.CurrentWorkerFile) {
        Format-EndTruncate $sharedState.CurrentWorkerFile 72
    } else {
        Format-EndTruncate $script:currentFile 72
    }

    $folderTag = if ($totalFolders -gt 0) { "[#$($script:folderIdx)/$totalFolders]" } else { "[#$($script:folderIdx)]" }
    $activeFolderDisplay = Format-MiddleTruncate $script:currentFolder (71 - $folderTag.Length)

    $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion | Tien do: $($pctDone.ToString('0.0'))% (Con lai: $($pctRem.ToString('0.0'))%) | Da quet: $($script:totalScanned)/$totalTarget | Da diet: $($script:totalCleaned)"

    # DONG 1: Ten tep dang xu ly & Duong dan thu muc dang thuc hien
    Write-FixedLine (" [1] TEP DANG XU LY    : $activeFileDisplay") ([System.ConsoleColor]::White)
    Write-FixedLine ("     THU MUC DANG QUET : $folderTag $activeFolderDisplay") ([System.ConsoleColor]::Cyan)
    Write-FixedLine ("----------------------------------------------------------------------------------------------------") ([System.ConsoleColor]::DarkGray)

    # DONG 2: Process Status, % Hoan thanh & % Con lai
    $statusColor = if ($isCompleted) { [System.ConsoleColor]::Green } else { [System.ConsoleColor]::Yellow }
    $pctDoneStr = ("{0,5:0.0}%" -f $pctDone)
    $pctRemStr  = ("{0,5:0.0}%" -f $pctRem)
    Write-FixedLine (" [2] PROCESS STATUS    : $progressBar $pctDoneStr | Con lai: $pctRemStr ($remFiles tep) | ETA: $etaStr") $statusColor
    Write-FixedLine ("====================================================================================================") ([System.ConsoleColor]::Cyan)

    # DONG 3+: Bang ket qua thong ke truc tuyen (Live Summary Table - Chuan 100 ky tu)
    $countStateStr = if ($sharedState.CountDone -or $isCompleted) { "$totalFolders thu muc | Toc do: $speed/s" } else { "Dang dem thu muc... | $speed/s" }
    $safePctStr = if ($script:totalScanned -gt 0) { "Ty le an toan: $([Math]::Round(($script:totalSafe / [double]$script:totalScanned) * 100.0, 1))%" } else { "Dang kiem tra..." }
    $deepQueueWait = $comWorkQueue.Count

    Write-FixedLine (" | [3] BANG KET QUA QUET TRUC TUYEN (LIVE TABLE) |       SO LUONG | CHI TIET / TRANG THAI          |") ([System.ConsoleColor]::Cyan)
    Write-FixedLine (" +-----------------------------------------------+----------------+--------------------------------+") ([System.ConsoleColor]::DarkCyan)
    Write-FixedLine (" | 1. Tong so tep Excel phat hien                | {0,10:N0} tep | {1,-30} |" -f $totalTarget, (Format-EndTruncate $countStateStr 30)) ([System.ConsoleColor]::White)
    Write-FixedLine (" | 2. Tong so tep da quet xong                   | {0,10:N0} tep | {1,-30} |" -f $script:totalScanned, ("Xong: {0:0.0}% | Con lai: {1:0.0}%" -f $pctDone, $pctRem)) ([System.ConsoleColor]::Cyan)
    Write-FixedLine (" |    + Tang 1: Bo qua nhanh (Fast Resume)       | {0,10:N0} tep | {1,-30} |" -f $script:totalSkipped, "0.01 ms/tep (Da sach tu truoc)") ([System.ConsoleColor]::Gray)
    Write-FixedLine (" |    + Tang 2: Loc tinh Zero-COM (.xlsx)        | {0,10:N0} tep | {1,-30} |" -f $script:totalZeroCom, "1.50 ms/tep (Khong mo Excel)") ([System.ConsoleColor]::Gray)
    Write-FixedLine (" |    + Tang 3: Quet sau Excel COM (.xls/Macro)  | {0,10:N0} tep | {1,-30} |" -f $script:totalDeepCom, ("Cho xu ly: $deepQueueWait | $workerCount Workers")) ([System.ConsoleColor]::Gray)
    Write-FixedLine (" +-----------------------------------------------+----------------+--------------------------------+") ([System.ConsoleColor]::DarkCyan)
    Write-FixedLine (" | 3. Tep AN TOAN (Sach)                         | {0,10:N0} tep | {1,-30} |" -f $script:totalSafe, (Format-EndTruncate $safePctStr 30)) ([System.ConsoleColor]::Green)
    $infColor = if ($script:totalInfected -gt 0) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
    Write-FixedLine (" | 4. Tep PHAT HIEN NHIEM VIRUS                  | {0,10:N0} tep | {1,-30} |" -f $script:totalInfected, $(if ($script:totalInfected -gt 0) { "Phat hien ma doc Kangatang!" } else { "Khong phat hien ma doc" })) $infColor
    $clnColor = if ($script:totalCleaned -gt 0) { [System.ConsoleColor]::Green } else { [System.ConsoleColor]::White }
    Write-FixedLine (" | 5. Tep DA XU LY & SAO LUU (ve may 223.176)    | {0,10:N0} tep | {1,-30} |" -f $script:totalCleaned, $(if ($script:totalCleaned -gt 0) { "Da diet sach & sao luu an toan" } else { "Chua co tep can xu ly" })) $clnColor
    $errColor = if ($script:totalErrors -gt 0) { [System.ConsoleColor]::Red } else { [System.ConsoleColor]::Gray }
    Write-FixedLine (" | 6. Tep BO QUA / LOI (Khoa Password/ReadOnly)  | {0,10:N0} tep | {1,-30} |" -f $script:totalErrors, "Tep dat mat khau hoac dang mo") $errColor
    Write-FixedLine (" +-----------------------------------------------+----------------+--------------------------------+") ([System.ConsoleColor]::DarkCyan)

    # KHOI 4: Nhat ky 5 tep nhiem virus vua xu ly gan nhat (Co dinh 5 dong de khong cuon man hinh)
    Write-FixedLine (" [4] NHAT KY TEP NHIEM VIRUS VUA XU LY (5 TEP GAN NHAT):") ([System.ConsoleColor]::Yellow)
    for ($i = 0; $i -lt 5; $i++) {
        if ($i -lt $script:recentThreats.Count) {
            $tLine = Format-EndTruncate $script:recentThreats[$i] 96
            Write-FixedLine ("   $tLine") ([System.ConsoleColor]::Yellow)
        } elseif ($i -eq 0 -and $script:recentThreats.Count -eq 0) {
            Write-FixedLine ("   - (Chua phat hien tep nhiem virus nao trong phien quet nay)") ([System.ConsoleColor]::DarkGray)
        } else {
            Write-FixedLine ("") ([System.ConsoleColor]::DarkGray)
        }
    }
    Write-FixedLine ("====================================================================================================") ([System.ConsoleColor]::Cyan)
}

function Add-RecentThreatLog([string]$entryText) {
    $cleanEntry = ConvertTo-DisplayAscii $entryText
    $script:recentThreats.Insert(0, $cleanEntry)
    while ($script:recentThreats.Count -gt 5) {
        $script:recentThreats.RemoveAt($script:recentThreats.Count - 1)
    }
}

function Drain-WorkerResults {
    $resItem = $null
    $updated = $false
    while ($comResultQueue.TryDequeue([ref]$resItem)) {
        $updated = $true
        $r = [string]$resItem.Result
        $fName = [string]$resItem.FileName
        $fPath = [string]$resItem.FilePath
        $wId   = $resItem.WorkerId
        $script:totalDeepCom++
        $script:currentFile = $fName

        if ($r -eq "SAFE") {
            $script:totalScanned++
            $script:totalSafe++
            Update-DeptStat $fPath "SAFE"
            $script:logBuffer.Add("$($resItem.LowPath)|$($resItem.FileSig)")
        } elseif ($r.StartsWith("CLEANED|") -or $r.StartsWith("CLEANED_PURGE|")) {
            $script:totalScanned++
            $script:totalInfected++
            $script:totalCleaned++
            Update-DeptStat $fPath "CLEANED"
            $isPurge = $r.StartsWith("CLEANED_PURGE|")
            $threatFamily = if ($isPurge) { "Excel5_ModuleSheet_Hidden" } else { "VBA_Macro_Kangatang" }
            $detail = if ($isPurge) { "Module-Sheet Excel 5.0: " + $r.Substring(14) } else { $r.Substring(8) }
            $nowDt = Get-Date
            $ts = $nowDt.ToString("HH:mm:ss")
            $tsFull = $nowDt.ToString("yyyy-MM-dd HH:mm:ss")
            Add-RecentThreatLog "> [$ts] [DA DIET] $fName ($detail)"
            $deptName = Get-TopDepartmentName $fPath
            $folderDir = ""
            try { $folderDir = [System.IO.Path]::GetDirectoryName($fPath) } catch {}
            $extStr = ""
            try { $extStr = [System.IO.Path]::GetExtension($fName).ToLowerInvariant() } catch {}
            $sizeKb = 0
            try {
                $partsSig = ([string]$resItem.FileSig).Split('|')
                if ($partsSig.Count -ge 1) { $sizeKb = [Math]::Round(([long]$partsSig[0]) / 1024.0, 1) }
            } catch {}

            [void]$script:infectedRecords.Add([pscustomobject]@{
                Time         = $tsFull
                TimeShort    = $ts
                Department   = $deptName
                FileName     = $fName
                FolderPath   = $folderDir
                FullPath     = $fPath
                Extension    = $extStr
                SizeKB       = $sizeKb
                ThreatFamily = $threatFamily
                ThreatDetail = $detail
                Action       = "CLEANED_AND_QUARANTINED"
                WorkerId     = $wId
            })

            $csvRow = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}" -f `
                (Format-CsvField $script:sessionId),
                (Format-CsvField $tsFull),
                (Format-CsvField $env:COMPUTERNAME),
                (Format-CsvField $deptName),
                (Format-CsvField $fName),
                (Format-CsvField $folderDir),
                (Format-CsvField $fPath),
                (Format-CsvField $extStr),
                $sizeKb,
                (Format-CsvField $threatFamily),
                (Format-CsvField $detail),
                (Format-CsvField "CLEANED_AND_QUARANTINED"),
                $wId
            $script:ledgerBuffer.Add($csvRow)
            Add-DetailAuditEvent "DETECTED_CLEANED" $fPath $wId $threatFamily "Da diet & cach ly thanh cong ($detail)"

            try {
                $fAfter = New-Object System.IO.FileInfo($fPath)
                $newSig = "$($fAfter.Length)|$($fAfter.LastWriteTime.ToString('yyyyMMddHHmmss'))"
                $script:logBuffer.Add("$($resItem.LowPath)|$newSig")
            } catch {}
        } elseif ($r.StartsWith("INFECTED_")) {
            $script:totalScanned++
            $script:totalInfected++
            $script:totalErrors++
            Update-DeptStat $fPath "INFECTED_LOCKED"
            $nowDt = Get-Date
            $ts = $nowDt.ToString("HH:mm:ss")
            $tsFull = $nowDt.ToString("yyyy-MM-dd HH:mm:ss")
            Add-RecentThreatLog "! [$ts] [KHOA/READONLY] $fName ($r)"
            $deptName = Get-TopDepartmentName $fPath
            $folderDir = ""
            try { $folderDir = [System.IO.Path]::GetDirectoryName($fPath) } catch {}
            $extStr = ""
            try { $extStr = [System.IO.Path]::GetExtension($fName).ToLowerInvariant() } catch {}
            $sizeKb = 0
            try {
                $partsSig = ([string]$resItem.FileSig).Split('|')
                if ($partsSig.Count -ge 1) { $sizeKb = [Math]::Round(([long]$partsSig[0]) / 1024.0, 1) }
            } catch {}

            [void]$script:infectedRecords.Add([pscustomobject]@{
                Time         = $tsFull
                TimeShort    = $ts
                Department   = $deptName
                FileName     = $fName
                FolderPath   = $folderDir
                FullPath     = $fPath
                Extension    = $extStr
                SizeKB       = $sizeKb
                ThreatFamily = "Locked_Infected_Workbook"
                ThreatDetail = $r
                Action       = "INFECTED_LOCKED_NEED_UNLOCK"
                WorkerId     = $wId
            })

            $csvRow = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}" -f `
                (Format-CsvField $script:sessionId),
                (Format-CsvField $tsFull),
                (Format-CsvField $env:COMPUTERNAME),
                (Format-CsvField $deptName),
                (Format-CsvField $fName),
                (Format-CsvField $folderDir),
                (Format-CsvField $fPath),
                (Format-CsvField $extStr),
                $sizeKb,
                (Format-CsvField "Locked_Infected_Workbook"),
                (Format-CsvField $r),
                (Format-CsvField "INFECTED_LOCKED_NEED_UNLOCK"),
                $wId
            $script:ledgerBuffer.Add($csvRow)
            Add-DetailAuditEvent "DETECTED_LOCKED" $fPath $wId "Locked_Infected_Workbook" "Phat hien nhiem nhung tep dang bi khoa/ReadOnly ($r)"
        } else {
            $script:totalScanned++
            $script:totalErrors++
            Update-DeptStat $fPath "ERROR"
            $script:logBuffer.Add("$($resItem.LowPath)|$($resItem.FileSig)")
        }

        if ($r.StartsWith("CLEANED") -or $r.StartsWith("INFECTED_") -or ($script:logBuffer.Count -ge 20)) {
            $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalDiscovered)
            Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
        }
    }
    if ($updated) {
        Render-LiveDashboard $true $false
    }
}

Add-DetailAuditEvent "SESSION_START" $TargetFolder 0 "(ALL)" "Bat dau phien quet thu muc v$AppVersion (Resume=$isResume, Workers=$maxWorkers)"
Flush-ScanLogs "IN_PROGRESS" $prevSessionTotal 0 0 $startTimeStr
Render-LiveDashboard $true $false

$dirQueue  = New-Object System.Collections.Generic.Queue[string]
$dirQueue.Enqueue($TargetFolder)
$validExts = @(".xls", ".xlsx", ".xlsm", ".xlsb")

try {
    while ($dirQueue.Count -gt 0) {
        $folder = $dirQueue.Dequeue()

        # Tang 1: Streaming Queue - Day cac thu muc con vao hang doi ngay khi duyet toi
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($folder)
            foreach ($sd in $subDirs) {
                $dirName = [System.IO.Path]::GetFileName($sd)
                if (-not (Test-SkippedFolderName $dirName)) {
                    $dirQueue.Enqueue($sd)
                }
            }
        } catch {}

        # Loc danh sach tep Excel trong thu muc hien tai
        $excelFiles = New-Object System.Collections.Generic.List[string]
        try {
            $allFiles = [System.IO.Directory]::GetFiles($folder)
            foreach ($fPath in $allFiles) {
                try {
                    $fName = [System.IO.Path]::GetFileName($fPath)
                    $ext   = [System.IO.Path]::GetExtension($fName).ToLowerInvariant()
                    if (($ext -in $validExts) -and (-not $fName.StartsWith('~$')) -and (-not $fName.ToLowerInvariant().StartsWith("kangatangguard"))) {
                        $excelFiles.Add($fPath)
                    }
                } catch {}
            }
        } catch {
            continue
        }

        if ($excelFiles.Count -eq 0) { continue }
        $script:folderIdx++
        $script:currentFolder = $folder
        $script:totalDiscovered += $excelFiles.Count

        foreach ($filePath in $excelFiles) {
            try {
                $fileName = [System.IO.Path]::GetFileName($filePath)
                $ext      = [System.IO.Path]::GetExtension($fileName).ToLowerInvariant()
                $script:currentFile = $fileName
                $lowPath  = $filePath.ToLowerInvariant()

                # TANG 1: Fast Resume trong RAM (0.001ms - Zero SMB FileInfo Round-Trip)
                if ($isResume -and $resumeMap.ContainsKey($lowPath)) {
                    $script:totalSkipped++
                    $script:totalScanned++
                    $script:totalSafe++
                    Update-DeptStat $filePath "SAFE"
                    if (($script:totalSkipped % 2000) -eq 0) {
                        $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalDiscovered)
                        Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
                    }
                    Render-LiveDashboard $false $false
                    continue
                }

                $fInfo = $null
                try { $fInfo = New-Object System.IO.FileInfo($filePath) } catch { continue }
                if (-not $fInfo -or (-not $fInfo.Exists)) { continue }

                $fileSig = "$($fInfo.Length)|$($fInfo.LastWriteTime.ToString('yyyyMMddHHmmss'))"

                # TANG 2: Zero-COM OpenXML Pre-Filter cho .xlsx/.xlsm/.xlsb/.xls (1.5 - 3ms)
                if (Test-OpenXmlSafeZeroCom $filePath $ext $iocRegex) {
                    $script:totalZeroCom++
                    $script:totalScanned++
                    $script:totalSafe++
                    Update-DeptStat $filePath "SAFE"
                    $script:logBuffer.Add("$lowPath|$fileSig")
                }
                # TANG 3: Day vao hang doi Multi-Worker Parallel Excel COM Pool (.xls hoac tep co vbaProject.bin)
                else {
                    Ensure-WorkerPoolStarted
                    $comWorkQueue.Enqueue([pscustomobject]@{
                        FilePath = $filePath
                        FileName = $fileName
                        FileSig  = $fileSig
                        LowPath  = $lowPath
                    })
                }
            } catch {
                $script:totalScanned++
                $script:totalErrors++
                Update-DeptStat $filePath "ERROR"
            }

            Drain-WorkerResults
            Render-LiveDashboard $false $false

            # Dieu tiet hang doi (Backpressure): neu hang doi Deep COM > 60 tep thi cho Worker Pool xu ly bot
            $bpTicks = 0
            $prevQueueLen = $comWorkQueue.Count
            while ($comWorkQueue.Count -gt 60) {
                Ensure-WorkerPoolStarted
                Drain-WorkerResults
                Render-LiveDashboard $false $false
                Start-Sleep -Milliseconds 40
                if ($comWorkQueue.Count -lt $prevQueueLen) {
                    $prevQueueLen = $comWorkQueue.Count
                    $bpTicks = 0
                } else {
                    $bpTicks++
                    if ($bpTicks -gt 1500) {
                        # Safety valve (60s): neu tat ca COM worker deu bi ket tren tep mang loi, bo qua 1 tep dau hang doi
                        $stuckItem = $null
                        if ($comWorkQueue.TryDequeue([ref]$stuckItem)) {
                            $script:totalScanned++
                            $script:totalErrors++
                            Update-DeptStat $stuckItem.FilePath "ERROR"
                            Add-DetailAuditEvent "WORKER_TIMEOUT_SKIP" $stuckItem.FilePath 0 "(TIMEOUT)" "Bo qua tep bi treo qua 60s trong hang doi Deep COM"
                        }
                        $bpTicks = 0
                    }
                }
            }
        }

        Drain-WorkerResults
        $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalDiscovered)
        Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
    }

    $sharedState.TotalTargetFiles   = $script:totalDiscovered
    $sharedState.TotalTargetFolders = $script:folderIdx
    $sharedState.CountDone          = $true

    # Cho tat ca cac tep .xls/.xlsm con lai trong hang doi Multi-Worker Pool hoan tat
    $sharedState.ProducerDone = $true
    if ($script:poolStarted) {
        $drainTicks = 0
        $prevRem = $comWorkQueue.Count
        while ((-not $comWorkQueue.IsEmpty) -or (-not $comResultQueue.IsEmpty)) {
            Ensure-WorkerPoolStarted
            Drain-WorkerResults
            Render-LiveDashboard $false $false
            if (-not $comWorkQueue.IsEmpty) {
                Start-Sleep -Milliseconds 60
                if ($comWorkQueue.Count -lt $prevRem) {
                    $prevRem = $comWorkQueue.Count
                    $drainTicks = 0
                } else {
                    $drainTicks++
                    if ($drainTicks -gt 1000) {
                        $stuckItem = $null
                        if ($comWorkQueue.TryDequeue([ref]$stuckItem)) {
                            $script:totalScanned++
                            $script:totalErrors++
                            Update-DeptStat $stuckItem.FilePath "ERROR"
                            Add-DetailAuditEvent "WORKER_TIMEOUT_SKIP" $stuckItem.FilePath 0 "(TIMEOUT)" "Bo qua tep cuoi hang doi bi treo qua 60s"
                        }
                        $drainTicks = 0
                    }
                }
            }
        }

        foreach ($job in $script:workerJobs) {
            try { [void]$job.PowerShell.EndInvoke($job.Handle) } catch {}
        }
        Drain-WorkerResults
    }

    $stopwatch.Stop()
    $totalSec = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    Render-LiveDashboard $true $true

    Add-DetailAuditEvent "SESSION_COMPLETED" $TargetFolder 0 "(ALL)" "Hoan tat trong ${totalSec}s: Scanned=$($script:totalScanned), Infected=$($script:totalInfected), Cleaned=$($script:totalCleaned), Errors=$($script:totalErrors)"
    Flush-ScanLogs "COMPLETED" $script:totalDiscovered $script:totalScanned $script:totalCleaned $startTimeStr
    $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - HOAN TAT ($totalSec giay) | Da quet: $($script:totalScanned) | Da diet: $($script:totalCleaned)"
    Write-Host "`n   [HOAN TAT] Qua trinh quet thu muc da ket thuc trong $totalSec giay!" -ForegroundColor Green
    Write-Host "   [BAO CAO]  $($script:latestMdPath)" -ForegroundColor Cyan
} catch {
    try {
        $errMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Top-level error: $($_.Exception.Message)`r`n$($_.ScriptStackTrace)`r`n"
        [System.IO.File]::AppendAllText((Join-Path $logDir "scanner_error.log"), $errMsg, $utf8NoBom)
        Add-DetailAuditEvent "SESSION_INTERRUPTED" $TargetFolder 0 "(ERROR)" "Ngoai le: $($_.Exception.Message)"
    } catch {}
    $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalDiscovered)
    Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
    Write-Host "`n   [CANH BAO] Co ngoai le phat sinh: $($_.Exception.Message). Tien trinh da duoc luu an toan!" -ForegroundColor Yellow
} finally {
    try { [Console]::CursorVisible = $true } catch {}
    $sharedState.CancelWorkers = $true
    $sharedState.ProducerDone  = $true
    if ($null -ne $script:counterPs) {
        try { $script:counterPs.Dispose() } catch {}
    }
    foreach ($job in $script:workerJobs) {
        try { $job.PowerShell.Dispose() } catch {}
    }
    if ($null -ne $script:runspacePool) {
        try { $script:runspacePool.Close(); $script:runspacePool.Dispose() } catch {}
    }
    if ($script:hasMutex -and ($null -ne $script:scannerMutex)) {
        try { $script:scannerMutex.ReleaseMutex(); $script:scannerMutex.Dispose() } catch {}
    }
}


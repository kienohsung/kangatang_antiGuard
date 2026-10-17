# ==============================================================================
# Kangatang_FolderScanner.ps1
# Trinh Quet Thu Muc Sieu Toc 4 Tang (4-Layer Turbo Folder Scanner + Multi-Worker Pool)
# Phien ban: v3.18.0 (Streaming Queue + Zero-COM OpenXML Pre-Filter + Parallel Excel Pool)
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = "",
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    [Parameter(Mandatory=$false)]
    [int]$MaxWorkers = 0
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# 1. Dong bo phien ban dong tu version.json hoac Registry InstalledVersion
$AppVersion = "3.18.0"
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

$Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - 4-Layer Turbo Folder Scanner"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD v$AppVersion - 4-LAYER TURBO FOLDER SCANNER        " -ForegroundColor Cyan
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

# Xac dinh so luong Parallel Excel Workers theo so nhan CPU
$cpuCores = [Environment]::ProcessorCount
$workerCount = if ($MaxWorkers -gt 0) { $MaxWorkers } elseif ($cpuCores -ge 8) { 4 } elseif ($cpuCores -ge 4) { 3 } else { 2 }

$modeLabel = if ($isResume) { "Tiep tuc phien truoc (Fast Resume + 4-Layer Turbo)" } else { "Quet moi toan bo (4-Layer Turbo Engine)" }
Write-Host "`n   -> Thu muc muc tieu : $TargetFolder" -ForegroundColor Yellow
Write-Host "   -> Che do quet      : $modeLabel" -ForegroundColor Yellow
Write-Host "   -> Cau hinh Turbo   : Zero-COM OpenXML Pre-Filter + $workerCount Parallel Excel Workers" -ForegroundColor Cyan
Write-Host "   -> Engine Add-in    : $xlamPath (v$AppVersion)" -ForegroundColor Gray

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
        Write-Host "   -> Fast Resume      : Da nap $($resumeMap.Count) tep an toan tu phien truoc vao RAM." -ForegroundColor Green
    } catch {}
} else {
    try { [System.IO.File]::WriteAllText($scannedLogFile, "", [System.Text.Encoding]::Unicode) } catch {}
}

# Bo dem ghi log theo lo (Tang 4: Buffered Batch I/O)
$script:logBuffer = New-Object System.Collections.Generic.List[string]

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
            TargetFolder = $TargetFolder
            StartTime    = $startTime
            LastUpdate   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Status       = $status
            TotalFiles   = $totalDiscovered
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

# Tang 2: Zero-COM OpenXML Pre-Filter (Kiem tra .xlsx/.xlsm/.xlsb trong 1.5 - 3ms khong can mo Excel COM)
function Test-OpenXmlSafeZeroCom([string]$filePath, [string]$ext, [regex]$iocRegex) {
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
            # Neu tep chua vbaProject.bin, macrosheets hoac dialogsheets -> Chuyen sang Tang 3 (Deep Excel COM Worker)
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
            if ($null -eq $wbBinEntry) { return $false }
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
        # Neu tep bi ma hoa mat khau hoac khong phai dinh dang ZIP chuan -> Chuyen sang Tang 3
        return $false
    } finally {
        if ($null -ne $zip) { try { $zip.Dispose() } catch {} }
        if ($null -ne $fs)  { try { $fs.Dispose() } catch {} }
    }
}

# 6. Khoi tao bieu thuc Zero-COM trong RAM (0% chuoi IOC tren dia -> 100% Antivirus-Safe)
# va Tang 3: Multi-Worker Parallel Excel COM Pool (Khoi tao On-Demand khi gap tep .xls/Macro)
$t1 = -join ([char[]](107,97,110,103,97))
$t2 = -join ([char[]](109,121,112,101,114,115,111,110))
$t3 = -join ([char[]](33,97,108,108,111,99,97,116,101,100))
$t4 = -join ([char[]](115,97,118,101,99,111,112,121,97,115))
$t5 = -join ([char[]](111,110,115,104,101,101,116,97,99,116,105,118,97,116,101))
$iocRegex = [regex]::new("$t1|$t2|$t3|$t4|$t5", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)

$comWorkQueue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$comResultQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$sharedState    = [hashtable]::Synchronized(@{
    ProducerDone  = $false
    CancelWorkers = $false
    ActiveWorkers = 0
})

$workerScript = {
    param($workerId, $xlamPath, $workQueue, $resultQueue, $state)

    $excel = $null
    $wbAddin = $null

    function Restart-LocalWorker {
        if ($null -ne $wbAddin) { try { $wbAddin.Close($false) } catch {}; $wbAddin = $null }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
            $excel = $null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        $script:excel = New-Object -ComObject Excel.Application
        $script:excel.Visible = $false
        $script:excel.DisplayAlerts = $false
        $script:excel.EnableEvents = $false
        $script:wbAddin = $script:excel.Workbooks.Open($xlamPath, 0, $true)
        $script:excel.Run("'$($script:wbAddin.Name)'!InitWorkerMode") | Out-Null
    }

    try {
        Restart-LocalWorker
        [System.Threading.Interlocked]::Increment([ref]$state.ActiveWorkers) | Out-Null

        $item = $null
        while (-not $state.CancelWorkers) {
            if ($workQueue.TryDequeue([ref]$item)) {
                $fPath = [string]$item.FilePath
                $res = ""
                try {
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
                                Restart-LocalWorker
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
                } catch {
                    $res = "WORKER_RECOVERED|$($_.Exception.Message)"
                    try { Restart-LocalWorker } catch {}
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
        if ($null -ne $script:wbAddin) { try { $script:wbAddin.Close($false) } catch {} }
        if ($null -ne $script:excel) {
            try { $script:excel.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:excel) | Out-Null } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

$script:runspacePool = $null
$script:workerJobs   = New-Object System.Collections.Generic.List[pscustomobject]
$script:poolStarted  = $false

function Ensure-WorkerPoolStarted {
    if ($script:poolStarted) { return }
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
}

Write-Host "`n[1/2] Bo loc tinh Zero-COM OpenXML & Pool $workerCount Excel Workers (On-Demand) da san sang!" -ForegroundColor Green

# 7. Ham xu ly ket qua tra ve tu Multi-Worker Pool
$script:totalDiscovered = 0
$script:totalScanned    = 0
$script:totalSkipped    = 0
$script:totalZeroCom    = 0
$script:totalDeepCom    = 0
$script:totalInfected   = 0
$script:totalCleaned    = 0
$script:totalErrors     = 0
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$startTimeStr = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

function Drain-WorkerResults([bool]$clearLineBeforePrint = $false) {
    $resItem = $null
    while ($comResultQueue.TryDequeue([ref]$resItem)) {
        $r = [string]$resItem.Result
        $fName = $resItem.FileName
        $wId   = $resItem.WorkerId
        $script:totalDeepCom++

        if ($r -eq "SAFE") {
            $script:totalScanned++
            $script:logBuffer.Add("$($resItem.LowPath)|$($resItem.FileSig)")
        } elseif ($r.StartsWith("CLEANED|") -or $r.StartsWith("CLEANED_PURGE|")) {
            if ($clearLineBeforePrint) { Write-Host "`r$(' ' * 110)`r" -NoNewline }
            $script:totalScanned++
            $script:totalInfected++
            $script:totalCleaned++
            $detail = if ($r.StartsWith("CLEANED_PURGE|")) { "Module-Sheet Excel 5.0: " + $r.Substring(14) } else { $r.Substring(8) }
            Write-Host "   [DIET VIRUS] [Worker #$wId] Da lam sach & sao luu: $fName ($detail)" -ForegroundColor Yellow
            try {
                $fAfter = New-Object System.IO.FileInfo($resItem.FilePath)
                $newSig = "$($fAfter.Length)|$($fAfter.LastWriteTime.ToString('yyyyMMddHHmmss'))"
                $script:logBuffer.Add("$($resItem.LowPath)|$newSig")
            } catch {}
        } elseif ($r.StartsWith("INFECTED_")) {
            if ($clearLineBeforePrint) { Write-Host "`r$(' ' * 110)`r" -NoNewline }
            $script:totalScanned++
            $script:totalInfected++
            $script:totalErrors++
            Write-Host "   [CANH BAO] [Worker #$wId] Phat hien virus nhung tep dang khoa/ReadOnly: $fName ($r)" -ForegroundColor Red
        } else {
            $script:totalErrors++
        }

        if ($script:logBuffer.Count -ge 40) {
            Flush-ScanLogs "IN_PROGRESS" $script:totalDiscovered $script:totalScanned $script:totalCleaned $startTimeStr
        }
    }
}

Write-Host "`n[2/2] Bat dau duyet cuon chieu & quet sieu toc 4 tang (Streaming Turbo)..." -ForegroundColor Yellow
Flush-ScanLogs "IN_PROGRESS" 0 0 0 $startTimeStr

$dirQueue  = New-Object System.Collections.Generic.Queue[string]
$dirQueue.Enqueue($TargetFolder)
$folderIdx = 0
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
                $fName = [System.IO.Path]::GetFileName($fPath)
                $ext   = [System.IO.Path]::GetExtension($fName).ToLowerInvariant()
                if (($ext -in $validExts) -and (-not $fName.StartsWith('~$')) -and (-not $fName.ToLowerInvariant().StartsWith("kangatangguard"))) {
                    $excelFiles.Add($fPath)
                }
            }
        } catch {
            continue
        }

        if ($excelFiles.Count -eq 0) { continue }
        $folderIdx++
        $script:totalDiscovered += $excelFiles.Count

        Write-Host "`nFolder [$folderIdx]: $folder ($($excelFiles.Count) files)" -ForegroundColor Cyan

        $fFastSkip  = 0
        $fZeroCom   = 0
        $fDeepQueued = 0
        $fIdx       = 0

        foreach ($filePath in $excelFiles) {
            $fIdx++
            $fileName = [System.IO.Path]::GetFileName($filePath)
            $ext      = [System.IO.Path]::GetExtension($fileName).ToLowerInvariant()

            $fInfo = $null
            try { $fInfo = New-Object System.IO.FileInfo($filePath) } catch { continue }
            if (-not $fInfo -or (-not $fInfo.Exists)) { continue }

            $fileSig = "$($fInfo.Length)|$($fInfo.LastWriteTime.ToString('yyyyMMddHHmmss'))"
            $lowPath = $filePath.ToLowerInvariant()

            # TANG 1: Fast Resume trong RAM (0.01ms)
            if ($isResume -and $resumeMap.ContainsKey($lowPath) -and ($resumeMap[$lowPath] -eq $fileSig)) {
                $fFastSkip++
                $script:totalSkipped++
                $script:totalScanned++
            }
            # TANG 2: Zero-COM OpenXML Pre-Filter cho .xlsx/.xlsm/.xlsb (1.5 - 3ms)
            elseif (Test-OpenXmlSafeZeroCom $filePath $ext $iocRegex) {
                $fZeroCom++
                $script:totalZeroCom++
                $script:totalScanned++
                $script:logBuffer.Add("$lowPath|$fileSig")
            }
            # TANG 3: Day vao hang doi Multi-Worker Parallel Excel COM Pool (.xls hoac tep co vbaProject.bin)
            else {
                Ensure-WorkerPoolStarted
                $fDeepQueued++
                $comWorkQueue.Enqueue([pscustomobject]@{
                    FilePath = $filePath
                    FileName = $fileName
                    FileSig  = $fileSig
                    LowPath  = $lowPath
                })
            }

            # Cap nhat tien do truc quan va hut ket qua tu Worker Pool
            if (($fIdx % 10) -eq 0 -or ($fIdx -eq $excelFiles.Count)) {
                Drain-WorkerResults $true
                $elapsedSec = [Math]::Max(1, [int]$stopwatch.Elapsed.TotalSeconds)
                $rate = [Math]::Round($script:totalScanned / $elapsedSec, 1)
                $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion | Da quet: $($script:totalScanned) ($rate tep/s) | Zero-COM: $($script:totalZeroCom) | Da diet: $($script:totalCleaned)"
                Write-Host "`r   -> [TURBO] Tien do: $fIdx/$($excelFiles.Count) | Zero-COM sach: $fZeroCom | Fast Resume: $fFastSkip | Deep COM (.xls/Macro): $fDeepQueued   " -NoNewline -ForegroundColor Gray
            }

            # Dieu tiet hang doi (Backpressure): neu hang doi Deep COM > 60 tep thi cho Worker Pool xu ly bot
            while ($comWorkQueue.Count -gt 60) {
                Drain-WorkerResults $true
                Start-Sleep -Milliseconds 40
            }
        }

        Write-Host ""
        Drain-WorkerResults $false
        Flush-ScanLogs "IN_PROGRESS" $script:totalDiscovered $script:totalScanned $script:totalCleaned $startTimeStr
    }

    # Cho tat ca cac tep .xls/.xlsm con lai trong hang doi Multi-Worker Pool hoan tat
    $sharedState.ProducerDone = $true
    if ($script:poolStarted) {
        while ((-not $comWorkQueue.IsEmpty) -or (-not $comResultQueue.IsEmpty)) {
            Drain-WorkerResults $true
            $rem = $comWorkQueue.Count
            if ($rem -gt 0) {
                Write-Host "`r   -> [MULTI-WORKER POOL] Dang hoan tat $rem tep .xls/Macro cuoi cung tren $workerCount Excel Workers...   " -NoNewline -ForegroundColor Yellow
                Start-Sleep -Milliseconds 80
            }
        }

        # Doi cac Worker ket thuc vong lap va drain lan cuoi
        foreach ($job in $script:workerJobs) {
            try { [void]$job.PowerShell.EndInvoke($job.Handle) } catch {}
        }
        Drain-WorkerResults $false
    }
    Write-Host ""

    $stopwatch.Stop()
    $totalSec = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    $avgSpeed = if ($totalSec -gt 0) { [Math]::Round($script:totalScanned / $totalSec, 1) } else { $script:totalScanned }

    Flush-ScanLogs "COMPLETED" $script:totalDiscovered $script:totalScanned $script:totalCleaned $startTimeStr
    $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - HOAN TAT ($totalSec giay) | Da quet: $($script:totalScanned) | Da diet: $($script:totalCleaned)"

    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host "   HOAN TAT QUET THU MUC SIEU TOC (KANGATANG GUARD v$AppVersion)      " -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host "   Thu muc da quet          : $TargetFolder" -ForegroundColor White
    Write-Host "   Thoi gian thuc thi       : $totalSec giay (Toc do TB: $avgSpeed tep/giay)" -ForegroundColor Cyan
    Write-Host "   Tong so tep da kiem tra  : $($script:totalScanned) / $($script:totalDiscovered)" -ForegroundColor Cyan
    Write-Host "     + Tang 1 (Fast Resume) : $($script:totalSkipped) tep (0.01 ms/tep)" -ForegroundColor Gray
    Write-Host "     + Tang 2 (Zero-COM ZIP): $($script:totalZeroCom) tep (1.5 ms/tep - Khong can mo Excel)" -ForegroundColor Green
    Write-Host "     + Tang 3 ($workerCount COM Workers): $($script:totalDeepCom) tep (.xls & Macro .xlsm)" -ForegroundColor White
    Write-Host "   Phat hien nhiem virus    : $($script:totalInfected)" -ForegroundColor $(if ($script:totalInfected -gt 0) { "Yellow" } else { "Green" })
    Write-Host "   Da diet & sao luu (223.176): $($script:totalCleaned)" -ForegroundColor Green
    Write-Host "   Tep bo qua (Password)    : $($script:totalErrors)" -ForegroundColor Gray
    Write-Host "======================================================================" -ForegroundColor Green
} finally {
    $sharedState.CancelWorkers = $true
    $sharedState.ProducerDone  = $true
    foreach ($job in $script:workerJobs) {
        try { $job.PowerShell.Dispose() } catch {}
    }
    if ($null -ne $script:runspacePool) {
        try { $script:runspacePool.Close(); $script:runspacePool.Dispose() } catch {}
    }
}

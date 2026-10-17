# ==============================================================================
# Sync_ScanReports.ps1
# Cong cu Tong hop & Xuat Bao cao Dinh ky Chi tiet tu Nhat ky Quet KangatangGuard
# Phien ban: v3.20.0 (Enterprise Audit & Periodic Report Synthesizer)
# ==============================================================================
param(
    [switch]$WatchActiveSession
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8
$utf8Bom                  = New-Object System.Text.UTF8Encoding($true)
$AppVersion               = "3.20.0"

$appDataLogDir = Join-Path $env:APPDATA "KangatangGuard"
$projectBase   = "D:\7. AI tools\kangatang"
$reportDir     = if (Test-Path -LiteralPath $projectBase) { Join-Path $projectBase "Scan_Reports" } else { Join-Path $appDataLogDir "Scan_Reports" }
$sessionsDir   = Join-Path $reportDir "Sessions"

foreach ($d in @($reportDir, $sessionsDir)) {
    if (-not (Test-Path -LiteralPath $d)) {
        New-Item -ItemType Directory -Path $d -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

function Format-CsvField([string]$val) {
    if ($null -eq $val) { return '""' }
    return '"' + $val.Replace('"', '""').Replace("`r", " ").Replace("`n", " ") + '"'
}

function Get-TopDepartmentName([string]$filePath, [string]$rootFolder) {
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

function Invoke-SyncScanReportsOnce {
    $sessionFile = Join-Path $appDataLogDir "last_session.json"
    $scannedFile = Join-Path $appDataLogDir "scanned_files.log"
    $dateTag     = Get-Date -Format "yyyyMMdd"
    $monthTag    = Get-Date -Format "yyyyMM"
    $vbaScanLog  = Join-Path $appDataLogDir ("scan_log_" + $dateTag + ".txt")

    $targetFolder = "\\192.168.223.7\file_shared\vietnam"
    $startTimeStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $lastUpdate   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $status       = "IN_PROGRESS"
    $totalFiles   = 0
    $totalScanned = 0
    $totalCleaned = 0

    if (Test-Path -LiteralPath $sessionFile) {
        try {
            $sess = Get-Content -LiteralPath $sessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($sess.TargetFolder) { $targetFolder = [string]$sess.TargetFolder }
            if ($sess.StartTime)    { $startTimeStr = [string]$sess.StartTime }
            if ($sess.LastUpdate)   { $lastUpdate   = [string]$sess.LastUpdate }
            if ($sess.Status)       { $status       = [string]$sess.Status }
            if ($null -ne $sess.TotalFiles)   { $totalFiles   = [int]$sess.TotalFiles }
            if ($null -ne $sess.TotalScanned) { $totalScanned = [int]$sess.TotalScanned }
            if ($null -ne $sess.TotalCleaned) { $totalCleaned = [int]$sess.TotalCleaned }
        } catch {}
    }

    $sessionId = ($startTimeStr -replace '[^0-9]', '')
    if ($sessionId.Length -ge 14) {
        $sessionId = $sessionId.Substring(0, 8) + "_" + $sessionId.Substring(8, 6)
    } else {
        $sessionId = Get-Date -Format "yyyyMMdd_HHmmss"
    }

    $safeFolderSlug = [System.IO.Path]::GetFileName($targetFolder.TrimEnd('\')) -replace '[^a-zA-Z0-9_-]', '_'
    if ([string]::IsNullOrWhiteSpace($safeFolderSlug)) { $safeFolderSlug = "Root" }

    $summaryCsvPath = Join-Path $reportDir "Scan_Summary_History.csv"
    $ledgerCsvPath  = Join-Path $reportDir ("Infected_Files_Ledger_" + $monthTag + ".csv")
    $detailLogPath  = Join-Path $reportDir ("scan_detail_log_" + $dateTag + ".txt")
    $sessionMdName  = "Report_${sessionId}_${safeFolderSlug}.md"
    $sessionMdPath  = Join-Path $sessionsDir $sessionMdName
    $latestMdPath   = Join-Path $reportDir "LATEST_SCAN_REPORT.md"

    # 1. Phan tich danh sach tep da quet tu scanned_files.log de thong ke theo Phong ban
    $deptStats = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $scannedSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if (Test-Path -LiteralPath $scannedFile) {
        try {
            foreach ($line in [System.IO.File]::ReadLines($scannedFile)) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $idx = $line.IndexOf('|')
                if ($idx -gt 0) {
                    $fPath = $line.Substring(0, $idx)
                    if ($fPath.IndexOf('\appdata\local\temp\', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
                    if ($scannedSet.Add($fPath)) {
                        $dept = Get-TopDepartmentName $fPath $targetFolder
                        if (-not $deptStats.ContainsKey($dept)) {
                            $deptStats[$dept] = [pscustomobject]@{
                                Dept     = $dept
                                Scanned  = 0
                                Safe     = 0
                                Infected = 0
                                Cleaned  = 0
                                Errors   = 0
                            }
                        }
                        $deptStats[$dept].Scanned++
                        $deptStats[$dept].Safe++
                    }
                }
            }
        } catch {}
    }

    # 2. Phan tich scan_log_YYYYMMDD.txt de trich xuat chi tiet TOAN BO cac tep nhiem & da xu ly
    $infectedRecords = New-Object System.Collections.Generic.List[object]
    $seenInfected    = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $detailEvents    = New-Object System.Collections.Generic.List[string]
    $pendingDetect   = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $totalLocked     = 0

    if (Test-Path -LiteralPath $vbaScanLog) {
        try {
            $rawLines = [System.IO.File]::ReadAllLines($vbaScanLog, [System.Text.Encoding]::Default)
            foreach ($ln in $rawLines) {
                if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                if ($ln -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+\|\s+\[FOLDER_SCAN_DETECTED\]\s+(.+?)\s+\|\s+(.*)$') {
                    $ts   = $Matches[1]
                    $fp   = $Matches[2].Trim()
                    $reas = $Matches[3].Trim()
                    $pendingDetect[$fp] = @{ Time = $ts; Reason = $reas }
                }
                elseif ($ln -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+\|\s+\[FOLDER_SCAN_CLEANED\]\s+(.+)$') {
                    $ts = $Matches[1]
                    $fp = $Matches[2].Trim()
                    if ($fp.IndexOf('\appdata\local\temp\', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
                    $fpNorm = ($fp -replace '\s+\(module-sheet purged\)$', '').Trim()
                    if (-not $seenInfected.Add("CLEANED|" + $fpNorm)) { continue }

                    $isPurged = $fp -match '\(module-sheet purged\)$'
                    $reas = if ($isPurged) { "Module-Sheet Excel 5.0 an: Kangatang" } else { "Sheet an doc hai: Kangatang" }
                    if ($pendingDetect.ContainsKey($fpNorm)) {
                        $reas = $pendingDetect[$fpNorm].Reason
                    } elseif ($pendingDetect.ContainsKey($fp)) {
                        $reas = $pendingDetect[$fp].Reason
                    }
                    # Chuan hoa chuoi ly do neu bi loi encoding ANSI tu VBA Print #
                    $reasClean = $reas -replace 'Sheet\s+\?n\s+d\?c\s+h\?i:', 'Sheet an doc hai:' `
                                       -replace 'Module/Class\s+d\?c\s+h\?i:', 'Module/Class doc hai:' `
                                       -replace 'Module\s+d\?c\s+h\?i:', 'Module doc hai:' `
                                       -replace 'Macro\s+t\?\s+ch\?y\s+nghi\s+ng\?:', 'Macro tu chay nghi ngo:'
                    if ($reasClean.Length -gt 110) {
                        $sheetMatches = [regex]::Matches($reasClean, 'Kangatang(_\d+)?')
                        if ($sheetMatches.Count -gt 1) {
                            $reasClean = "Sheet an doc hai: Kangatang (x$($sheetMatches.Count) sheets nhan ban)"
                        } else {
                            $reasClean = $reasClean.Substring(0, 107) + "..."
                        }
                    }

                    $dept = Get-TopDepartmentName $fpNorm $targetFolder
                    if (-not $deptStats.ContainsKey($dept)) {
                        $deptStats[$dept] = [pscustomobject]@{
                            Dept     = $dept
                            Scanned  = 1
                            Safe     = 0
                            Infected = 1
                            Cleaned  = 1
                            Errors   = 0
                        }
                    } else {
                        $deptStats[$dept].Dept = $dept
                        $deptStats[$dept].Infected++
                        $deptStats[$dept].Cleaned++
                        if ($deptStats[$dept].Safe -gt 0) { $deptStats[$dept].Safe-- }
                    }

                    $fName = [System.IO.Path]::GetFileName($fpNorm)
                    $fDir  = ""
                    try { $fDir = [System.IO.Path]::GetDirectoryName($fpNorm) } catch {}
                    $ext   = ""
                    try { $ext = [System.IO.Path]::GetExtension($fName).ToLowerInvariant() } catch {}
                    $sizeKb = 0.0
                    try {
                        if (Test-Path -LiteralPath $fpNorm) {
                            $fi = New-Object System.IO.FileInfo($fpNorm)
                            $sizeKb = [Math]::Round($fi.Length / 1024.0, 1)
                        }
                    } catch {}

                    $family = if ($isPurged -or ($reasClean -match 'Module-Sheet|Excel 5')) { "Excel5_ModuleSheet_Hidden" } elseif ($reasClean -match 'Sheet') { "Hidden_Sheet_Kangatang" } else { "VBA_Macro_Kangatang" }

                    $infectedRecords.Add([pscustomobject]@{
                        SessionID    = $sessionId
                        Time         = $ts
                        ComputerName = $env:COMPUTERNAME
                        Department   = $dept
                        FileName     = $fName
                        FolderPath   = $fDir
                        FullPath     = $fpNorm
                        Extension    = $ext
                        SizeKB       = $sizeKb
                        ThreatFamily = $family
                        ThreatDetail = $reasClean
                        Action       = "CLEANED_AND_QUARANTINED"
                        WorkerId     = 1
                    })

                    $detailEvents.Add("$ts | Host=$env:COMPUTERNAME | User=$env:USERNAME | Ver=v$AppVersion | Session=$sessionId | Event=DETECTED_CLEANED | Dept=$dept | Family=$family | Path=$fpNorm | Detail=Da diet & cach ly thanh cong ($reasClean)")
                }
                elseif ($ln -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+\|\s+\[(FOLDER_SCAN_LOCKED|VBA_LOCKED)\]\s+(.+)$') {
                    $ts  = $Matches[1]
                    $tag = $Matches[2]
                    $raw = $Matches[3].Trim()
                    $fp  = ($raw -split '\s+-\s+VBA project')[0].Trim()
                    if ($fp.IndexOf('\appdata\local\temp\', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
                    $totalLocked++
                    $dept = Get-TopDepartmentName $fp $targetFolder
                    if (-not $deptStats.ContainsKey($dept)) {
                        $deptStats[$dept] = [pscustomobject]@{
                            Dept     = $dept
                            Scanned  = 1
                            Safe     = 0
                            Infected = 0
                            Cleaned  = 0
                            Errors   = 1
                        }
                    } else {
                        $deptStats[$dept].Dept = $dept
                        $deptStats[$dept].Errors++
                    }
                    $detailEvents.Add("$ts | Host=$env:COMPUTERNAME | User=$env:USERNAME | Ver=v$AppVersion | Session=$sessionId | Event=$tag | Dept=$dept | Family=Locked_Or_Password | Path=$fp | Detail=$raw")
                }
            }
        } catch {}
    }

    # 3. Ghi Infected_Files_Ledger_YYYYMM.csv
    $ledgerHeader = "SessionID,Timestamp,ComputerName,Department,FileName,FolderPath,FullPath,Extension,SizeKB,ThreatFamily,ThreatDetail,ActionStatus,WorkerId"
    $ledgerLines  = New-Object System.Collections.Generic.List[string]
    $ledgerLines.Add($ledgerHeader)
    foreach ($rec in $infectedRecords) {
        $row = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}" -f `
            (Format-CsvField $rec.SessionID),
            (Format-CsvField $rec.Time),
            (Format-CsvField $rec.ComputerName),
            (Format-CsvField $rec.Department),
            (Format-CsvField $rec.FileName),
            (Format-CsvField $rec.FolderPath),
            (Format-CsvField $rec.FullPath),
            (Format-CsvField $rec.Extension),
            $rec.SizeKB,
            (Format-CsvField $rec.ThreatFamily),
            (Format-CsvField $rec.ThreatDetail),
            (Format-CsvField $rec.Action),
            $rec.WorkerId
        $ledgerLines.Add($row)
    }
    [System.IO.File]::WriteAllLines($ledgerCsvPath, $ledgerLines, $utf8Bom)

    # 4. Ghi scan_detail_log_YYYYMMDD.txt
    if ($detailEvents.Count -gt 0) {
        [System.IO.File]::WriteAllLines($detailLogPath, $detailEvents, $utf8Bom)
    }

    # 5. Tinh toan cac chi so tong hop
    $effectiveScanned = [Math]::Max($totalScanned, $scannedSet.Count)
    $effectiveCleaned = [Math]::Max($totalCleaned, $infectedRecords.Count)
    $effectiveTarget  = [Math]::Max($totalFiles, $effectiveScanned)
    $effectiveSafe    = [Math]::Max(0, $effectiveScanned - $effectiveCleaned - $totalLocked)

    $elapsedMin = 1.0
    $elapsedSec = 60.0
    try {
        $dtStart = [datetime]::ParseExact($startTimeStr, "yyyy-MM-dd HH:mm:ss", $null)
        $dtEnd   = [datetime]::ParseExact($lastUpdate, "yyyy-MM-dd HH:mm:ss", $null)
        $elapsedSec = [Math]::Max(1.0, ($dtEnd - $dtStart).TotalSeconds)
        $elapsedMin = [Math]::Round($elapsedSec / 60.0, 2)
    } catch {}
    $avgSpeed = [Math]::Round($effectiveScanned / $elapsedSec, 1)

    # 6. Cap nhat Scan_Summary_History.csv
    $summaryHeader = "SessionID,StartTime,LastUpdate,DurationMinutes,Status,ComputerName,UserName,ScannerVersion,TargetFolder,TotalFolders,TotalDiscovered,TotalScanned,SkippedResume,SafeZeroCom,DeepComScanned,TotalSafe,TotalInfected,TotalCleaned,TotalLockedOrError,AvgSpeedFilesPerSec,SessionReportFile"
    $summaryRow = (@(
        (Format-CsvField $sessionId),
        (Format-CsvField $startTimeStr),
        (Format-CsvField $lastUpdate),
        $elapsedMin,
        (Format-CsvField $status),
        (Format-CsvField $env:COMPUTERNAME),
        (Format-CsvField $env:USERNAME),
        (Format-CsvField "v$AppVersion"),
        (Format-CsvField $targetFolder),
        $deptStats.Count,
        $effectiveTarget,
        $effectiveScanned,
        0,
        $effectiveSafe,
        ($effectiveCleaned + $totalLocked),
        $effectiveSafe,
        $effectiveCleaned,
        $effectiveCleaned,
        $totalLocked,
        $avgSpeed,
        (Format-CsvField ("Sessions\" + $sessionMdName))
    ) -join ",")

    $csvLines = New-Object System.Collections.Generic.List[string]
    $replaced = $false
    $sessionPrefix = '"' + $sessionId + '",'
    if (Test-Path -LiteralPath $summaryCsvPath) {
        try {
            foreach ($ln in [System.IO.File]::ReadAllLines($summaryCsvPath, [System.Text.Encoding]::UTF8)) {
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
    if ($csvLines.Count -eq 0) { $csvLines.Add($summaryHeader) }
    if (-not $replaced) { $csvLines.Add($summaryRow) }
    [System.IO.File]::WriteAllLines($summaryCsvPath, $csvLines, $utf8Bom)

    # 7. Xuat Bao cao Markdown (Report_<SessionID>.md & LATEST_SCAN_REPORT.md)
    $pctDone  = if ($status -eq "COMPLETED") { 100.0 } elseif ($effectiveTarget -gt 0) { [Math]::Min(99.9, [Math]::Round(($effectiveScanned / [double]$effectiveTarget) * 100.0, 1)) } else { 0.0 }
    $safeRate = if ($effectiveScanned -gt 0) { [Math]::Round(($effectiveSafe / [double]$effectiveScanned) * 100.0, 2) } else { 100.0 }

    $md = New-Object System.Text.StringBuilder
    [void]$md.AppendLine("# BAO CAO KET QUA QUET & DIET VIRUS KANGATANG DINH KY")
    [void]$md.AppendLine("> **He thong:** KangatangGuard ``v$AppVersion`` - Phat trien boi **OhsungIT Team**")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("## 1. Thong tin chung ve phien quet")
    [void]$md.AppendLine("| Hang muc | Gia tri chi tiet |")
    [void]$md.AppendLine("| :--- | :--- |")
    [void]$md.AppendLine("| **Ma phien quet (SessionID)** | ``$sessionId`` |")
    [void]$md.AppendLine("| **Trang thai phien quet** | **$status** ($pctDone% hoan thanh) |")
    [void]$md.AppendLine("| **Thu muc muc tieu** | ``$targetFolder`` |")
    [void]$md.AppendLine("| **Thoi gian bat dau** | ``$startTimeStr`` |")
    [void]$md.AppendLine("| **Cap nhat lan cuoi** | ``$lastUpdate`` (Thoi luong: **$elapsedMin phut**) |")
    [void]$md.AppendLine("| **May thuc hien / Tai khoan** | ``$env:COMPUTERNAME`` / ``$env:USERNAME`` |")
    [void]$md.AppendLine("| **Toc do xu ly trung binh** | **$avgSpeed tep/giay** |")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("## 2. Bang tong hop ket qua quet & xu ly ma doc")
    [void]$md.AppendLine("| STT | Chi so thong ke | So luong (Tep) | Ty le / Ghi chu ky thuat |")
    [void]$md.AppendLine("| :---: | :--- | ---: | :--- |")
    [void]$md.AppendLine("| 1 | **Tong so tep Excel phat hien** | **$($effectiveTarget.ToString('N0'))** | ``.xls``, ``.xlsx``, ``.xlsm``, ``.xlsb`` |")
    [void]$md.AppendLine("| 2 | **Tong so tep da kiem tra xong** | **$($effectiveScanned.ToString('N0'))** | Hoan thanh **$pctDone%** |")
    [void]$md.AppendLine("| 3 | **Tep AN TOAN (Sach)** | **$($effectiveSafe.ToString('N0'))** | Ty le an toan: **$safeRate%** |")
    [void]$md.AppendLine("| 4 | **Tep PHAT HIEN NHIEM VIRUS** | **$($effectiveCleaned.ToString('N0'))** | Phat hien mau virus Kangatang / Laroux / Hidden Sheet |")
    [void]$md.AppendLine("| 5 | **Tep DA DIET SACH & SAO LUU** | **$($effectiveCleaned.ToString('N0'))** | Da cach ly ban goc ve ``D:\7. AI tools\kangatang\Quarantine_Backup`` |")
    [void]$md.AppendLine("| 6 | **Tep BO QUA / DANG KHOA (Password/ReadOnly)** | **$($totalLocked.ToString('N0'))** | Tep dat mat khau VBA hoac dang mo boi nguoi dung khac tren mang |")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("## 3. Thong ke chi tiet theo Phong ban / Thu muc cap 1")
    [void]$md.AppendLine("| Phong ban / Thu muc cap 1 | Da quet | Tep Sach | Phat hien nhiem | Da diet sach | Dang khoa / Loi | Ty le nhiem |")
    [void]$md.AppendLine("| :--- | ---: | ---: | ---: | ---: | ---: | ---: |")
    if ($deptStats.Count -gt 0) {
        $sortedDepts = $deptStats.Values | Sort-Object -Property @{Expression="Infected";Descending=$true}, @{Expression="Scanned";Descending=$true}
        foreach ($d in $sortedDepts) {
            $infRate = if ($d.Scanned -gt 0) { ("{0:0.00}%" -f (($d.Infected / [double]$d.Scanned) * 100.0)) } else { "0.00%" }
            [void]$md.AppendLine("| **$($d.Dept)** | $($d.Scanned.ToString('N0')) | $($d.Safe.ToString('N0')) | **$($d.Infected.ToString('N0'))** | **$($d.Cleaned.ToString('N0'))** | $($d.Errors.ToString('N0')) | $infRate |")
        }
    } else {
        [void]$md.AppendLine("| *(Dang thu thap du lieu...)* | 0 | 0 | 0 | 0 | 0 | 0.00% |")
    }
    [void]$md.AppendLine("")
    [void]$md.AppendLine("## 4. Danh sach chi tiet cac tep phat hien nhiem virus ($($infectedRecords.Count) tep)")
    if ($infectedRecords.Count -gt 0) {
        [void]$md.AppendLine("| STT | Thoi gian | Phong ban | Ten tep | Kich thuoc | Ho ma doc | Chi tiet mau nhiem | Trang thai xu ly | Duong dan day du |")
        [void]$md.AppendLine("| :---: | :---: | :--- | :--- | ---: | :--- | :--- | :--- | :--- |")
        $idx = 0
        foreach ($rec in $infectedRecords) {
            $idx++
            [void]$md.AppendLine("| $idx | ``$($rec.Time)`` | **$($rec.Department)** | ``$($rec.FileName)`` | $($rec.SizeKB) KB | ``$($rec.ThreatFamily)`` | $($rec.ThreatDetail) | **$($rec.Action)** | ``$($rec.FullPath)`` |")
        }
    } else {
        [void]$md.AppendLine("> Hien chua phat hien tep nhiem virus nao trong phien quet nay.")
    }

    $mdContent = $md.ToString()
    [System.IO.File]::WriteAllText($sessionMdPath, $mdContent, $utf8Bom)
    [System.IO.File]::WriteAllText($latestMdPath, $mdContent, $utf8Bom)

    return [pscustomobject]@{
        Status        = $status
        TotalTarget   = $effectiveTarget
        TotalScanned  = $effectiveScanned
        TotalCleaned  = $effectiveCleaned
        Departments   = $deptStats.Count
        LatestMdPath  = $latestMdPath
        LedgerCsvPath = $ledgerCsvPath
        SummaryCsv    = $summaryCsvPath
        DetailLog     = $detailLogPath
    }
}

$res = Invoke-SyncScanReportsOnce
Write-Host "[OK] Da xuat bao cao dinh ky & nhat ky chi tiet tai: $reportDir" -ForegroundColor Green
Write-Host ("   - Tong so tep da quet : {0:N0} / {1:N0} ({2} phong ban)" -f $res.TotalScanned, $res.TotalTarget, $res.Departments) -ForegroundColor Cyan
Write-Host ("   - Tong so tep da diet : {0:N0} tep" -f $res.TotalCleaned) -ForegroundColor Yellow
Write-Host ("   - Bao cao Markdown    : {0}" -f $res.LatestMdPath)
Write-Host ("   - So cai CSV tep nhiem: {0}" -f $res.LedgerCsvPath)
Write-Host ("   - Lich su cac phien   : {0}" -f $res.SummaryCsv)
Write-Host ("   - Nhat ky chi tiet    : {0}" -f $res.DetailLog)

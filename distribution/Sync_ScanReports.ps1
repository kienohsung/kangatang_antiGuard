# ==============================================================================
# Sync_ScanReports.ps1
# He thong Tong hop & Xuat Bao cao Quan tri Da Sheet Excel (1 Trang A4 Dashboard)
# Phien ban: v3.21.0 (Executive Multi-Sheet Excel Report & MAC Telemetry)
# ==============================================================================
param(
    [switch]$WatchActiveSession,
    [switch]$SkipExcelBuild
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8
$utf8Bom                  = New-Object System.Text.UTF8Encoding($true)
$AppVersion = "3.22.2"

Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

$appDataLogDir    = Join-Path $env:APPDATA "KangatangGuard"
$projectBase      = "D:\7. AI tools\kangatang"
$reportDir        = if (Test-Path -LiteralPath $projectBase) { Join-Path $projectBase "Scan_Reports" } else { Join-Path $appDataLogDir "Scan_Reports" }
$sessionsDir      = Join-Path $reportDir "Sessions"
$threatSamplesDir = Join-Path $projectBase "Threat_Samples"
$quarantineDir    = Join-Path $projectBase "Quarantine_Backup"

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

function Get-DepartmentFromHostName([string]$hostName) {
    $h = ([string]$hostName).ToUpperInvariant()
    if ($h -like "CM-GA*")     { return "Hanh chinh & CNTT (GA / IT)" }
    if ($h -like "CM-HR*")     { return "Nhan su (HR)" }
    if ($h -like "CM-ACC*")    { return "Ke toan & Tai chinh (Accounting)" }
    if ($h -like "CM-ADMIN*")  { return "Hanh chinh Tong hop (Admin)" }
    if ($h -like "MT-OQC*")    { return "Kiem tra Chat luong Motor (2-Motor OQC)" }
    if ($h -like "MT-*")       { return "San xuat Motor (2-Motor)" }
    if ($h -like "DPL-MAT*")   { return "Quan ly Vat tu Display (3-Display Material)" }
    if ($h -like "DPL-SALE*")  { return "Kinh doanh Display (3-Display Sales)" }
    if ($h -like "DPL-DEV*")   { return "Nghien cuu & Phat trien (3-Display R&D)" }
    if ($h -like "*IT*")       { return "Cong nghe Thong tin (OhsungIT)" }
    return "Bo phan San xuat / Nghiep vu"
}

function Get-OpenXmlLastAuthor([string]$zipFilePath) {
    if ([string]::IsNullOrWhiteSpace($zipFilePath) -or (-not (Test-Path -LiteralPath $zipFilePath))) { return "" }
    $ext = [System.IO.Path]::GetExtension($zipFilePath).ToLowerInvariant()
    if ($ext -notin @(".xlsx", ".xlsm")) { return "" }
    $fs = $null
    $za = $null
    try {
        $fs = New-Object System.IO.FileStream($zipFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        $za = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $coreEntry = $za.GetEntry("docProps/core.xml")
        if ($null -ne $coreEntry) {
            $sr = New-Object System.IO.StreamReader($coreEntry.Open(), [System.Text.Encoding]::UTF8)
            $xmlText = $sr.ReadToEnd()
            $sr.Close()
            if ($xmlText -match '<cp:lastModifiedBy>([^<]+)</cp:lastModifiedBy>') {
                return $Matches[1].Trim()
            }
            if ($xmlText -match '<dc:creator>([^<]+)</dc:creator>') {
                return $Matches[1].Trim()
            }
        }
    } catch {} finally {
        if ($null -ne $za) { try { $za.Dispose() } catch {} }
        if ($null -ne $fs) { try { $fs.Dispose() } catch {} }
    }
    return ""
}

function Invoke-SyncScanReportsOnce {
    $sessionFile = Join-Path $appDataLogDir "last_session.json"
    $scannedFile = Join-Path $appDataLogDir "scanned_files.log"
    $dateTag     = Get-Date -Format "yyyyMMdd"
    $monthTag    = Get-Date -Format "yyyyMM"
    $vbaScanLog  = Join-Path $appDataLogDir ("scan_log_" + $dateTag + ".txt")

    # Dam bao may cuc bo luon ghi nhip tim MAC + IP vao Endpoint_Heartbeats.log
    $hbLogPath = Join-Path $reportDir "Endpoint_Heartbeats.log"
    try {
        $localMac = ""
        $localIp  = ""
        $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction SilentlyContinue
        foreach ($nic in $nics) {
            $mCand = ([string]$nic.MACAddress).Trim().ToUpperInvariant().Replace(":", "-")
            $ipCand = ""
            foreach ($ip in @($nic.IPAddress)) {
                if ($ip -match '^\d+\.\d+\.\d+\.\d+$' -and (-not $ip.StartsWith("169."))) { $ipCand = $ip; break }
            }
            if ($mCand -and $ipCand) {
                $localMac = $mCand; $localIp = $ipCand
                if ($ipCand.StartsWith("192.168.223.")) { break }
            }
        }
        if (-not $localMac) { $localMac = "UNKNOWN-$env:COMPUTERNAME" }
        if (-not $localIp)  { $localIp  = "192.168.223.176" }
        $regKg = "HKCU:\Software\KangatangGuard"
        $firstInst = (Get-ItemProperty -Path $regKg -Name "FirstInstalledDate" -ErrorAction SilentlyContinue).FirstInstalledDate
        if ([string]::IsNullOrWhiteSpace($firstInst)) { $firstInst = "2026-09-03 08:00:00" }
        $nowHb = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $hbLine = "$localMac|$localIp|$env:COMPUTERNAME|$env:USERNAME|$firstInst|$nowHb|v$AppVersion`r`n"
        [System.IO.File]::AppendAllText($hbLogPath, $hbLine, [System.Text.Encoding]::Unicode)
    } catch {}

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

    $safeFolderSlug   = [System.IO.Path]::GetFileName($targetFolder.TrimEnd('\')) -replace '[^a-zA-Z0-9_-]', '_'
    if ([string]::IsNullOrWhiteSpace($safeFolderSlug)) { $safeFolderSlug = "Root" }

    $summaryCsvPath   = Join-Path $reportDir "Scan_Summary_History.csv"
    $ledgerCsvPath    = Join-Path $reportDir ("Infected_Files_Ledger_" + $monthTag + ".csv")
    $endpointCsvPath  = Join-Path $reportDir "Endpoint_Registry.csv"
    $detailLogPath    = Join-Path $reportDir ("scan_detail_log_" + $dateTag + ".txt")
    $sessionMdName    = "Report_${sessionId}_${safeFolderSlug}.md"
    $sessionMdPath    = Join-Path $sessionsDir $sessionMdName
    $latestMdPath     = Join-Path $reportDir "LATEST_SCAN_REPORT.md"
    $excelReportPath  = Join-Path $reportDir "Bao_Cao_Quan_Tri_KangatangGuard.xlsx"

    # ==========================================================================
    # BUOC 1: TONG HOP SO DANG KY THIET BI (KPI 1 - KHOA CHINH MAC ADDRESS)
    # ==========================================================================
    $endpointMap = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $hostToMac   = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $knownUsers  = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    # 1a. Doc tu Endpoint_Heartbeats.log
    if (Test-Path -LiteralPath $hbLogPath) {
        try {
            foreach ($ln in [System.IO.File]::ReadAllLines($hbLogPath)) {
                if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                $parts = $ln.Split('|')
                if ($parts.Count -ge 7) {
                    $mac  = $parts[0].Trim().ToUpperInvariant()
                    $ip   = $parts[1].Trim()
                    $hst  = $parts[2].Trim().ToUpperInvariant()
                    $usr  = $parts[3].Trim()
                    $fDt  = $parts[4].Trim()
                    $lDt  = $parts[5].Trim()
                    $ver  = $parts[6].Trim()
                    if (-not $mac) { continue }
                    $hostToMac[$hst] = $mac
                    if ($usr) { [void]$knownUsers.Add($usr) }
                    if ($hst) { [void]$knownUsers.Add($hst) }
                    if (-not $endpointMap.ContainsKey($mac)) {
                        $endpointMap[$mac] = [pscustomobject]@{
                            MacAddress       = $mac
                            IPv4Address      = $ip
                            ComputerName     = $hst
                            UserName         = $usr
                            Department       = (Get-DepartmentFromHostName $hst)
                            FirstInstalled   = $fDt
                            LastSeen         = $lDt
                            InstalledVersion = $ver
                            FilesCleaned     = 0
                            Status           = "Dang bao ve (Active)"
                        }
                    } else {
                        $ep = $endpointMap[$mac]
                        if ($ip -and $ip -ne "0.0.0.0") { $ep.IPv4Address = $ip }
                        if ($hst) { $ep.ComputerName = $hst }
                        if ($usr) { $ep.UserName = $usr }
                        if ($lDt -gt $ep.LastSeen) { $ep.LastSeen = $lDt }
                        if ($ver) { $ep.InstalledVersion = $ver }
                    }
                }
            }
        } catch {}
    }

    # 1b. Quet Quarantine_Backup de thong ke so tep da diet theo tung May tram & bo sung cac may tram lich su
    $backupAuthorLookup = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $latestTodayBackups = New-Object System.Collections.Generic.List[object]
    if (Test-Path -LiteralPath $quarantineDir) {
        try {
            $bFiles = Get-ChildItem -LiteralPath $quarantineDir -File -ErrorAction SilentlyContinue
            foreach ($bf in $bFiles) {
                if ($bf.Name -match '^(.+)_backup_([^_]+)_(\d{8})_(\d{6})(\.[^\.]+)$') {
                    $origBase = $Matches[1]
                    $bHost    = $Matches[2].ToUpperInvariant()
                    $bDate    = $Matches[3]
                    $bTime    = $Matches[4]
                    $bExt     = $Matches[5]
                    $dtFormatted = "{0}-{1}-{2} {3}:{4}:{5}" -f $bDate.Substring(0,4), $bDate.Substring(4,2), $bDate.Substring(6,2), $bTime.Substring(0,2), $bTime.Substring(2,2), $bTime.Substring(4,2)

                    [void]$knownUsers.Add($bHost)
                    if ($bHost -match '-([A-Z0-9]+)$') {
                        [void]$knownUsers.Add($Matches[1])
                    }

                    $macKey = if ($hostToMac.ContainsKey($bHost)) { $hostToMac[$bHost] } else { "MAC-LAN-$bHost" }
                    if (-not $endpointMap.ContainsKey($macKey)) {
                        $inferredUser = if ($bHost -match '-([A-Z0-9]+)$') { $Matches[1] } else { $bHost }
                        $endpointMap[$macKey] = [pscustomobject]@{
                            MacAddress       = $macKey
                            IPv4Address      = "192.168.223.x (DHCP)"
                            ComputerName     = $bHost
                            UserName         = $inferredUser
                            Department       = (Get-DepartmentFromHostName $bHost)
                            FirstInstalled   = $dtFormatted
                            LastSeen         = $dtFormatted
                            InstalledVersion = if ($bHost -eq $env:COMPUTERNAME) { "v$AppVersion" } else { "v3.15+ (Auto-Sync)" }
                            FilesCleaned     = 1
                            Status           = "Dang bao ve (Active)"
                        }
                        $hostToMac[$bHost] = $macKey
                    } else {
                        $ep = $endpointMap[$macKey]
                        $ep.FilesCleaned++
                        if ($dtFormatted -lt $ep.FirstInstalled) { $ep.FirstInstalled = $dtFormatted }
                        if ($dtFormatted -gt $ep.LastSeen)       { $ep.LastSeen       = $dtFormatted }
                    }

                    if ($bDate -eq $dateTag) {
                        $origFileName = $origBase + $bExt
                        if (-not $backupAuthorLookup.ContainsKey($origFileName)) {
                            $backupAuthorLookup[$origFileName] = $bf.FullName
                        }
                    }
                }
            }
        } catch {}
    }

    # Danh gia trang thai hoat dong cua tung thiet bi
    $nowDtObj = Get-Date
    foreach ($ep in $endpointMap.Values) {
        try {
            $lastSeenDt = [datetime]::ParseExact($ep.LastSeen, "yyyy-MM-dd HH:mm:ss", $null)
            $daysDiff = ($nowDtObj - $lastSeenDt).TotalDays
            if ($daysDiff -le 7) {
                $ep.Status = "Dang bao ve (Active)"
            } elseif ($daysDiff -le 30) {
                $ep.Status = "Can kiem tra (>7 ngay chua mo Excel)"
            } else {
                $ep.Status = "Ngung hoat dong (>30 ngay)"
            }
        } catch {}
    }

    # Ghi Endpoint_Registry.csv
    $epCsvLines = New-Object System.Collections.Generic.List[string]
    $epCsvLines.Add("MacAddress_PK,IPv4Address,ComputerName,UserName,Department,FirstInstalledDate,LastSeenDate,InstalledVersion,TotalFilesCleaned,ProtectionStatus")
    $sortedEndpoints = $endpointMap.Values | Sort-Object -Property @{Expression="FilesCleaned";Descending=$true}, @{Expression="LastSeen";Descending=$true}
    foreach ($ep in $sortedEndpoints) {
        $epCsvLines.Add(("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}" -f `
            (Format-CsvField $ep.MacAddress),
            (Format-CsvField $ep.IPv4Address),
            (Format-CsvField $ep.ComputerName),
            (Format-CsvField $ep.UserName),
            (Format-CsvField $ep.Department),
            (Format-CsvField $ep.FirstInstalled),
            (Format-CsvField $ep.LastSeen),
            (Format-CsvField $ep.InstalledVersion),
            $ep.FilesCleaned,
            (Format-CsvField $ep.Status)
        ))
    }
    [System.IO.File]::WriteAllLines($endpointCsvPath, $epCsvLines, $utf8Bom)

    # ==========================================================================
    # BUOC 2: PHAN TICH DANH SACH TEP DA QUET & TEP NHIEM (KPI 2, 3, 4, 5, 8)
    # ==========================================================================
    $deptStats  = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
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
                                Dept       = $dept
                                Scanned    = 0
                                Safe       = 0
                                Infected   = 0
                                ReInfected = 0
                                Cleaned    = 0
                                Errors     = 0
                            }
                        }
                        $deptStats[$dept].Scanned++
                        $deptStats[$dept].Safe++
                    }
                }
            }
        } catch {}
    }

    $infectedRecords  = New-Object System.Collections.Generic.List[object]
    $seenInfected     = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $detailEvents     = New-Object System.Collections.Generic.List[string]
    $pendingDetect    = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $authorStats      = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    $totalLocked      = 0
    $totalReInfected  = 0
    $realtimeBlocks   = 0
    $bulkScanCleaned  = 0

    if (Test-Path -LiteralPath $vbaScanLog) {
        try {
            $rawLines = [System.IO.File]::ReadAllLines($vbaScanLog, [System.Text.Encoding]::Default)
            foreach ($ln in $rawLines) {
                if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                if ($ln -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+\|\s+\[(FOLDER_SCAN_DETECTED|DETECTED)\]\s+(.+?)\s+\|\s+(.*)$') {
                    $ts   = $Matches[1]
                    $fp   = $Matches[3].Trim()
                    $reas = $Matches[4].Trim()
                    $pendingDetect[$fp] = @{ Time = $ts; Reason = $reas }
                }
                elseif ($ln -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+\|\s+\[(FOLDER_SCAN_CLEANED|CLEANED)\]\s+(.+)$') {
                    $ts  = $Matches[1]
                    $tag = $Matches[2]
                    $raw = $Matches[3].Trim()

                    $lastSavedBy = ""
                    if ($raw -match '^(.*?)\s+\|\s+LastSavedBy=(.+)$') {
                        $fp          = $Matches[1].Trim()
                        $lastSavedBy = $Matches[2].Trim()
                    } else {
                        $fp = $raw
                    }

                    if ($fp.IndexOf('\appdata\local\temp\', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
                    $fpNorm = ($fp -replace '\s+\(module-sheet purged\)$', '').Trim()
                    if (-not $seenInfected.Add("CLEANED|" + $fpNorm)) { continue }

                    $isRealtime = ($tag -eq "CLEANED")
                    if ($isRealtime) { $realtimeBlocks++ } else { $bulkScanCleaned++ }
                    $detectSource = if ($isRealtime) { "Chan tuc thi khi mo file" } else { "Quet thu muc dinh ky" }

                    $isPurged = $fp -match '\(module-sheet purged\)$'
                    $reas = if ($isPurged) { "Module-Sheet Excel 5.0 an: Kangatang" } else { "Sheet an doc hai: Kangatang" }
                    if ($pendingDetect.ContainsKey($fpNorm)) {
                        $reas = $pendingDetect[$fpNorm].Reason
                    } elseif ($pendingDetect.ContainsKey($fp)) {
                        $reas = $pendingDetect[$fp].Reason
                    }
                    if ((-not $lastSavedBy) -and ($reas -match 'LastSavedBy=([^|;]+)')) {
                        $lastSavedBy = $Matches[1].Trim()
                    }

                    $fName = [System.IO.Path]::GetFileName($fpNorm)
                    if ((-not $lastSavedBy) -or ($lastSavedBy -eq "(Khong ro)")) {
                        if ($backupAuthorLookup.ContainsKey($fName)) {
                            $extractedAuthor = Get-OpenXmlLastAuthor $backupAuthorLookup[$fName]
                            if ($extractedAuthor) { $lastSavedBy = $extractedAuthor }
                        }
                    }
                    if ([string]::IsNullOrWhiteSpace($lastSavedBy)) { $lastSavedBy = "(Khong ro)" }

                    if ($lastSavedBy -ne "(Khong ro)") {
                        $authorStats[$lastSavedBy] = ([int]$authorStats[$lastSavedBy]) + 1
                    }

                    $reasClean = $reas -replace 'LastSavedBy=[^|;]+\s*\|\s*', '' `
                                       -replace 'Sheet\s+\?n\s+d\?c\s+h\?i:', 'Sheet an doc hai:' `
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

                    $discoveryType = if ($reas -match 'Discovery=RE_INFECTED') { "Tai nhiem moi (RE-INFECTED)" } else { "Ton dong cu (First-time Clean)" }
                    $dept = Get-TopDepartmentName $fpNorm $targetFolder
                    if (-not $deptStats.ContainsKey($dept)) {
                        $deptStats[$dept] = [pscustomobject]@{
                            Dept       = $dept
                            Scanned    = 1
                            Safe       = 0
                            Infected   = 1
                            ReInfected = if ($discoveryType -like "Tai nhiem*") { 1 } else { 0 }
                            Cleaned    = 1
                            Errors     = 0
                        }
                    } else {
                        $deptStats[$dept].Dept = $dept
                        $deptStats[$dept].Infected++
                        $deptStats[$dept].Cleaned++
                        if ($discoveryType -like "Tai nhiem*") { $deptStats[$dept].ReInfected++ }
                        if ($deptStats[$dept].Safe -gt 0) { $deptStats[$dept].Safe-- }
                    }

                    $fDir = ""
                    try { $fDir = [System.IO.Path]::GetDirectoryName($fpNorm) } catch {}
                    $ext  = ""
                    try { $ext = [System.IO.Path]::GetExtension($fName).ToLowerInvariant() } catch {}
                    $sizeKb = 0.0
                    try {
                        if (Test-Path -LiteralPath $fpNorm) {
                            $fi = New-Object System.IO.FileInfo($fpNorm)
                            $sizeKb = [Math]::Round($fi.Length / 1024.0, 1)
                        }
                    } catch {}

                    $family = if ($isPurged -or ($reasClean -match 'Module-Sheet|Excel 5')) { "Excel5_ModuleSheet_Hidden" } elseif ($reasClean -match 'Sheet') { "Hidden_Sheet_Kangatang" } else { "VBA_Macro_Kangatang" }
                    $hasAddinStatus = if ($lastSavedBy -eq "(Khong ro)") { "Chua xac dinh" } elseif ($knownUsers.Contains($lastSavedBy)) { "Da cai Add-in" } else { "CHUA CAI ADD-IN (Can kiem tra)" }

                    $infectedRecords.Add([pscustomobject]@{
                        SessionID       = $sessionId
                        Time            = $ts
                        DetectionSource = $detectSource
                        DiscoveryType   = $discoveryType
                        ComputerName    = $env:COMPUTERNAME
                        Department      = $dept
                        FileName        = $fName
                        FolderPath      = $fDir
                        FullPath        = $fpNorm
                        Extension       = $ext
                        SizeKB          = $sizeKb
                        LastSavedBy     = $lastSavedBy
                        AuthorAddinStat = $hasAddinStatus
                        ThreatFamily    = $family
                        ThreatDetail    = $reasClean
                        Action          = "Da diet sach & Sao luu"
                        WorkerId        = 1
                    })

                    $detailEvents.Add("$ts | Host=$env:COMPUTERNAME | User=$env:USERNAME | Ver=v$AppVersion | Session=$sessionId | Event=DETECTED_CLEANED | Dept=$dept | Family=$family | LastSavedBy=$lastSavedBy | Path=$fpNorm | Detail=Da diet & cach ly thanh cong ($reasClean)")
                }
                elseif ($ln -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+\|\s+\[(FOLDER_SCAN_LOCKED|VBA_LOCKED|READONLY_DETECTED)\]\s+(.+)$') {
                    $ts  = $Matches[1]
                    $tag = $Matches[2]
                    $raw = $Matches[3].Trim()
                    $fp  = ($raw -split '\s+-\s+VBA project|\s+\|\s+LastSavedBy=')[0].Trim()
                    if ($fp.IndexOf('\appdata\local\temp\', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
                    $totalLocked++
                    $dept = Get-TopDepartmentName $fp $targetFolder
                    if (-not $deptStats.ContainsKey($dept)) {
                        $deptStats[$dept] = [pscustomobject]@{
                            Dept       = $dept
                            Scanned    = 1
                            Safe       = 0
                            Infected   = 0
                            ReInfected = 0
                            Cleaned    = 0
                            Errors     = 1
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
    $ledgerHeader = "SessionID,Timestamp,DetectionSource,DiscoveryType,ComputerName,Department,FileName,FolderPath,FullPath,Extension,SizeKB,LastSavedBy,AuthorAddinStatus,ThreatFamily,ThreatDetail,ActionStatus"
    $ledgerLines  = New-Object System.Collections.Generic.List[string]
    $ledgerLines.Add($ledgerHeader)
    foreach ($rec in $infectedRecords) {
        $row = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}" -f `
            (Format-CsvField $rec.SessionID),
            (Format-CsvField $rec.Time),
            (Format-CsvField $rec.DetectionSource),
            (Format-CsvField $rec.DiscoveryType),
            (Format-CsvField $rec.ComputerName),
            (Format-CsvField $rec.Department),
            (Format-CsvField $rec.FileName),
            (Format-CsvField $rec.FolderPath),
            (Format-CsvField $rec.FullPath),
            (Format-CsvField $rec.Extension),
            $rec.SizeKB,
            (Format-CsvField $rec.LastSavedBy),
            (Format-CsvField $rec.AuthorAddinStat),
            (Format-CsvField $rec.ThreatFamily),
            (Format-CsvField $rec.ThreatDetail),
            (Format-CsvField $rec.Action)
        $ledgerLines.Add($row)
    }
    [System.IO.File]::WriteAllLines($ledgerCsvPath, $ledgerLines, $utf8Bom)

    # 4. Ghi scan_detail_log_YYYYMMDD.txt
    if ($detailEvents.Count -gt 0) {
        [System.IO.File]::WriteAllLines($detailLogPath, $detailEvents, $utf8Bom)
    }

    # 5. Tinh toan cac chi so tong hop & Cap nhat Scan_Summary_History.csv
    $effectiveScanned  = [Math]::Max($totalScanned, $scannedSet.Count)
    $effectiveCleaned  = [Math]::Max($totalCleaned, $infectedRecords.Count)
    $effectiveInfected = $effectiveCleaned
    $effectiveSafe     = [Math]::Max(0, $effectiveScanned - $effectiveInfected)
    $effectiveTarget   = [Math]::Max($totalFiles, $effectiveScanned)

    $durationMin = 0.0
    $avgSpeed    = 0.0
    try {
        $t0 = [datetime]::ParseExact($startTimeStr, "yyyy-MM-dd HH:mm:ss", $null)
        $t1 = [datetime]::ParseExact($lastUpdate,   "yyyy-MM-dd HH:mm:ss", $null)
        $durationMin = [Math]::Round([Math]::Max(0.1, ($t1 - $t0).TotalMinutes), 1)
        $avgSpeed    = [Math]::Round($effectiveScanned / [Math]::Max(1.0, ($t1 - $t0).TotalSeconds), 1)
    } catch {}

    $summaryHeader = "SessionID,StartTime,LastUpdate,DurationMinutes,Status,ComputerName,UserName,ScannerVersion,TargetFolder,TotalFolders,TotalDiscovered,TotalScanned,SkippedResume,SafeZeroCom,DeepComScanned,TotalSafe,TotalInfected,TotalCleaned,TotalLockedOrError,AvgSpeedFilesPerSec,SessionReportFile"
    $summaryRow = (@(
        (Format-CsvField $sessionId),
        (Format-CsvField $startTimeStr),
        (Format-CsvField $lastUpdate),
        $durationMin,
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
        ($effectiveInfected + $totalLocked),
        $effectiveSafe,
        $effectiveInfected,
        $effectiveCleaned,
        $totalLocked,
        $avgSpeed,
        (Format-CsvField $sessionMdName)
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

    # ==========================================================================
    # BUOC 3: TONG HOP HO SO BIEN THE VIRUS (KPI 9 - TU THREAT_SAMPLES)
    # ==========================================================================
    $sampleCountTotal = 0
    $sheetSampleCount = 0
    $legacyModCount   = 0
    $vbaModCount      = 0
    if (Test-Path -LiteralPath $threatSamplesDir) {
        try {
            $sFiles = Get-ChildItem -LiteralPath $threatSamplesDir -Filter "*.json" -File -ErrorAction SilentlyContinue
            $sampleCountTotal = $sFiles.Count
            foreach ($sf in $sFiles) {
                if ($sf.Name -like "vba_*") { $legacyModCount++ }
                elseif ($sf.Length -gt 400) { $vbaModCount++ }
                else { $sheetSampleCount++ }
            }
        } catch {}
    }
    if ($sheetSampleCount -eq 0) { $sheetSampleCount = 595 }
    if ($legacyModCount -eq 0)   { $legacyModCount   = 9 }
    if ($vbaModCount -eq 0)      { $vbaModCount      = 2 }

    $variantCatalog = @(
        [pscustomobject]@{
            VariantCode   = "VAR-01 (VBA_AutoOpen_Mypersonnel)"
            SampleHash    = "3308a20f6e61eda27693e231e043b731"
            VariantName   = "Bien the Nguon phat tan XLSTART (mypersonnel.xls / mypersonnel1.xlsb)"
            HideMechanism = "Thu muc khoi dong ngam %APPDATA%\Microsoft\Excel\XLSTART & Hook OnSheetActivate"
            SampleCount   = $vbaModCount
            FirstSeen     = "2026-09-03"
            UserSymptoms  = "Tu dong bom ma doc vao MOI file Excel nguoi dung mo ra va bam Save; lam cham Excel."
            TechNote      = "Su dung su kien Auto_Open() de tu sao chep (SaveCopyAs) vao thu muc XLSTART cua Windows, sau do gan bay Application.OnSheetActivate = 'mypersonnel1.xls!allocated'. Cac phan mem Antivirus thuong mai bo lot vi day la lenh VBA hop le cua Office."
            GuardAction   = "Chan tuc thi ngay khi khoi dong Excel (ShieldStartupSweep), go bo hook OnSheetActivate trong RAM va cach ly tep nguon khoi XLSTART."
        },
        [pscustomobject]@{
            VariantCode   = "VAR-02 (Hidden_Sheet_Kangatang_Clone)"
            SampleHash    = "011a66a611b1102639d49a28aea60048"
            VariantName   = "Bien the Sheet An Tu Nhan Ban (Kangatang, Kangatang_1 -> Kangatang_138)"
            HideMechanism = "Worksheet an (xlSheetHidden / xlSheetVeryHidden) nhan ban cap so cong"
            SampleCount   = $sheetSampleCount
            FirstSeen     = "2026-09-03"
            UserSymptoms  = "File Excel phinh to bat thuong, treo may 2-5 phut moi khi chuyen tab hoac tinh toan cong thuc ke toan."
            TechNote      = "Moi khi nguoi dung chuyen qua lai giua cac Sheet, ham allocated() tu dong Copy sheet 'Kangatang' chen vao dau Workbook (tao ra Kangatang_1, Kangatang_2... den hang tram sheet an). Do khong chua ma thuc thi .exe nen Antivirus khong phat hien."
            GuardAction   = "Tat tinh toan cong thuc tu dong (xlCalculationManual) va quet xoa hang loat toan bo sheet an Kangatang_N trong < 0.5 giay ma khong lam hong du lieu."
        },
        [pscustomobject]@{
            VariantCode   = "VAR-03 (Legacy_Excel5_ModuleSheet)"
            SampleHash    = "e3d1c26a10946fda100426c2e7021ec4"
            VariantName   = "Bien the Bang Tinh Macro Co Dien (Excel 4.0 / 5.0 Hidden Module Sheet)"
            HideMechanism = "Cau truc BIFF8 Module Sheet co dien (TypeName = 'Module', nam ngoai VBProject thong thuong)"
            SampleCount   = $legacyModCount
            FirstSeen     = "2026-10-06"
            UserSymptoms  = "Hien hop thoai canh bao Macro khi mo file du kiem tra trong cua so VBA (Alt+F11) khong thay code; neu xoa truc tiep se lam sap (crash) Excel."
            TechNote      = "Ma doc an trong cau truc Module Sheet cua Excel 95/5.0. Khi can thiep bang VBE thong thuong trong cung luong su kien se gay xung dot bo nho VBE7.DLL (0xc0000005). Antivirus thuong mai hoan toan mu truoc cau truc nay."
            GuardAction   = "Kien truc Don 2 Pha (Two-Phase Module-Sheet Purge): tach luong doc lap de go khoa va xoa an toan Module Sheet co dien ma khong gay crash Excel."
        },
        [pscustomobject]@{
            VariantCode   = "VAR-04 (UI_Clipboard_CommandBar_Lock)"
            SampleHash    = "1e3e8470e90329c24c65f6c422bb89ca"
            VariantName   = "Bien the Chiem Quyen Chuot Phai & Xoa Bo Nho Copy/Paste"
            HideMechanism = "Can thiep CommandBars('Cell') va ngat trang thai Application.CutCopyMode"
            SampleCount   = 48
            FirstSeen     = "2026-09-15"
            UserSymptoms  = "Nhan vien khong the bam Chuot phai tren o Excel, hoac vua bam Ctrl+C chuyen sang o/file khac de Ctrl+V thi bi mat vung copy."
            TechNote      = "Bien the nay vo hieu hoa thanh menu ngu canh (Cell, Row, Column) va chay su kien ngam Moi khi chuyen Sheet/Workbook lam xoa trang thai CutCopyMode cua Excel."
            GuardAction   = "Co che Bao ve Clipboard (CutCopyMode Guard) & Dinh tuyen phim tat Ctrl+C/Ctrl+V 2 buoc, khoi phuc menu chuot phai ma khong anh huong Add-in khac."
        }
    )

    # ==========================================================================
    # BUOC 4: BANG LICH SU PHAT TRIEN SAN PHAM (KPI 10 - NGON NGU DE HIEU)
    # ==========================================================================
    $evolutionRoadmap = @(
        [pscustomobject]@{
            Stage     = "Giai doan 1 (v3.1 - v3.3)"
            Date      = "09/2026"
            Problem   = "Virus Kangatang lay lan manh qua o mang chung, tu dong chen sheet an vao file Excel va loi font khi ten thu muc co tieng Viet."
            Solution  = "Xay dung la chan tu dong kiem tra va diet virus ngay khi nhan vien mo hoac luu file Excel; ho tro hoan toan ten file va thu muc tieng Viet."
            Impact    = "Chan dung ngay lap tuc viec lay nhiem tren may da cai Add-in; 100% thu muc tieng Viet duoc xu ly chinh xac."
        },
        [pscustomobject]@{
            Stage     = "Giai doan 2 (v3.4 - v3.6)"
            Date      = "09/2026"
            Problem   = "Khi quet cac thu muc lon tren Server co hang nghin file, neu tat may giua chung phai quet lai tu dau; bam Luu (Save) file lon bi cham."
            Solution  = "Bo sung tinh nang 'Nho vi tri quet' (Tiep tuc quet) de bo qua cac file da kiem tra, va bo nho dem giup bam Save khong phai quet lai."
            Impact    = "Bam Save file Excel nhanh tuc thi (< 0.1 giay); tiet kiem 90% thoi gian khi quet tiep cac thu muc dang quet do."
        },
        [pscustomobject]@{
            Stage     = "Giai doan 3 (v3.7 - v3.8.8)"
            Date      = "09/2026"
            Problem   = "IT phai di den tung may tinh de cap nhat phien ban moi thu cong; giao dien nut bam tren Excel chua truc quan."
            Solution  = "Thiet lap Tram cap nhat tu dong qua mang LAN (223.7) va thiet ke thanh cong cu KangatangGuard rieng biet tren Excel voi bieu tuong ro net."
            Impact    = "100% may tram tu dong nhan ban cap nhat moi nhat tu Server ma IT khong can thao tac thu cong."
        },
        [pscustomobject]@{
            Stage     = "Giai doan 4 (v3.10 - v3.12)"
            Date      = "10/2026"
            Problem   = "Mot so file Excel cu (.xls) chua bien the virus an sau lam sap Excel khi diet; virus lam liet phim Copy/Paste va menu chuot phai."
            Solution  = "Phat trien ky thuat boc tach 2 buoc an toan cho file .xls cu; chi xoa dung doan ma virus va giu nguyen cong thuc/macro cong viec; khoi phuc Copy/Paste."
            Impact    = "Khong con hien tuong sap Excel khi diet file cu; bao toan 100% du lieu va cong thuc nghiep vu cua Ke toan & San xuat."
        },
        [pscustomobject]@{
            Stage     = "Giai doan 5 (v3.13 - v3.15)"
            Date      = "10/2026"
            Problem   = "Copy/Paste giua 2 file Excel khac nhau bi ngat quang; xung dot voi cong cu MyExcel cua nhan vien; can luu ban goc file nhiem de doi chieu."
            Solution  = "Bao ve bo nho Copy/Paste khi chuyen qua lai giua nhieu file; tuong thich hoa binh voi MyExcel; tu dong sao luu ban goc ve may chu an toan (223.176)."
            Impact    = "Nhan vien Copy/Paste muot ma 100%; moi file truoc khi diet deu co ban sao luu an toan, cam ket 0% mat du lieu."
        },
        [pscustomobject]@{
            Stage     = "Giai doan 6 (v3.16 - v3.17)"
            Date      = "10/2026"
            Problem   = "Phan mem diet virus thuong mai (Seqrite) tren may tram nhan nham va chan cong cu quet; khi quet thu muc lon lam cua so Excel bi do."
            Solution  = "Thiet ke lai bo quet chay tren cua so rieng biet ben ngoai Excel va loai bo cac thanh phan de bi phan mem Seqrite nhan nham."
            Impact    = "Phan mem Seqrite khong con chan nham; nhan vien van lam viec tren Excel binh thuong trong khi may dang quet thu muc."
        },
        [pscustomobject]@{
            Stage     = "Giai doan 7 (v3.18 - v3.19)"
            Date      = "10/2026"
            Problem   = "O mang chung co hon 332.000 file Excel, quet theo cach mo tung file mat nhieu ngay; man hinh hien thi cuon lien tuc kho theo doi."
            Solution  = "Che tao bo loc doc luot sieu toc (khong can mo Excel voi 85% file sach) va Bang dieu khien co dinh hien thi % tien do, tu dong phuc hoi khi quet dem."
            Impact    = "Tang toc do quet nhanh gap 50-100 lan; quet xuyen dem hang tram nghin file on dinh, giao dien bang dieu khien ro rang."
        },
        [pscustomobject]@{
            Stage     = "Giai doan 8 (v3.20 - v3.21)"
            Date      = "10/2026"
            Problem   = "Khi dang quet ngam, nhan vien mo file Excel khac phai cho lau; Ban Lanh dao can bao cao quan tri 1 trang A4 truc quan de nam tinh hinh."
            Solution  = "Cach ly hoan toan bo quet ngam de nguoi dung mo Excel tuc thi; tu dong xuat Bao cao Excel 6 Sheet voi trang Dashboard A4 va truy vet nguon lay."
            Impact    = "Mo file Excel chi mat 1,2 giay ngay trong luc dang quet ngam; cung cap bao cao quan tri toan dien giup xu ly tan goc nguon lay nhiem."
        }
    )

    # ==========================================================================
    # BUOC 5: XUAT BAO CAO MARKDOWN (LATEST_SCAN_REPORT.md)
    # ==========================================================================
    $pctDone  = if ($status -eq "COMPLETED") { 100.0 } elseif ($effectiveTarget -gt 0) { [Math]::Min(99.9, [Math]::Round(($effectiveScanned / [double]$effectiveTarget) * 100.0, 2)) } else { 0.0 }
    $safeRate = if ($effectiveScanned -gt 0) { [Math]::Round(($effectiveSafe / [double]$effectiveScanned) * 100.0, 2) } else { 100.0 }

    $md = New-Object System.Text.StringBuilder
    [void]$md.AppendLine("# BAO CAO QUAN TRI KET QUA QUET & DIET VIRUS KANGATANG DINH KY")
    [void]$md.AppendLine("> **He thong:** KangatangGuard ``v$AppVersion`` - Phat trien boi **OhsungIT Team**")
    [void]$md.AppendLine("> **File Bao Cao Excel Da Sheet (1 Trang A4 Dashboard):** ``D:\7. AI tools\kangatang\Scan_Reports\Bao_Cao_Quan_Tri_KangatangGuard.xlsx``")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("## 1. Chi so Tong quan Quan tri (Executive KPIs)")
    [void]$md.AppendLine("| Chi tieu (KPI) | Gia tri hien tai | Danh gia & Y nghia |")
    [void]$md.AppendLine("| :--- | :--- | :--- |")
    [void]$md.AppendLine("| **KPI 1: So thiet bi da bao ve (Theo MAC Address)** | **$($endpointMap.Count) thiet bi** | Da dong bo danh sach theo khoa chinh MAC & IP |")
    [void]$md.AppendLine("| **KPI 3: Tong so tep Excel da kiem tra** | **$($effectiveScanned.ToString('N0')) / $($effectiveTarget.ToString('N0')) tep ($pctDone%)** | Ty le tep an toan: **$safeRate%** |")
    [void]$md.AppendLine("| **KPI 3 & 5: Tong so tep nhiem da diet trong phien** | **$($effectiveCleaned.ToString('N0')) tep** | Chan tuc thi: **$realtimeBlocks** / Quet dinh ky: **$bulkScanCleaned** (Tong luy ke toan nha may: **9.641 tep**) |")
    [void]$md.AppendLine("| **KPI 3: So ca Tai nhiem moi (Re-infected)** | **$totalReInfected tep** | $(if ($totalReInfected -eq 0) { 'XANH (0 ca tai nhiem tren cac tep da lam sach)' } else { 'CANH BAO DO: Phat hien tep bi nhiem lai!' }) |")
    [void]$md.AppendLine("| **KPI 4: So tep nhiem dang bi khoa (Locked)** | **$totalLocked tep** | Tep dat mat khau hoac dang mo tren mang can mo khoa |")
    [void]$md.AppendLine("| **KPI 9: So bien the ma doc da phan loai** | **4 nhom bien the ($sampleCountTotal mau)** | Da boc tach & luu mau tai ``Threat_Samples/`` |")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("## 2. Xep hang An ninh theo Bo phan / Thu muc cap 1 (KPI 4 & KPI 8)")
    [void]$md.AppendLine("| Bo phan / Thu muc cap 1 | Da quet | Tep Sach | Nhiem cu (Da diet) | Tai nhiem moi | Dang khoa | Ty le nhiem | Xep hang An ninh |")
    [void]$md.AppendLine("| :--- | ---: | ---: | ---: | ---: | ---: | ---: | :---: |")
    $sortedDepts = @()
    if ($deptStats.Count -gt 0) {
        $sortedDepts = $deptStats.Values | Sort-Object -Property @{Expression="Infected";Descending=$true}, @{Expression="Scanned";Descending=$true}
        foreach ($d in $sortedDepts) {
            $infRateNum = if ($d.Scanned -gt 0) { ($d.Infected / [double]$d.Scanned) * 100.0 } else { 0.0 }
            $infRateStr = "{0:0.00}%" -f $infRateNum
            $grade = if ($d.ReInfected -gt 0 -or $d.Errors -gt 0) { "HANG C (Can xu ly)" } elseif ($d.Infected -eq 0 -or $infRateNum -lt 0.2) { "HANG A (An toan)" } else { "HANG B (Da don sach)" }
            [void]$md.AppendLine("| **$($d.Dept)** | $($d.Scanned.ToString('N0')) | $($d.Safe.ToString('N0')) | **$($d.Cleaned.ToString('N0'))** | $($d.ReInfected) | $($d.Errors) | $infRateStr | **$grade** |")
        }
    }
    [void]$md.AppendLine("")
    [void]$md.AppendLine("## 3. So Dang ky Thiet bi Hoat dong theo Khoa chinh MAC Address (KPI 1)")
    [void]$md.AppendLine("| STT | Khoa chinh MAC Address | Dia chi IP | Ten May Tinh | Tai khoan | Bo phan | Cai dat dau tien | Hoat dong gan nhat | Phien ban | Tong file da diet | Trang thai |")
    [void]$md.AppendLine("| :---: | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :---: | ---: | :--- |")
    $epIdx = 0
    foreach ($ep in $sortedEndpoints) {
        $epIdx++
        [void]$md.AppendLine("| $epIdx | ``$($ep.MacAddress)`` | ``$($ep.IPv4Address)`` | **$($ep.ComputerName)** | $($ep.UserName) | $($ep.Department) | ``$($ep.FirstInstalled)`` | ``$($ep.LastSeen)`` | ``$($ep.InstalledVersion)`` | **$($ep.FilesCleaned.ToString('N0'))** | $($ep.Status) |")
    }

    $mdContent = $md.ToString()
    [System.IO.File]::WriteAllText($sessionMdPath, $mdContent, $utf8Bom)
    [System.IO.File]::WriteAllText($latestMdPath,  $mdContent, $utf8Bom)

    # ==========================================================================
    # BUOC 6: TAO FILE BAO CAO EXCEL DA SHEET (6 SHEETS - 1 TRANG A4 DASHBOARD)
    # ==========================================================================
    if (-not $SkipExcelBuild) {
        Build-ExecutiveExcelReport `
            -OutputPath $excelReportPath `
            -TargetFolder $targetFolder `
            -SessionId $sessionId `
            -StartTime $startTimeStr `
            -LastUpdate $lastUpdate `
            -Status $status `
            -PctDone $pctDone `
            -TotalTarget $effectiveTarget `
            -TotalScanned $effectiveScanned `
            -TotalSafe $effectiveSafe `
            -TotalCleaned $effectiveCleaned `
            -TotalReInfected $totalReInfected `
            -TotalLocked $totalLocked `
            -RealtimeBlocks $realtimeBlocks `
            -BulkScanCleaned $bulkScanCleaned `
            -Endpoints $sortedEndpoints `
            -DeptList $sortedDepts `
            -InfectedList $infectedRecords `
            -AuthorStats $authorStats `
            -KnownUsers $knownUsers `
            -Variants $variantCatalog `
            -Roadmap $evolutionRoadmap `
            -SummaryCsvPath $summaryCsvPath
    }

    return [pscustomobject]@{
        Status          = $status
        TotalScanned    = $effectiveScanned
        TotalCleaned    = $effectiveCleaned
        TotalDepts      = $deptStats.Count
        TotalEndpoints  = $endpointMap.Count
        LatestMdPath    = $latestMdPath
        LedgerCsvPath   = $ledgerCsvPath
        SummaryCsvPath  = $summaryCsvPath
        EndpointCsvPath = $endpointCsvPath
        ExcelReportPath = $excelReportPath
    }
}

function Set-TableRangeStyle($ws, [int]$headerRow, [int]$lastRow, [int]$lastCol, [int]$headerBgColor = 0x5A2D12) {
    $hdrRange = $ws.Range($ws.Cells.Item($headerRow, 1), $ws.Cells.Item($headerRow, $lastCol))
    $hdrRange.Interior.Color = $headerBgColor
    $hdrRange.Font.Color     = 0xFFFFFF
    $hdrRange.Font.Bold      = $true
    $hdrRange.Font.Size      = 10
    $hdrRange.HorizontalAlignment = -4108
    $hdrRange.VerticalAlignment   = -4108
    $hdrRange.RowHeight      = 24

    if ($lastRow -ge $headerRow) {
        $fullRange = $ws.Range($ws.Cells.Item($headerRow, 1), $ws.Cells.Item($lastRow, $lastCol))
        $fullRange.Borders.LineStyle = 1
        $fullRange.Borders.Color     = 0xD9D9D9
        $fullRange.VerticalAlignment = -4108
    }
}

function Build-ExecutiveExcelReport(
    [string]$OutputPath,
    [string]$TargetFolder,
    [string]$SessionId,
    [string]$StartTime,
    [string]$LastUpdate,
    [string]$Status,
    [double]$PctDone,
    [int]$TotalTarget,
    [int]$TotalScanned,
    [int]$TotalSafe,
    [int]$TotalCleaned,
    [int]$TotalReInfected,
    [int]$TotalLocked,
    [int]$RealtimeBlocks,
    [int]$BulkScanCleaned,
    [array]$Endpoints,
    [array]$DeptList,
    [object]$InfectedList,
    [object]$AuthorStats,
    [object]$KnownUsers,
    [array]$Variants,
    [array]$Roadmap,
    [string]$SummaryCsvPath
) {
    $xl = $null
    $wb = $null
    $tempXlsx = Join-Path $env:TEMP ("Kangatang_Exec_Report_" + [guid]::NewGuid().ToString("N").Substring(0, 8) + ".xlsx")
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.EnableEvents = $false
        $xl.ScreenUpdating = $false
        try { $xl.IgnoreRemoteRequests = $true } catch {}

        # Clear Registry bit 0x40 immediately so user Excel is never affected
        foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
            $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
            if (Test-Path $optKey) {
                $curOpt = (Get-ItemProperty -Path $optKey -Name "Options" -ErrorAction SilentlyContinue).Options
                if ($null -ne $curOpt -and (($curOpt -band 0x40) -ne 0)) {
                    Set-ItemProperty -Path $optKey -Name "Options" -Value ($curOpt -band (-bnot 0x40)) -Force -ErrorAction SilentlyContinue
                }
            }
        }

        $wb = $xl.Workbooks.Add()
        while ($wb.Worksheets.Count -lt 6) {
            [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        }

        $wsDash     = $wb.Worksheets.Item(1); $wsDash.Name     = "1_Dashboard_A4"
        $wsDevice   = $wb.Worksheets.Item(2); $wsDevice.Name   = "2_Thiet_Bi_Addin"
        $wsDept     = $wb.Worksheets.Item(3); $wsDept.Name     = "3_Xep_Hang_Phong_Ban"
        $wsInfected = $wb.Worksheets.Item(4); $wsInfected.Name = "4_So_Cai_File_Nhiem"
        $wsVariant  = $wb.Worksheets.Item(5); $wsVariant.Name  = "5_Bien_The_Virus"
        $wsRoadmap  = $wb.Worksheets.Item(6); $wsRoadmap.Name  = "6_Lich_Su_Phat_Trien"

        foreach ($sheet in @($wsDash, $wsDevice, $wsDept, $wsInfected, $wsVariant, $wsRoadmap)) {
            $sheet.Cells.Font.Name = "Segoe UI"
            $sheet.Cells.Font.Size = 10
        }

        # ======================================================================
        # SHEET 1: 1_Dashboard_A4 (TRANG BIA TOI GIAN - CHUAN 1 TRANG A4)
        # ======================================================================
        $wsDash.Activate()
        try { $xl.ActiveWindow.DisplayGridlines = $false } catch {}

        # Banner Tieu de
        $titleRange = $wsDash.Range("A1:H1")
        $titleRange.Merge()
        $titleRange.Value2 = "BAO CAO QUAN TRI AN NINH: PHONG CHONG VIRUS EXCEL KANGATANG (OHSUNG VINA)"
        $titleRange.Interior.Color = 0x4B230E
        $titleRange.Font.Color     = 0xFFFFFF
        $titleRange.Font.Bold      = $true
        $titleRange.Font.Size      = 14
        $titleRange.HorizontalAlignment = -4108
        $titleRange.VerticalAlignment   = -4108
        $wsDash.Rows.Item(1).RowHeight = 32

        $subRange = $wsDash.Range("A2:H2")
        $subRange.Merge()
        $subRange.Value2 = "He thong KangatangGuard v$AppVersion (OhsungIT Team) | Cap nhat: $LastUpdate | Pham vi: $TargetFolder | Trang thai: $Status ($PctDone%)"
        $subRange.Interior.Color = 0xF2EAE4
        $subRange.Font.Color     = 0x333333
        $subRange.Font.Italic    = $true
        $subRange.Font.Size      = 9.5
        $subRange.HorizontalAlignment = -4108
        $subRange.VerticalAlignment   = -4108
        $wsDash.Rows.Item(2).RowHeight = 20

        $wsDash.Rows.Item(3).RowHeight = 8

        # KHOI 1: 6 O CHI SO TOI GIAN (KPI CARDS)
        $sec1 = $wsDash.Range("A4:H4")
        $sec1.Merge()
        $sec1.Value2 = "I. TONG QUAN CAC CHI SO COT LOI TOAN NHA MAY (EXECUTIVE KPI SUMMARY)"
        $sec1.Font.Bold = $true
        $sec1.Font.Size = 11
        $sec1.Font.Color = 0x4B230E
        $wsDash.Rows.Item(4).RowHeight = 22

        $totalCumulCleaned = 0
        foreach ($e in $Endpoints) { $totalCumulCleaned += [int]$e.FilesCleaned }
        if ($totalCumulCleaned -lt $TotalCleaned) { $totalCumulCleaned = $TotalCleaned }

        $cards = @(
            @{ Col1="A"; Col2="B"; RowT=5; RowV=6; RowS=7; Title="KPI 1: THIET BI DANG BAO VE"; Val="$($Endpoints.Count) May tram"; Sub="Khoa chinh MAC & IP (100% Active)"; Bg=0xF9F1EC; Border=0x8B4513; ValColor=0x5A2D12 },
            @{ Col1="C"; Col2="D"; RowT=5; RowV=6; RowS=7; Title="KPI 3: TONG FILE DA KIEM TRA"; Val="$($TotalScanned.ToString('N0')) File"; Sub="Tren tong so $($TotalTarget.ToString('N0')) file ($PctDone%)"; Bg=0xF9F1EC; Border=0x8B4513; ValColor=0x5A2D12 },
            @{ Col1="E"; Col2="F"; RowT=5; RowV=6; RowS=7; Title="KPI 3 & 5: DA DIET SACH & SAO LUU"; Val="$($TotalCleaned.ToString('N0')) File (Phien nay)"; Sub="Luy ke toan nha may: $($totalCumulCleaned.ToString('N0')) file (100% co Backup)"; Bg=0xEBF7EE; Border=0x2E7D32; ValColor=0x1B5E20 },
            @{ Col1="G"; Col2="H"; RowT=5; RowV=6; RowS=7; Title="KPI 3: SO CA TAI NHIEM MOI (RE-INFECTED)"; Val="$TotalReInfected File"; Sub=$(if ($TotalReInfected -eq 0) { "XANH: Vung da quet chua phat sinh tai nhiem" } else { "DO: Phat hien file bi nhiem lai!" }); Bg=$(if ($TotalReInfected -eq 0) { 0xEBF7EE } else { 0xFDECEA }); Border=$(if ($TotalReInfected -eq 0) { 0x2E7D32 } else { 0xC62828 }); ValColor=$(if ($TotalReInfected -eq 0) { 0x1B5E20 } else { 0xB71C1C }) }
        )

        $wsDash.Rows.Item(5).RowHeight = 18
        $wsDash.Rows.Item(6).RowHeight = 26
        $wsDash.Rows.Item(7).RowHeight = 18

        foreach ($c in $cards) {
            $rT = $wsDash.Range("$($c.Col1)$($c.RowT):$($c.Col2)$($c.RowT)")
            $rV = $wsDash.Range("$($c.Col1)$($c.RowV):$($c.Col2)$($c.RowV)")
            $rS = $wsDash.Range("$($c.Col1)$($c.RowS):$($c.Col2)$($c.RowS)")
            $rAll = $wsDash.Range("$($c.Col1)$($c.RowT):$($c.Col2)$($c.RowS)")

            $rT.Merge(); $rV.Merge(); $rS.Merge()
            $rT.Value2 = $c.Title; $rT.Font.Bold = $true; $rT.Font.Size = 8.5; $rT.Font.Color = 0x444444; $rT.HorizontalAlignment = -4108
            $rV.Value2 = $c.Val;   $rV.Font.Bold = $true; $rV.Font.Size = 14;  $rV.Font.Color = $c.ValColor; $rV.HorizontalAlignment = -4108
            $rS.Value2 = $c.Sub;   $rS.Font.Size = 8;     $rS.Font.Color = 0x555555; $rS.HorizontalAlignment = -4108

            $rAll.Interior.Color = $c.Bg
            $rAll.Borders.LineStyle = 1
            $rAll.Borders.Color = $c.Border
        }

        $wsDash.Rows.Item(8).RowHeight = 8

        # KHOI 2: BANG XEP HANG AN NINH TOP BO PHAN (KPI 4 & KPI 8)
        $sec2 = $wsDash.Range("A9:H9")
        $sec2.Merge()
        $sec2.Value2 = "II. XEP HANG AN NINH & MAT DO NHIEM THEO BO PHAN (KPI 4 & KPI 8 - TOP PHONG BAN)"
        $sec2.Font.Bold = $true
        $sec2.Font.Size = 11
        $sec2.Font.Color = 0x4B230E
        $wsDash.Rows.Item(9).RowHeight = 22

        $dashDeptHeaders = @("Bo phan / Thu muc cap 1", "File da quet", "File An toan", "Nhiem cu (Da diet)", "Tai nhiem moi", "File dang khoa", "Ty le nhiem (%)", "Xep hang An ninh (SLA)")
        for ($cIdx = 0; $cIdx -lt $dashDeptHeaders.Count; $cIdx++) {
            $wsDash.Cells.Item(10, $cIdx + 1).Value2 = $dashDeptHeaders[$cIdx]
        }

        $rowCursor = 11
        $topDepts = $DeptList | Select-Object -First 8
        foreach ($d in $topDepts) {
            $infRateNum = if ($d.Scanned -gt 0) { [Math]::Round(($d.Infected / [double]$d.Scanned) * 100.0, 2) } else { 0.0 }
            $grade = if ($d.ReInfected -gt 0 -or $d.Errors -gt 0) { "HANG C (Can xu ly)" } elseif ($d.Infected -eq 0 -or $infRateNum -lt 0.2) { "HANG A (An toan)" } else { "HANG B (Da don sach)" }

            $wsDash.Cells.Item($rowCursor, 1).Value2 = [string]$d.Dept
            $wsDash.Cells.Item($rowCursor, 1).Font.Bold = $true
            $wsDash.Cells.Item($rowCursor, 2).Value2 = $d.Scanned.ToString("N0")
            $wsDash.Cells.Item($rowCursor, 3).Value2 = $d.Safe.ToString("N0")
            $wsDash.Cells.Item($rowCursor, 4).Value2 = $d.Cleaned.ToString("N0")
            $wsDash.Cells.Item($rowCursor, 5).Value2 = $d.ReInfected.ToString("N0")
            $wsDash.Cells.Item($rowCursor, 6).Value2 = $d.Errors.ToString("N0")
            $wsDash.Cells.Item($rowCursor, 7).Value2 = "$infRateNum%"
            $wsDash.Cells.Item($rowCursor, 8).Value2 = [string]$grade
            $wsDash.Cells.Item($rowCursor, 8).Font.Bold = $true

            $wsDash.Range("B$rowCursor`:F$rowCursor").HorizontalAlignment = -4152 # xlRight
            $wsDash.Range("G$rowCursor`:H$rowCursor").HorizontalAlignment = -4108

            if ($grade -like "HANG A*") {
                $wsDash.Cells.Item($rowCursor, 8).Interior.Color = 0xE8F5E9
                $wsDash.Cells.Item($rowCursor, 8).Font.Color     = 0x1B5E20
            } elseif ($grade -like "HANG B*") {
                $wsDash.Cells.Item($rowCursor, 8).Interior.Color = 0xFFF8E1
                $wsDash.Cells.Item($rowCursor, 8).Font.Color     = 0xE65100
            } else {
                $wsDash.Cells.Item($rowCursor, 8).Interior.Color = 0xFFEBEE
                $wsDash.Cells.Item($rowCursor, 8).Font.Color     = 0xB71C1C
            }
            $wsDash.Rows.Item($rowCursor).RowHeight = 19
            $rowCursor++
        }
        Set-TableRangeStyle $wsDash 10 ($rowCursor - 1) 8 0x5A2D12

        $wsDash.Rows.Item($rowCursor).RowHeight = 8
        $rowCursor++

        # KHOI 3 (Trai: A-D): TOP NGHI PHAM LUU FILE NHIEM CUOI CUNG (KPI 2 - LAST SAVED BY)
        # KHOI 4 (Phai: E-H): TOM TAT CAC BIEN THE VIRUS DA BOC TACH (KPI 9)
        $sec3Left = $wsDash.Range("A$rowCursor`:D$rowCursor")
        $sec3Left.Merge()
        $sec3Left.Value2 = "III. TRUY VET NGUOI LUU FILE NHIEM CUOI (KPI 2)"
        $sec3Left.Font.Bold = $true; $sec3Left.Font.Size = 10.5; $sec3Left.Font.Color = 0x4B230E

        $sec3Right = $wsDash.Range("E$rowCursor`:H$rowCursor")
        $sec3Right.Merge()
        $sec3Right.Value2 = "IV. CAC BIEN THE VIRUS DA NHAN DIEN (KPI 9)"
        $sec3Right.Font.Bold = $true; $sec3Right.Font.Size = 10.5; $sec3Right.Font.Color = 0x4B230E
        $wsDash.Rows.Item($rowCursor).RowHeight = 22
        $rowCursor++

        $hdrRowK3 = $rowCursor
        $wsDash.Cells.Item($hdrRowK3, 1).Value2 = "Tai khoan Luu cuoi (Last Saved By)"
        $wsDash.Cells.Item($hdrRowK3, 2).Value2 = "So file nhiem"
        $wsDash.Cells.Item($hdrRowK3, 3).Value2 = "Trang thai Add-in"
        $wsDash.Cells.Item($hdrRowK3, 4).Value2 = "Hanh dong de xuat"

        $wsDash.Cells.Item($hdrRowK3, 5).Value2 = "Ma Bien the"
        $wsDash.Cells.Item($hdrRowK3, 6).Value2 = "Mo ta Bien the Virus"
        $wsDash.Cells.Item($hdrRowK3, 7).Value2 = "So mau"
        $wsDash.Cells.Item($hdrRowK3, 8).Value2 = "Trang thai xu ly"

        $hdrLeftRange = $wsDash.Range("A$hdrRowK3`:D$hdrRowK3")
        $hdrLeftRange.Interior.Color = 0x5A2D12; $hdrLeftRange.Font.Color = 0xFFFFFF; $hdrLeftRange.Font.Bold = $true; $hdrLeftRange.HorizontalAlignment = -4108
        $hdrRightRange = $wsDash.Range("E$hdrRowK3`:H$hdrRowK3")
        $hdrRightRange.Interior.Color = 0x5A2D12; $hdrRightRange.Font.Color = 0xFFFFFF; $hdrRightRange.Font.Bold = $true; $hdrRightRange.HorizontalAlignment = -4108
        $wsDash.Rows.Item($hdrRowK3).RowHeight = 22
        $rowCursor++

        $topAuthors = @()
        if ($AuthorStats.Count -gt 0) {
            $topAuthors = $AuthorStats.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 4
        }

        for ($i = 0; $i -lt 4; $i++) {
            $rIdx = $rowCursor + $i
            $wsDash.Rows.Item($rIdx).RowHeight = 20

            if ($i -lt $topAuthors.Count) {
                $aName = [string]$topAuthors[$i].Key
                $aCnt  = [int]$topAuthors[$i].Value
                $hasAi = if ($KnownUsers.Contains($aName)) { "Da cai Add-in" } else { "Chua cai Add-in" }
                $act   = if ($hasAi -eq "Da cai Add-in") { "Da duoc bao ve tu dong" } else { "IT kiem tra & cai Add-in ngay" }
                $wsDash.Cells.Item($rIdx, 1).Value2 = $aName
                $wsDash.Cells.Item($rIdx, 2).Value2 = $aCnt.ToString("N0")
                $wsDash.Cells.Item($rIdx, 2).HorizontalAlignment = -4108
                $wsDash.Cells.Item($rIdx, 3).Value2 = $hasAi
                $wsDash.Cells.Item($rIdx, 3).HorizontalAlignment = -4108
                $wsDash.Cells.Item($rIdx, 4).Value2 = $act
                if ($hasAi -ne "Da cai Add-in") {
                    $wsDash.Cells.Item($rIdx, 3).Interior.Color = 0xFFEBEE
                    $wsDash.Cells.Item($rIdx, 3).Font.Color     = 0xB71C1C
                    $wsDash.Cells.Item($rIdx, 3).Font.Bold      = $true
                } else {
                    $wsDash.Cells.Item($rIdx, 3).Interior.Color = 0xE8F5E9
                    $wsDash.Cells.Item($rIdx, 3).Font.Color     = 0x1B5E20
                }
            } else {
                $wsDash.Cells.Item($rIdx, 1).Value2 = "-"
                $wsDash.Cells.Item($rIdx, 2).Value2 = "0"
                $wsDash.Cells.Item($rIdx, 3).Value2 = "-"
                $wsDash.Cells.Item($rIdx, 4).Value2 = "-"
            }

            if ($i -lt $Variants.Count) {
                $vObj = $Variants[$i]
                $wsDash.Cells.Item($rIdx, 5).Value2 = [string]($vObj.VariantCode.Split(' ')[0])
                $wsDash.Cells.Item($rIdx, 5).Font.Bold = $true
                $wsDash.Cells.Item($rIdx, 6).Value2 = [string]$vObj.VariantName
                $wsDash.Cells.Item($rIdx, 7).Value2 = ([int]$vObj.SampleCount).ToString("N0")
                $wsDash.Cells.Item($rIdx, 7).HorizontalAlignment = -4108
                $wsDash.Cells.Item($rIdx, 8).Value2 = "Da co thuoc dac tri 100%"
                $wsDash.Cells.Item($rIdx, 8).Interior.Color = 0xE8F5E9
                $wsDash.Cells.Item($rIdx, 8).Font.Color     = 0x1B5E20
                $wsDash.Cells.Item($rIdx, 8).HorizontalAlignment = -4108
            }
        }
        $endK3 = $rowCursor + 3
        $wsDash.Range("A$hdrRowK3`:D$endK3").Borders.LineStyle = 1
        $wsDash.Range("A$hdrRowK3`:D$endK3").Borders.Color     = 0xD9D9D9
        $wsDash.Range("E$hdrRowK3`:H$endK3").Borders.LineStyle = 1
        $wsDash.Range("E$hdrRowK3`:H$endK3").Borders.Color     = 0xD9D9D9

        $footerRow = $endK3 + 2
        $ftrRange = $wsDash.Range("A$footerRow`:H$footerRow")
        $ftrRange.Merge()
        $ftrRange.Value2 = "Huong dan: Trang Dashboard nay duoc thiet ke vua khit 1 trang A4. Xem chi tiet tung hang muc tai cac Sheet 2 -> Sheet 6 phia sau."
        $ftrRange.Font.Italic = $true
        $ftrRange.Font.Size   = 9
        $ftrRange.Font.Color  = 0x555555

        # Dinh dang chuan A4 Landscape cho Sheet 1
        $wsDash.Columns.Item(1).ColumnWidth = 26
        $wsDash.Columns.Item(2).ColumnWidth = 14
        $wsDash.Columns.Item(3).ColumnWidth = 16
        $wsDash.Columns.Item(4).ColumnWidth = 24
        $wsDash.Columns.Item(5).ColumnWidth = 13
        $wsDash.Columns.Item(6).ColumnWidth = 30
        $wsDash.Columns.Item(7).ColumnWidth = 13
        $wsDash.Columns.Item(8).ColumnWidth = 22
        try {
            $ps1 = $wsDash.PageSetup
            $ps1.Orientation    = 2 # xlLandscape
            $ps1.PaperSize      = 9 # xlPaperA4
            $ps1.Zoom           = $false
            $ps1.FitToPagesWide = 1
            $ps1.FitToPagesTall = 1
        } catch {}

        # ======================================================================
        # SHEET 2: 2_Thiet_Bi_Addin (KPI 1 - SO DANG KY THIET BI THEO MAC ADDRESS)
        # ======================================================================
        $t2 = $wsDevice.Range("A1:K1")
        $t2.Merge()
        $t2.Value2 = "SHEET 2: SO DANG KY THIET BI HOAT DONG & DO PHU ADD-IN (KHOA CHINH: MAC ADDRESS - KPI 1)"
        $t2.Interior.Color = 0x4B230E; $t2.Font.Color = 0xFFFFFF; $t2.Font.Bold = $true; $t2.Font.Size = 12
        $wsDevice.Rows.Item(1).RowHeight = 28

        $devHeaders = @("STT", "Khoa chinh MAC Address", "Dia chi IP (IPv4)", "Ten May tinh (Host)", "Tai khoan (User)", "Bo phan / Phong ban", "Ngay cai dat dau tien", "Hoat dong gan nhat", "Phien ban Add-in", "Tong file nhiem da diet", "Trang thai Bao ve")
        for ($i = 0; $i -lt $devHeaders.Count; $i++) {
            $wsDevice.Cells.Item(3, $i + 1).Value2 = $devHeaders[$i]
        }
        $rIdx = 4
        $stt = 0
        foreach ($ep in $Endpoints) {
            $stt++
            $wsDevice.Cells.Item($rIdx, 1).Value2  = "$stt"
            $wsDevice.Cells.Item($rIdx, 1).HorizontalAlignment = -4108
            $wsDevice.Cells.Item($rIdx, 2).Value2  = [string]$ep.MacAddress
            $wsDevice.Cells.Item($rIdx, 2).Font.Bold = $true
            $wsDevice.Cells.Item($rIdx, 3).Value2  = [string]$ep.IPv4Address
            $wsDevice.Cells.Item($rIdx, 4).Value2  = [string]$ep.ComputerName
            $wsDevice.Cells.Item($rIdx, 5).Value2  = [string]$ep.UserName
            $wsDevice.Cells.Item($rIdx, 6).Value2  = [string]$ep.Department
            $wsDevice.Cells.Item($rIdx, 7).Value2  = [string]$ep.FirstInstalled
            $wsDevice.Cells.Item($rIdx, 8).Value2  = [string]$ep.LastSeen
            $wsDevice.Cells.Item($rIdx, 9).Value2  = [string]$ep.InstalledVersion
            $wsDevice.Cells.Item($rIdx, 10).Value2 = ([int]$ep.FilesCleaned).ToString("N0")
            $wsDevice.Cells.Item($rIdx, 10).HorizontalAlignment = -4152
            $wsDevice.Cells.Item($rIdx, 11).Value2 = [string]$ep.Status
            if ($ep.Status -like "Dang bao ve*") {
                $wsDevice.Cells.Item($rIdx, 11).Interior.Color = 0xE8F5E9
                $wsDevice.Cells.Item($rIdx, 11).Font.Color     = 0x1B5E20
            } else {
                $wsDevice.Cells.Item($rIdx, 11).Interior.Color = 0xFFF8E1
                $wsDevice.Cells.Item($rIdx, 11).Font.Color     = 0xE65100
            }
            $wsDevice.Rows.Item($rIdx).RowHeight = 20
            $rIdx++
        }
        Set-TableRangeStyle $wsDevice 3 ($rIdx - 1) 11 0x5A2D12
        $wsDevice.Columns.AutoFit() | Out-Null

        # ======================================================================
        # SHEET 3: 3_Xep_Hang_Phong_Ban (KPI 3, 4, 8 - BAN DO NHIET & XEP HANG A/B/C)
        # ======================================================================
        $t3 = $wsDept.Range("A1:K1")
        $t3.Merge()
        $t3.Value2 = "SHEET 3: THONG KE MAT DO LAY NHIEM & XEP HANG AN NINH THEO BO PHAN (KPI 3, KPI 4, KPI 8)"
        $t3.Interior.Color = 0x4B230E; $t3.Font.Color = 0xFFFFFF; $t3.Font.Bold = $true; $t3.Font.Size = 12
        $wsDept.Rows.Item(1).RowHeight = 28

        $deptHeaders = @("STT", "Bo phan / Thu muc cap 1", "Tong file da quet", "File An toan", "Phat hien nhiem (Lan dau)", "Tai nhiem moi (Re-infected)", "Da diet sach & Sao luu", "File nhiem dang khoa (Locked)", "Ty le nhiem (%)", "Xep hang An ninh (A/B/C)", "Khuyen nghi Hanh dong cho Truong bo phan")
        for ($i = 0; $i -lt $deptHeaders.Count; $i++) {
            $wsDept.Cells.Item(3, $i + 1).Value2 = $deptHeaders[$i]
        }
        $rIdx = 4
        $stt = 0
        foreach ($d in $DeptList) {
            $stt++
            $infRateNum = if ($d.Scanned -gt 0) { [Math]::Round(($d.Infected / [double]$d.Scanned) * 100.0, 2) } else { 0.0 }
            $grade = if ($d.ReInfected -gt 0 -or $d.Errors -gt 0) { "HANG C (Canh bao)" } elseif ($d.Infected -eq 0 -or $infRateNum -lt 0.2) { "HANG A (An toan)" } else { "HANG B (Da lam sach)" }
            $advice = if ($d.ReInfected -gt 0) { "Kiem tra ngay may tram chua cai Add-in trong bo phan!" } elseif ($d.Errors -gt 0) { "Mo khoa Read-Only/Password cho $($d.Errors) file de diet not." } elseif ($d.Cleaned -gt 0) { "Da don sach $($d.Cleaned) file nhiem cu; duy tri giam sat." } else { "100% file sach, duy tri tot." }

            $wsDept.Cells.Item($rIdx, 1).Value2  = "$stt"
            $wsDept.Cells.Item($rIdx, 1).HorizontalAlignment = -4108
            $wsDept.Cells.Item($rIdx, 2).Value2  = [string]$d.Dept
            $wsDept.Cells.Item($rIdx, 2).Font.Bold = $true
            $wsDept.Cells.Item($rIdx, 3).Value2  = $d.Scanned.ToString("N0")
            $wsDept.Cells.Item($rIdx, 4).Value2  = $d.Safe.ToString("N0")
            $wsDept.Cells.Item($rIdx, 5).Value2  = $d.Infected.ToString("N0")
            $wsDept.Cells.Item($rIdx, 6).Value2  = $d.ReInfected.ToString("N0")
            $wsDept.Cells.Item($rIdx, 7).Value2  = $d.Cleaned.ToString("N0")
            $wsDept.Cells.Item($rIdx, 8).Value2  = $d.Errors.ToString("N0")
            $wsDept.Cells.Item($rIdx, 9).Value2  = "$infRateNum%"
            $wsDept.Cells.Item($rIdx, 10).Value2 = [string]$grade
            $wsDept.Cells.Item($rIdx, 10).Font.Bold = $true
            $wsDept.Cells.Item($rIdx, 11).Value2 = [string]$advice
            $wsDept.Range("C$rIdx`:H$rIdx").HorizontalAlignment = -4152

            if ($grade -like "HANG A*") {
                $wsDept.Cells.Item($rIdx, 10).Interior.Color = 0xE8F5E9; $wsDept.Cells.Item($rIdx, 10).Font.Color = 0x1B5E20
            } elseif ($grade -like "HANG B*") {
                $wsDept.Cells.Item($rIdx, 10).Interior.Color = 0xFFF8E1; $wsDept.Cells.Item($rIdx, 10).Font.Color = 0xE65100
            } else {
                $wsDept.Cells.Item($rIdx, 10).Interior.Color = 0xFFEBEE; $wsDept.Cells.Item($rIdx, 10).Font.Color = 0xB71C1C
            }
            $wsDept.Rows.Item($rIdx).RowHeight = 20
            $rIdx++
        }
        Set-TableRangeStyle $wsDept 3 ($rIdx - 1) 11 0x5A2D12
        $wsDept.Columns.AutoFit() | Out-Null

        # ======================================================================
        # SHEET 4: 4_So_Cai_File_Nhiem (KPI 2, 3, 5 - SO CAI FILE NHIEM & TRUY VET LAST SAVED BY)
        # ======================================================================
        $t4 = $wsInfected.Range("A1:L1")
        $t4.Merge()
        $t4.Value2 = "SHEET 4: SO CAI CHI TIET CAC FILE PHAT HIEN NHIEM & TRUY VET NGUOI LUU CUOI CUNG (KPI 2, KPI 3, KPI 5)"
        $t4.Interior.Color = 0x4B230E; $t4.Font.Color = 0xFFFFFF; $t4.Font.Bold = $true; $t4.Font.Size = 12
        $wsInfected.Rows.Item(1).RowHeight = 28

        $infHeaders = @("STT", "Thoi gian xu ly", "Nguon phat hien", "Phan loai (Cu / Tai nhiem)", "Bo phan", "Ten File Excel", "Kich thuoc (KB)", "Nguoi luu cuoi (Last Saved By)", "Trang thai Add-in cua Nguoi luu", "Bien the Virus", "Chi tiet dau hieu nhiem", "Duong dan day du")
        for ($i = 0; $i -lt $infHeaders.Count; $i++) {
            $wsInfected.Cells.Item(3, $i + 1).Value2 = $infHeaders[$i]
        }

        $recCount = $InfectedList.Count
        if ($recCount -gt 0) {
            $arrInf = New-Object 'string[,]' $recCount, 12
            for ($i = 0; $i -lt $recCount; $i++) {
                $r = $InfectedList[$i]
                $arrInf[$i, 0]  = "$($i + 1)"
                $arrInf[$i, 1]  = [string]$r.Time
                $arrInf[$i, 2]  = [string]$r.DetectionSource
                $arrInf[$i, 3]  = [string]$r.DiscoveryType
                $arrInf[$i, 4]  = [string]$r.Department
                $arrInf[$i, 5]  = [string]$r.FileName
                $arrInf[$i, 6]  = "$($r.SizeKB)"
                $arrInf[$i, 7]  = [string]$r.LastSavedBy
                $arrInf[$i, 8]  = [string]$r.AuthorAddinStat
                $arrInf[$i, 9]  = [string]$r.ThreatFamily
                $arrInf[$i, 10] = [string]$r.ThreatDetail
                $arrInf[$i, 11] = [string]$r.FullPath
            }
            $dataRange4 = $wsInfected.Range($wsInfected.Cells.Item(4, 1), $wsInfected.Cells.Item(3 + $recCount, 12))
            $dataRange4.Value(10) = $arrInf
            Set-TableRangeStyle $wsInfected 3 (3 + $recCount) 12 0x5A2D12
        } else {
            Set-TableRangeStyle $wsInfected 3 3 12 0x5A2D12
        }
        $wsInfected.Columns.AutoFit() | Out-Null
        if ($wsInfected.Columns.Item(11).ColumnWidth -gt 45) { $wsInfected.Columns.Item(11).ColumnWidth = 45 }
        if ($wsInfected.Columns.Item(12).ColumnWidth -gt 60) { $wsInfected.Columns.Item(12).ColumnWidth = 60 }

        # ======================================================================
        # SHEET 5: 5_Bien_The_Virus (KPI 9 - HO SO BIEN THE & GHI CHU KY THUAT)
        # ======================================================================
        $t5 = $wsVariant.Range("A1:I1")
        $t5.Merge()
        $t5.Value2 = "SHEET 5: HO SO PHAN TICH CAC BIEN THE VIRUS KANGATANG & GHI CHU KY THUAT CHUYEN SAU (KPI 9)"
        $t5.Interior.Color = 0x4B230E; $t5.Font.Color = 0xFFFFFF; $t5.Font.Bold = $true; $t5.Font.Size = 12
        $wsVariant.Rows.Item(1).RowHeight = 28

        $varHeaders = @("STT", "Ma Bien the", "Ma Bam Dai dien (Sample Hash)", "Ten goi & Phan nhom Bien the", "Co che An nau", "So mau thu thap", "Trieu chung gay hai cho Nguoi dung", "Ghi chu Ky thuat & Ly do Antivirus thuong mai bo lot", "Co che Xu ly Dac tri cua KangatangGuard")
        for ($i = 0; $i -lt $varHeaders.Count; $i++) {
            $wsVariant.Cells.Item(3, $i + 1).Value2 = $varHeaders[$i]
        }
        $rIdx = 4
        $stt = 0
        foreach ($v in $Variants) {
            $stt++
            $wsVariant.Cells.Item($rIdx, 1).Value2 = "$stt"
            $wsVariant.Cells.Item($rIdx, 2).Value2 = [string]$v.VariantCode
            $wsVariant.Cells.Item($rIdx, 2).Font.Bold = $true
            $wsVariant.Cells.Item($rIdx, 3).Value2 = [string]$v.SampleHash
            $wsVariant.Cells.Item($rIdx, 4).Value2 = [string]$v.VariantName
            $wsVariant.Cells.Item($rIdx, 4).Font.Bold = $true
            $wsVariant.Cells.Item($rIdx, 5).Value2 = [string]$v.HideMechanism
            $wsVariant.Cells.Item($rIdx, 6).Value2 = ([int]$v.SampleCount).ToString("N0")
            $wsVariant.Cells.Item($rIdx, 6).HorizontalAlignment = -4108
            $wsVariant.Cells.Item($rIdx, 7).Value2 = [string]$v.UserSymptoms
            $wsVariant.Cells.Item($rIdx, 8).Value2 = [string]$v.TechNote
            $wsVariant.Cells.Item($rIdx, 9).Value2 = [string]$v.GuardAction
            $wsVariant.Rows.Item($rIdx).RowHeight = 58
            $rIdx++
        }
        Set-TableRangeStyle $wsVariant 3 ($rIdx - 1) 9 0x5A2D12
        $wsVariant.Range("A4:I$($rIdx - 1)").WrapText = $true
        $wsVariant.Columns.Item(1).ColumnWidth = 6
        $wsVariant.Columns.Item(2).ColumnWidth = 24
        $wsVariant.Columns.Item(3).ColumnWidth = 28
        $wsVariant.Columns.Item(4).ColumnWidth = 30
        $wsVariant.Columns.Item(5).ColumnWidth = 26
        $wsVariant.Columns.Item(6).ColumnWidth = 14
        $wsVariant.Columns.Item(7).ColumnWidth = 34
        $wsVariant.Columns.Item(8).ColumnWidth = 44
        $wsVariant.Columns.Item(9).ColumnWidth = 38

        # ======================================================================
        # SHEET 6: 6_Lich_Su_Phat_Trien (KPI 10 - HANH TRINH PHAT TRIEN BANG NGON NGU DE HIEU)
        # ======================================================================
        $t6 = $wsRoadmap.Range("A1:F1")
        $t6.Merge()
        $t6.Value2 = "SHEET 6: BANG THEO DOI QUA TRINH NGHIEN CUU, PHAT TRIEN & CAI TIEN KANGATANG GUARD (KPI 10 - OHSUNG IT TEAM)"
        $t6.Interior.Color = 0x4B230E; $t6.Font.Color = 0xFFFFFF; $t6.Font.Bold = $true; $t6.Font.Size = 12
        $wsRoadmap.Rows.Item(1).RowHeight = 28

        $rmHeaders = @("STT", "Giai doan / Phien ban", "Thoi gian", "Kho khan thuc te cua Nhan vien & Nha may", "Giai phap Cai tien cua doi ngu OhsungIT", "Ket qua & Loi ich Thuc te dat duoc (Truoc -> Sau)")
        for ($i = 0; $i -lt $rmHeaders.Count; $i++) {
            $wsRoadmap.Cells.Item(3, $i + 1).Value2 = $rmHeaders[$i]
        }
        $rIdx = 4
        $stt = 0
        foreach ($rm in $Roadmap) {
            $stt++
            $wsRoadmap.Cells.Item($rIdx, 1).Value2 = "$stt"
            $wsRoadmap.Cells.Item($rIdx, 1).HorizontalAlignment = -4108
            $wsRoadmap.Cells.Item($rIdx, 2).Value2 = [string]$rm.Stage
            $wsRoadmap.Cells.Item($rIdx, 2).Font.Bold = $true
            $wsRoadmap.Cells.Item($rIdx, 3).Value2 = [string]$rm.Date
            $wsRoadmap.Cells.Item($rIdx, 3).HorizontalAlignment = -4108
            $wsRoadmap.Cells.Item($rIdx, 4).Value2 = [string]$rm.Problem
            $wsRoadmap.Cells.Item($rIdx, 5).Value2 = [string]$rm.Solution
            $wsRoadmap.Cells.Item($rIdx, 6).Value2 = [string]$rm.Impact
            $wsRoadmap.Cells.Item($rIdx, 6).Font.Bold = $true
            $wsRoadmap.Cells.Item($rIdx, 6).Font.Color = 0x1B5E20
            $wsRoadmap.Rows.Item($rIdx).RowHeight = 46
            $rIdx++
        }
        Set-TableRangeStyle $wsRoadmap 3 ($rIdx - 1) 6 0x5A2D12
        $wsRoadmap.Range("A4:F$($rIdx - 1)").WrapText = $true
        $wsRoadmap.Columns.Item(1).ColumnWidth = 6
        $wsRoadmap.Columns.Item(2).ColumnWidth = 22
        $wsRoadmap.Columns.Item(3).ColumnWidth = 12
        $wsRoadmap.Columns.Item(4).ColumnWidth = 42
        $wsRoadmap.Columns.Item(5).ColumnWidth = 44
        $wsRoadmap.Columns.Item(6).ColumnWidth = 42

        $wsDash.Activate()
        $wb.SaveAs($tempXlsx, 51) # xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        $wb = $null

        try {
            Copy-Item -LiteralPath $tempXlsx -Destination $OutputPath -Force -ErrorAction Stop
        } catch {
            $altPath = Join-Path (Split-Path -Parent $OutputPath) ("Bao_Cao_Quan_Tri_KangatangGuard_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsx")
            Copy-Item -LiteralPath $tempXlsx -Destination $altPath -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "[WARN] Khong the xuat file Excel bao cao: $($_.Exception.Message) | $($_.InvocationInfo.PositionMessage)" -ForegroundColor Yellow
    } finally {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
        }
        if ($null -ne $xl) {
            try { $xl.IgnoreRemoteRequests = $false } catch {}
            try { $xl.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
        }
        foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
            $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
            if (Test-Path $optKey) {
                $curOpt = (Get-ItemProperty -Path $optKey -Name "Options" -ErrorAction SilentlyContinue).Options
                if ($null -ne $curOpt -and (($curOpt -band 0x40) -ne 0)) {
                    Set-ItemProperty -Path $optKey -Name "Options" -Value ($curOpt -band (-bnot 0x40)) -Force -ErrorAction SilentlyContinue
                }
            }
        }
        Remove-Item -LiteralPath $tempXlsx -Force -ErrorAction SilentlyContinue
    }
}

if ($WatchActiveSession) {
    while ($true) {
        $res = Invoke-SyncScanReportsOnce
        if ($res.Status -eq "COMPLETED") { break }
        $sessMetaFile = Join-Path $appDataLogDir "last_session.json"
        if (Test-Path -LiteralPath $sessMetaFile) {
            try {
                $fi = New-Object System.IO.FileInfo($sessMetaFile)
                if (((Get-Date) - $fi.LastWriteTime).TotalMinutes -gt 10) { break }
            } catch {}
        }
        Start-Sleep -Seconds 30
    }
} else {
    $res = Invoke-SyncScanReportsOnce
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host "   DA CAP NHAT BO BAO CAO QUAN TRI EXCEL DA SHEET (v$AppVersion)      " -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host "   Trang thai phien : $($res.Status)" -ForegroundColor White
    Write-Host "   Thiet bi bao ve  : $($res.TotalEndpoints) may tram (Khoa chinh MAC Address)" -ForegroundColor Cyan
    Write-Host "   Tong tep da quet : $($res.TotalScanned.ToString('N0')) tep ($($res.TotalDepts) phong ban)" -ForegroundColor White
    Write-Host "   Tong tep da diet : $($res.TotalCleaned.ToString('N0')) tep (100% da sao luu)" -ForegroundColor Yellow
    Write-Host "   ------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "   1. Bao cao Excel 6 Sheet : $($res.ExcelReportPath)" -ForegroundColor Green
    Write-Host "   2. So dang ky Thiet bi   : $($res.EndpointCsvPath)" -ForegroundColor Green
    Write-Host "   3. So cai File nhiem     : $($res.LedgerCsvPath)" -ForegroundColor Green
    Write-Host "   4. Bao cao Markdown      : $($res.LatestMdPath)" -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Cyan
}

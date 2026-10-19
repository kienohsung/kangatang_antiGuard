# ==============================================================================
# Pull_And_Clean_From_223_7.ps1
# Dich vu Cau noi Dong bo Thoi gian thuc & Ho tro COM Worker tu xa (223.7 <-> 223.176)
# Phien ban: v3.21.4
# ==============================================================================

$ErrorActionPreference = "SilentlyContinue"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$utf8Bom   = New-Object System.Text.UTF8Encoding($true)

$remoteHubRoot   = "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang"
$remoteBridgeDir = Join-Path $remoteHubRoot "Sync_To_223_176"
$remoteStateDir  = Join-Path $remoteHubRoot "Shared_Resume_State"

$localProjectDir = "D:\7. AI tools\kangatang"
$localReportDir  = Join-Path $localProjectDir "Scan_Reports"
$localStateDir   = Join-Path $localReportDir "Shared_Resume_State"
$localQuarDir1   = "D:\test\Quarantine_Backup"
$localQuarDir2   = Join-Path $localProjectDir "Quarantine_Backup"
$localSampleDir  = Join-Path $localProjectDir "Threat_Samples"
$xlamPath        = Join-Path $localProjectDir "addin\KangatangGuard.xlam"

foreach ($d in @($localReportDir, (Join-Path $localReportDir "Sessions"), $localStateDir, $localQuarDir1, $localQuarDir2, $localSampleDir)) {
    if (-not (Test-Path -LiteralPath $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}

$script:fsoBridge = $null
function Get-BridgeShortPath([string]$p) {
    if ([string]::IsNullOrEmpty($p) -or $p.Length -lt 215) { return $p }
    try {
        if ($null -eq $script:fsoBridge) { $script:fsoBridge = New-Object -ComObject Scripting.FileSystemObject }
        $lp = if ($p.StartsWith("\\?\")) { $p } elseif ($p.StartsWith("\\")) { "\\?\UNC\" + $p.Substring(2) } else { "\\?\" + $p }
        $sh = $script:fsoBridge.GetFile($lp).ShortPath
        if ($sh) {
            if ($sh.StartsWith("\\?\UNC\")) { return "\\" + $sh.Substring(8) }
            if ($sh.StartsWith("\\?\")) { return $sh.Substring(4) }
            return $sh
        }
    } catch {}
    return $p
}

$script:excelWorker = $null
$script:wbWorker    = $null
$workerXlam = Join-Path $env:APPDATA "KangatangGuard\KangatangGuard_BridgeWorker.xlam"

function Ensure-BridgeExcelWorker {
    if (($null -ne $script:excelWorker) -and ($null -ne $script:wbWorker)) { return $true }
    try {
        Copy-Item -LiteralPath $xlamPath -Destination $workerXlam -Force -ErrorAction SilentlyContinue
        $script:excelWorker = New-Object -ComObject Excel.Application
        $script:excelWorker.Visible = $false
        try { $script:excelWorker.AutomationSecurity = 1 } catch {}
        $script:excelWorker.DisplayAlerts = $false
        $script:excelWorker.EnableEvents = $false
        $script:wbWorker = $script:excelWorker.Workbooks.Open($workerXlam)
        $script:excelWorker.Run("'$($script:wbWorker.Name)'!InitWorkerMode") | Out-Null
        return $true
    } catch {
        $script:excelWorker = $null
        $script:wbWorker    = $null
        return $false
    }
}

$processedReqIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$lastReqLineCount = 0
$lastLogSyncTime  = [datetime]::MinValue

Write-Output "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Bridge Daemon v3.23.3 started: Monitoring $remoteBridgeDir"

while ($true) {
    try {
        if (Test-Path -LiteralPath $remoteHubRoot) {
            # 1. Dong bo Shared_Resume_State (Server_Profile, last_session.json; scanned_files.log moi 60s de tranh nghen SMB)
            if (Test-Path -LiteralPath $remoteStateDir) {
                Get-ChildItem -LiteralPath $remoteStateDir -Filter "Server_Profile_*.json" -ErrorAction SilentlyContinue | ForEach-Object {
                    $dstProf = Join-Path $localReportDir $_.Name
                    if ((-not (Test-Path -LiteralPath $dstProf)) -or ($_.LastWriteTime -gt (Get-Item -LiteralPath $dstProf).LastWriteTime)) {
                        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $localStateDir $_.Name) -Force -ErrorAction SilentlyContinue
                        Copy-Item -LiteralPath $_.FullName -Destination $dstProf -Force -ErrorAction SilentlyContinue
                    }
                }
                $nowUtc = [datetime]::UtcNow
                if (($nowUtc - $lastLogSyncTime).TotalSeconds -ge 60) {
                    $lastLogSyncTime = $nowUtc
                    $remSess = Join-Path $remoteStateDir "last_session.json"
                    $locSess = Join-Path $localStateDir "last_session.json"
                    if (Test-Path -LiteralPath $remSess) {
                        Copy-Item -LiteralPath $remSess -Destination $locSess -Force -ErrorAction SilentlyContinue
                    }
                    $remLog = Join-Path $remoteStateDir "scanned_files.log"
                    $locLog = Join-Path $localStateDir "scanned_files.log"
                    if ((Test-Path -LiteralPath $remLog) -and (Test-Path -LiteralPath $locLog)) {
                        $rItem = Get-Item -LiteralPath $remLog -ErrorAction SilentlyContinue
                        $lItem = Get-Item -LiteralPath $locLog -ErrorAction SilentlyContinue
                        if ($rItem -and $lItem -and ($rItem.LastWriteTime -gt $lItem.LastWriteTime) -and ($rItem.Length -ne $lItem.Length)) {
                            Copy-Item -LiteralPath $remLog -Destination $locLog -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
            }

            # 2. Dong bo Bao cao & Log tu Sync_To_223_176\Scan_Reports ve D:\7. AI tools\kangatang\Scan_Reports
            $remRepDir = Join-Path $remoteBridgeDir "Scan_Reports"
            if (Test-Path -LiteralPath $remRepDir) {
                Get-ChildItem -LiteralPath $remRepDir -File -ErrorAction SilentlyContinue | ForEach-Object {
                    $dstFile = Join-Path $localReportDir $_.Name
                    if ((-not (Test-Path -LiteralPath $dstFile)) -or ($_.LastWriteTime -gt (Get-Item -LiteralPath $dstFile).LastWriteTime) -or ($_.Length -ne (Get-Item -LiteralPath $dstFile).Length)) {
                        Copy-Item -LiteralPath $_.FullName -Destination $dstFile -Force -ErrorAction SilentlyContinue
                    }
                }
                $remSessDir = Join-Path $remRepDir "Sessions"
                $locSessDir = Join-Path $localReportDir "Sessions"
                if (Test-Path -LiteralPath $remSessDir) {
                    Get-ChildItem -LiteralPath $remSessDir -File -ErrorAction SilentlyContinue | ForEach-Object {
                        $dstSess = Join-Path $locSessDir $_.Name
                        if ((-not (Test-Path -LiteralPath $dstSess)) -or ($_.LastWriteTime -gt (Get-Item -LiteralPath $dstSess).LastWriteTime)) {
                            Copy-Item -LiteralPath $_.FullName -Destination $dstSess -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
            }

            # 3. Rut & Xoa sach Quarantine_Backup va Threat_Samples tu 223.7 ve 223.176
            $remQuarDir = Join-Path $remoteBridgeDir "Quarantine_Backup"
            if (Test-Path -LiteralPath $remQuarDir) {
                Get-ChildItem -LiteralPath $remQuarDir -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    Copy-Item -LiteralPath $_.FullName -Destination $localQuarDir1 -Recurse -Force -ErrorAction SilentlyContinue
                    Copy-Item -LiteralPath $_.FullName -Destination $localQuarDir2 -Recurse -Force -ErrorAction SilentlyContinue
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            $remSampDir = Join-Path $remoteBridgeDir "Threat_Samples"
            if (Test-Path -LiteralPath $remSampDir) {
                Get-ChildItem -LiteralPath $remSampDir -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    Copy-Item -LiteralPath $_.FullName -Destination $localSampleDir -Recurse -Force -ErrorAction SilentlyContinue
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
            }

            # 4. Xu ly Hang doi Remote_COM_Queue (Chi doc cac dong moi tu $lastReqLineCount)
            $remQueueDir = Join-Path $remoteBridgeDir "Remote_COM_Queue"
            $reqFile = Join-Path $remQueueDir "remote_com_requests.jsonl"
            $resFile = Join-Path $remQueueDir "remote_com_results.jsonl"
            if (Test-Path -LiteralPath $reqFile) {
                $reqLines = [System.IO.File]::ReadAllLines($reqFile, $utf8NoBom)
                if ($reqLines.Count -lt $lastReqLineCount) { $lastReqLineCount = 0 }
                if ($reqLines.Count -gt $lastReqLineCount) {
                    $outLines = New-Object System.Collections.Generic.List[string]
                    for ($idxReq = $lastReqLineCount; $idxReq -lt $reqLines.Count; $idxReq++) {
                        $ln = $reqLines[$idxReq]
                        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                        $reqObj = $ln | ConvertFrom-Json
                        if ($reqObj -and $reqObj.ReqId -and $processedReqIds.Add([string]$reqObj.ReqId)) {
                            $uncTarget = Get-BridgeShortPath ([string]$reqObj.UncPath)
                            $resStr = "SAFE"
                            if (Ensure-BridgeExcelWorker) {
                                try {
                                    $resStr = [string]$script:excelWorker.Run("'$($script:wbWorker.Name)'!ScanSingleFileForWorker", $uncTarget)
                                    if ($resStr.StartsWith("PURGE_NEEDED|")) {
                                        $detail = $resStr.Substring(13)
                                        $purgeRes = "MORE_PURGE"
                                        for ($att = 1; $att -le 5; $att++) {
                                            try { $script:excelWorker.Run("'$($script:wbWorker.Name)'!WorkerDeleteOneModuleSheet", $uncTarget) | Out-Null } catch {}
                                            try {
                                                $purgeRes = [string]$script:excelWorker.Run("'$($script:wbWorker.Name)'!WorkerFinishCleanAndClose", $uncTarget)
                                            } catch {
                                                $purgeRes = "INFECTED_ERR|RPC error"
                                                $script:excelWorker = $null
                                                $script:wbWorker = $null
                                                break
                                            }
                                            if ($purgeRes -ne "MORE_PURGE") { break }
                                        }
                                        if ($purgeRes.StartsWith("CLEANED|")) {
                                            $resStr = "CLEANED_PURGE|$detail"
                                        } else {
                                            $resStr = $purgeRes
                                        }
                                    }
                                } catch {
                                    $resStr = "WORKER_RECOVERED|$($_.Exception.Message)"
                                    $script:excelWorker = $null
                                    $script:wbWorker = $null
                                }
                            }
                            $ansObj = [ordered]@{
                                ReqId    = [string]$reqObj.ReqId
                                Result   = $resStr
                                WorkerId = 1
                            }
                            $outLines.Add(($ansObj | ConvertTo-Json -Compress))
                        }
                    }
                    $lastReqLineCount = $reqLines.Count
                    if ($outLines.Count -gt 0) {
                        $chunk = ($outLines -join "`r`n") + "`r`n"
                        [System.IO.File]::AppendAllText($resFile, $chunk, $utf8NoBom)
                    }
                }
            }
        }
    } catch {}
    Start-Sleep -Seconds 3
}

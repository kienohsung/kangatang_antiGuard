# ==============================================================================
# Kangatang_FolderScanner.ps1
# Trinh Quet Thu Muc Sieu Toc 4 Tang & Bang Dieu Khien Truc Tuyen (Live Dashboard)
# Phien ban: v3.24.0 (Split Tier-2/COM Lanes, Streaming LOH-free Zero-COM, Excel PID Watchdog & XLSTART Vaccine)
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = "",
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    [Parameter(Mandatory=$false)]
    [int]$MaxWorkers = 0,          # So Tier-2 Workers (Zero-COM, khong dung Excel)
    [Parameter(Mandatory=$false)]
    [int]$ComWorkers = 0,          # So COM Workers (Excel COM, gioi han cung so EXCEL.EXE cua scanner)
    [Parameter(Mandatory=$false)]
    [int]$IoConcurrency = 0,       # So luong doc dia dong thoi toi da (0 = bang so Tier-2 Workers)
    [Parameter(Mandatory=$false)]
    [switch]$InteractivePrompt
)

try { [System.Runtime.GCSettings]::LatencyMode = [System.Runtime.GCLatencyMode]::SustainedLowLatency } catch {}

# Ham thao tac chuoi duong dan an toan tuyet doi (Khong bao gio nem PathTooLongException tren .NET 4.5 khi >= 260 ky tu)
function Join-ScannerPath([string]$parent, [string]$child) {
    if ([string]::IsNullOrEmpty($parent)) { return $child }
    if ($parent.EndsWith("\")) { return ($parent + $child) }
    return ($parent + "\" + $child)
}

function Normalize-ScannerPath([string]$p) {
    if ([string]::IsNullOrEmpty($p)) { return "" }
    $s = $p.Trim()
    if ($s.StartsWith("\\")) {
        return ("\\" + $s.Substring(2).Replace("\\", "\"))
    }
    while ($s.Contains("\\")) {
        $s = $s.Replace("\\", "\")
    }
    return $s
}

function Get-SafeParentDir([string]$p) {
    if ([string]::IsNullOrEmpty($p)) { return "" }
    $idx = $p.LastIndexOf('\')
    if ($idx -le 0) { return "" }
    if ($idx -eq 2 -and $p[1] -eq ':') { return $p.Substring(0, 3) }
    return $p.Substring(0, $idx)
}

function Get-SafeFileName([string]$p) {
    if ([string]::IsNullOrEmpty($p)) { return "" }
    $idx = $p.LastIndexOf('\')
    if ($idx -lt 0) { return $p }
    return $p.Substring($idx + 1)
}

function Get-SafeExtension([string]$p) {
    $fn = Get-SafeFileName $p
    if ([string]::IsNullOrEmpty($fn)) { return "" }
    $dot = $fn.LastIndexOf('.')
    if ($dot -lt 0) { return "" }
    return $fn.Substring($dot).ToLowerInvariant()
}

function ConvertTo-Win32LongPath([string]$p) {
    if ([string]::IsNullOrEmpty($p)) { return $p }
    if ($p.StartsWith("\\?\")) { return $p }
    if ($p.StartsWith("\\")) { return "\\?\UNC\" + $p.Substring(2) }
    return "\\?\" + $p
}

$script:fso = $null
function Get-SafeShortPath([string]$p, [switch]$IsDirectory) {
    if ([string]::IsNullOrEmpty($p) -or $p.Length -lt 240) { return $p }
    try {
        if ($null -eq $script:fso) { $script:fso = New-Object -ComObject Scripting.FileSystemObject }
        $lp = ConvertTo-Win32LongPath $p
        $short = if ($IsDirectory) { $script:fso.GetFolder($lp).ShortPath } else { $script:fso.GetFile($lp).ShortPath }
        if ($short) {
            if ($short.StartsWith("\\?\UNC\")) { return "\\" + $short.Substring(8) }
            if ($short.StartsWith("\\?\")) { return $short.Substring(4) }
            return $short
        }
    } catch {}
    return $p
}

# 0. Quan ly Da Cua So Song Song (Multi-Window Concurrency & Per-Folder Mutex + Orphan Worker Cleanup)
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$scanReqFile     = Join-Path $logDir "scan_request.json"
$lastSessionFile = Join-Path $logDir "last_session.json"
$scannedLogFile  = Join-Path $logDir "scanned_files.log"

$script:scannerMutex = $null
$script:hasMutex     = $false

function Get-OtherScannerPids {
    $pids = New-Object 'System.Collections.Generic.List[int]'
    try {
        $procs = Get-WmiObject -Class Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue
        foreach ($pr in $procs) {
            $pIdVal = [int]$pr.ProcessId
            $cmdLine = [string]$pr.CommandLine
            if (($pIdVal -ne $PID) -and ($cmdLine -match "Kangatang_FolderScanner")) {
                $pids.Add($pIdVal)
            }
        }
    } catch {}
    return $pids
}

function Invoke-WithLogMutex([scriptblock]$Action) {
    $logMutex = $null
    $locked   = $false
    try {
        $logMutex = New-Object System.Threading.Mutex($false, "Local\KangatangGuard_LogWrite_Mutex")
        try {
            $locked = $logMutex.WaitOne(3000, $false)
        } catch [System.Threading.AbandonedMutexException] {
            $locked = $true
        }
        & $Action
    } catch {} finally {
        if ($locked -and ($null -ne $logMutex)) {
            try { $logMutex.ReleaseMutex() } catch {}
        }
        if ($null -ne $logMutex) {
            try { $logMutex.Dispose() } catch {}
        }
    }
}

# Chi don dep EXCEL.EXE chay ngam mo coi khi KHONG CO cua so Kangatang_FolderScanner nao khac dang chay!
try {
    $otherPidsAtStart = @(Get-OtherScannerPids)
    $alivePidSet = New-Object 'System.Collections.Generic.HashSet[int]'
    [void]$alivePidSet.Add($PID)
    foreach ($op in $otherPidsAtStart) { [void]$alivePidSet.Add([int]$op) }

    Get-ChildItem -LiteralPath $logDir -Filter "KangatangGuard_Worker_*.xlam" -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Name -match '^KangatangGuard_Worker_(\d+)_\d+(_\d+)?\.xlam$') {
            $wOwnerPid = [int]$Matches[1]
            if (-not $alivePidSet.Contains($wOwnerPid)) {
                try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch {}
            }
        } elseif ($otherPidsAtStart.Count -eq 0) {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    if ($otherPidsAtStart.Count -eq 0) {
        Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
        }
        # v3.24.0: Don EXCEL.EXE tu dong hoa (/automation -Embedding) mo coi trong CUNG phien dang nhap
        # (Excel nguoi dung mo bang tay khong co tham so nay -> khong bao gio bi dong nham)
        try {
            $mySession = (Get-Process -Id $PID).SessionId
            Get-WmiObject -Class Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue | ForEach-Object {
                $cl = [string]$_.CommandLine
                if ($cl -match '(?i)/automation|-embedding') {
                    $pr = Get-Process -Id ([int]$_.ProcessId) -ErrorAction SilentlyContinue
                    if ($pr -and $pr.SessionId -eq $mySession) {
                        try { Stop-Process -Id $pr.Id -Force -ErrorAction SilentlyContinue } catch {}
                    }
                }
            }
        } catch {}
        Start-Sleep -Milliseconds 200
    }
} catch {}

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
$AppVersion = "3.24.0"
$verJsonPath = Join-Path $ScriptDir "version.json"
if (Test-Path $verJsonPath) {
    try {
        $vMeta = Get-Content -Path $verJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($vMeta.version -and ($vMeta.version -gt $AppVersion)) { $AppVersion = $vMeta.version }
    } catch {}
}

$Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion [PID:$PID] - Bang Dieu Khien Quet Thu Muc (Multi-Window)"
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

$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$addinsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path $addinsDir)) { try { New-Item -ItemType Directory -Path $addinsDir -Force | Out-Null } catch {} }

# 2. Dam bao AccessVBOM = 1, VBAWarnings = 1, Trusted Location cho Worker & tat hop thoai WER
try {
    $werKey = "HKCU:\Software\Microsoft\Windows\Windows Error Reporting"
    if (-not (Test-Path $werKey)) { New-Item -Path $werKey -Force | Out-Null }
    Set-ItemProperty -Path $werKey -Name "DontShowUI" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
} catch {}

foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
    $excelBaseKey = "HKCU:\Software\Microsoft\Office\$ver\Excel"
    $secPath = "$excelBaseKey\Security"
    if (-not (Test-Path $secPath)) {
        try { New-Item -Path $secPath -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    if (Test-Path $secPath) {
        Set-ItemProperty -Path $secPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $secPath -Name "VBAWarnings" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        try {
            $tlKey = "$secPath\Trusted Locations"
            if (-not (Test-Path $tlKey)) { New-Item -Path $tlKey -Force -ErrorAction SilentlyContinue | Out-Null }
            Set-ItemProperty -Path $tlKey -Name "AllowNetworkLocations" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            $locKg = "$tlKey\KangatangGuardWorker"
            if (-not (Test-Path $locKg)) { New-Item -Path $locKg -Force -ErrorAction SilentlyContinue | Out-Null }
            Set-ItemProperty -Path $locKg -Name "Path" -Value $logDir -Type String -Force -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $locKg -Name "AllowSubfolders" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            $locAddins = "$tlKey\KangatangGuardAddIns"
            if (-not (Test-Path $locAddins)) { New-Item -Path $locAddins -Force -ErrorAction SilentlyContinue | Out-Null }
            Set-ItemProperty -Path $locAddins -Name "Path" -Value $addinsDir -Type String -Force -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $locAddins -Name "AllowSubfolders" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        } catch {}
    }
    foreach ($sub in @("DocumentRecovery", "StartupItems")) {
        $resKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\$sub"
        if (Test-Path $resKey) {
            Remove-Item -Path $resKey -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# 3. Doc yeu cau quet tu scan_request.json & Dong bo Shared_Resume_State + Sync_To_223_176
$scanReqFile     = Join-Path $logDir "scan_request.json"
$lastSessionFile = Join-Path $logDir "last_session.json"
$scannedLogFile  = Join-Path $logDir "scanned_files.log"

$sharedStateGroups = @(
    @("D:\file_shared\vietnam\z. ETC\1. addinKangatang\Shared_Resume_State", (Join-Path $ScriptDir "Shared_Resume_State"), "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\Shared_Resume_State"),
    @("D:\7. AI tools\kangatang\Scan_Reports\Shared_Resume_State", "\\192.168.223.176\test\Scan_Reports\Shared_Resume_State")
)
$script:sharedResumeDirs = New-Object System.Collections.Generic.List[string]
foreach ($grp in $sharedStateGroups) {
    foreach ($sc in $grp) {
        if (-not $sc) { continue }
        $parentSc = Get-SafeParentDir $sc
        if ($parentSc -and (Test-Path -LiteralPath $parentSc)) {
            if (-not (Test-Path -LiteralPath $sc)) {
                try { New-Item -ItemType Directory -Path $sc -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
            }
            if (Test-Path -LiteralPath $sc) {
                if (-not $script:sharedResumeDirs.Contains($sc)) {
                    $script:sharedResumeDirs.Add($sc)
                }
                break # Chi lay 1 duong dan tot nhat trong moi nhom (Uu tien o dia Local truoc UNC de tranh ghi lap 2 lan tren 223.7)
            }
        }
    }
}

# Thiet lap thu muc cau noi Sync_To_223_176 (Uu tien thu muc chia se D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176 de 223.176 keo ve ngay)
$bridgeCandidates = @(
    "D:\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176",
    "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\Sync_To_223_176",
    (Join-Path $ScriptDir "Sync_To_223_176")
)
$script:bridgeRootDir = ""
foreach ($bc in $bridgeCandidates) {
    if (-not $bc) { continue }
    $parentBc = Get-SafeParentDir $bc
    if ($parentBc -and (Test-Path -LiteralPath $parentBc)) {
        if (-not (Test-Path -LiteralPath $bc)) {
            try { New-Item -ItemType Directory -Path $bc -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        }
        if (Test-Path -LiteralPath $bc) {
            $script:bridgeRootDir = $bc
            break
        }
    }
}
if ($script:bridgeRootDir) {
    foreach ($subB in @("Scan_Reports", "Scan_Reports\Sessions", "Quarantine_Backup", "Threat_Samples", "Remote_COM_Queue")) {
        $pSub = Join-Path $script:bridgeRootDir $subB
        if (-not (Test-Path -LiteralPath $pSub)) {
            try { New-Item -ItemType Directory -Path $pSub -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        }
    }
}

# Neu may hien tai (vd 223.7) chua co scanned_files.log hoac ban tren Shared_Resume_State moi/day du hon -> ke thua ngay!
foreach ($sDir in $script:sharedResumeDirs) {
    $sharedLog  = Join-Path $sDir "scanned_files.log"
    $sharedSess = Join-Path $sDir "last_session.json"
    try {
        if (Test-Path -LiteralPath $sharedLog) {
            $shLen  = (Get-Item -LiteralPath $sharedLog -ErrorAction SilentlyContinue).Length
            $locLen = if (Test-Path -LiteralPath $scannedLogFile) { (Get-Item -LiteralPath $scannedLogFile -ErrorAction SilentlyContinue).Length } else { 0 }
            if ($shLen -gt $locLen) {
                Copy-Item -LiteralPath $sharedLog -Destination $scannedLogFile -Force -ErrorAction SilentlyContinue
            }
        }
        if ((Test-Path -LiteralPath $sharedSess) -and (-not (Test-Path -LiteralPath $lastSessionFile))) {
            Copy-Item -LiteralPath $sharedSess -Destination $lastSessionFile -Force -ErrorAction SilentlyContinue
        } elseif ((Test-Path -LiteralPath $sharedSess) -and (Test-Path -LiteralPath $lastSessionFile)) {
            $shMeta  = Get-Content -LiteralPath $sharedSess -Raw -Encoding UTF8 | ConvertFrom-Json
            $locMeta = Get-Content -LiteralPath $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($shMeta -and $locMeta -and (([int]$shMeta.TotalScanned -gt [int]$locMeta.TotalScanned) -or ([int]$shMeta.TotalFiles -gt [int]$locMeta.TotalFiles))) {
                Copy-Item -LiteralPath $sharedSess -Destination $lastSessionFile -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {}
}

$isResume = [bool]$Resume
if ((-not $InteractivePrompt) -and (-not $TargetFolder) -and (Test-Path $scanReqFile)) {
    try {
        $reqMeta = Get-Content -Path $scanReqFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($reqMeta.TargetFolder) { $TargetFolder = [string]$reqMeta.TargetFolder }
        if ($reqMeta.Resume -eq $true) { $isResume = $true }
    } catch {}
    Remove-Item -Path $scanReqFile -Force -ErrorAction SilentlyContinue
}

# Ham chuan hoa duong dan tuong doi \vietnam\... giua Local Disk (tren 223.7) va UNC (\\192.168.223.7\file_shared\vietnam)
function Get-VietnamRootPrefix([string]$pathStr) {
    if ([string]::IsNullOrWhiteSpace($pathStr)) { return "" }
    if ($pathStr -match '(?i)^(.*?\\vietnam)(?:\\.*)?$') {
        return $Matches[1]
    }
    return ""
}

function ConvertTo-TargetVietnamPath([string]$rawPath, [string]$localVnRoot) {
    if ([string]::IsNullOrWhiteSpace($rawPath)) { return "" }
    $normPath = Normalize-ScannerPath $rawPath
    if ([string]::IsNullOrWhiteSpace($localVnRoot)) { return $normPath }
    if ($normPath -match '(?i)^.*?\\vietnam(\\.+)$') {
        return (Normalize-ScannerPath ($localVnRoot + $Matches[1]))
    }
    if ($normPath -match '(?i)^.*?\\vietnam$') {
        return (Normalize-ScannerPath $localVnRoot)
    }
    return $normPath
}

if ($InteractivePrompt -or (-not $TargetFolder)) {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    # Tu dong tim duong dan O CUNG NOI BO cua file_shared\vietnam neu dang chay truc tiep tren may chu 223.7
    $defaultFolder = ""
    try {
        $smbShare = Get-SmbShare -Name "file_shared" -ErrorAction SilentlyContinue
        if ($smbShare -and $smbShare.Path) {
            $candVn = Join-Path $smbShare.Path "vietnam"
            if (Test-Path -LiteralPath $candVn) { $defaultFolder = $candVn }
            elseif (Test-Path -LiteralPath $smbShare.Path) { $defaultFolder = $smbShare.Path }
        }
    } catch {}
    if (-not $defaultFolder) {
        foreach ($drv in @("D:", "E:", "F:", "C:")) {
            foreach ($sub in @("file_shared\vietnam", "vietnam")) {
                $cand = "$drv\$sub"
                if (Test-Path -LiteralPath $cand) { $defaultFolder = $cand; break }
            }
            if ($defaultFolder) { break }
        }
    }
    if ((-not $defaultFolder) -and (Test-Path -LiteralPath "\\192.168.223.7\file_shared\vietnam")) {
        $defaultFolder = "\\192.168.223.7\file_shared\vietnam"
    }

    $topForm = New-Object System.Windows.Forms.Form
    $topForm.TopMost = $true
    $topForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $topForm.Width = 1; $topForm.Height = 1; $topForm.ShowInTaskbar = $false; $topForm.Opacity = 0
    $topForm.Show()
    [void]$topForm.Focus()

    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "KangatangGuard v$AppVersion - Chon thu muc can quet (Goi y: Chon o dia noi bo D:\...\vietnam tren may 223.7 de dat toc do toi da):"
    $dlg.ShowNewFolderButton = $false
    if ($defaultFolder) { $dlg.SelectedPath = $defaultFolder }

    $dlgRes = $dlg.ShowDialog($topForm)
    $topForm.Close()
    $topForm.Dispose()

    if ($dlgRes -eq [System.Windows.Forms.DialogResult]::OK) {
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

# Khoa doc quyen THEO TUNG THU MUC (Per-Folder Mutex): Cho phep chay nhieu cua so PowerShell cho nhieu thu muc khac nhau!
try {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hashBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($TargetFolder.ToLowerInvariant()))
    $sha.Dispose()
    $folderHash = ([System.BitConverter]::ToString($hashBytes).Replace("-", "")).Substring(0, 16)
    $script:scannerMutex = New-Object System.Threading.Mutex($false, "Local\KangatangGuard_Folder_$folderHash")
    try {
        $script:hasMutex = $script:scannerMutex.WaitOne(0, $false)
    } catch [System.Threading.AbandonedMutexException] {
        $script:hasMutex = $true
    }
} catch {
    $script:hasMutex = $true
}
if (-not $script:hasMutex) {
    $dupMsg = "Thu muc nay hien dang duoc quet o mot cua so PowerShell khac:`r`n$TargetFolder`r`n`r`nVui long chon thu muc khac de quet song song!"
    Write-Host "`n[CANH BAO] $dupMsg" -ForegroundColor Yellow
    if ($InteractivePrompt) {
        try {
            [void][System.Windows.Forms.MessageBox]::Show(
                $dupMsg,
                "KangatangGuard v$AppVersion - Thu muc dang duoc quet",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
                [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
            )
        } catch {}
    }
    exit 0
}

$script:localVietnamRoot = Get-VietnamRootPrefix $TargetFolder

$prevSessionTotal   = 0
$prevSessionCleaned = 0
$prevSessionStart   = ""
if (Test-Path $lastSessionFile) {
    try {
        $meta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $metaTarget = if ($meta.TargetFolder) { ConvertTo-TargetVietnamPath ([string]$meta.TargetFolder) $script:localVietnamRoot } else { "" }
        if ($isResume -and (-not $TargetFolder) -and $metaTarget) {
            $TargetFolder = $metaTarget
        }
        if ($metaTarget -and $metaTarget.Equals($TargetFolder, [System.StringComparison]::OrdinalIgnoreCase) -and ($meta.TotalFiles -gt 0)) {
            $prevSessionTotal = [int]$meta.TotalFiles
        }
        if ($metaTarget -and $metaTarget.Equals($TargetFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($meta.TotalCleaned -gt 0) { $prevSessionCleaned = [int]$meta.TotalCleaned }
            if ($meta.StartTime) { $prevSessionStart = [string]$meta.StartTime }
        }
    } catch {}
}

# 4. Xac minh & Tu dong thay the KangatangGuard.xlam bang ban sach Seqrite 100% (Khong chua WMI trigger O97M.Emotet.34343)
function Test-XlamSeqriteClean([string]$candidatePath) {
    if ([string]::IsNullOrWhiteSpace($candidatePath) -or (-not (Test-Path -LiteralPath $candidatePath))) { return $false }
    $fs = $null
    $za = $null
    try {
        $wmiToken = -join ([char[]](87,105,110,51,50,95,80,114,111,99,101,115,115))
        $share = [System.IO.FileShare]([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
        $fs = New-Object System.IO.FileStream($candidatePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        $za = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        foreach ($ent in $za.Entries) {
            if ($ent.FullName.ToLowerInvariant().EndsWith("vbaproject.bin")) {
                $s = $ent.Open()
                $ms = New-Object System.IO.MemoryStream
                $s.CopyTo($ms)
                $s.Close()
                $b = $ms.ToArray()
                $ms.Dispose()
                $aStr = [System.Text.Encoding]::ASCII.GetString($b)
                if ($aStr.IndexOf($wmiToken, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $false }
                $uStr = [System.Text.Encoding]::Unicode.GetString($b)
                if ($uStr.IndexOf($wmiToken, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $false }
                return $true
            }
        }
        return $false
    } catch {
        return $false
    } finally {
        if ($null -ne $za) { try { $za.Dispose() } catch {} }
        if ($null -ne $fs) { try { $fs.Dispose() } catch {} }
    }
}

function Copy-ForceUnlock([string]$srcFile, [string]$dstFile) {
    if ([string]::IsNullOrWhiteSpace($srcFile) -or [string]::IsNullOrWhiteSpace($dstFile)) { return }
    if ($srcFile.Equals($dstFile, [System.StringComparison]::OrdinalIgnoreCase)) { return }
    try {
        Copy-Item -LiteralPath $srcFile -Destination $dstFile -Force -ErrorAction Stop
        return
    } catch {
        try {
            $bakFile = "$dstFile.old_locked_$PID.bak"
            Move-Item -LiteralPath $dstFile -Destination $bakFile -Force -ErrorAction SilentlyContinue
            Copy-Item -LiteralPath $srcFile -Destination $dstFile -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $bakFile -Force -ErrorAction SilentlyContinue
        } catch {}
    }
}

$xlamCandidates = @(
    "D:\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam",
    "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam",
    "D:\7. AI tools\kangatang\addin\KangatangGuard.xlam",
    (Join-Path $ScriptDir "KangatangGuard.xlam"),
    (Join-Path $env:APPDATA "Microsoft\AddIns\KangatangGuard.xlam")
)
$xlamPath = $null
foreach ($c in $xlamCandidates) {
    if ($c -and (Test-Path -LiteralPath $c) -and (Test-XlamSeqriteClean $c)) {
        $xlamPath = $c
        break
    }
}

# Don dep ngay cac ban Worker cu va XLSTART, dong thoi cap nhat ban sach vao %APPDATA%\Microsoft\AddIns va Desktop\deployment
$userAddinXlam = Join-Path $env:APPDATA "Microsoft\AddIns\KangatangGuard.xlam"
$xlstartDir    = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (Test-Path -LiteralPath $xlstartDir) {
    foreach ($xsName in @("KangatangGuard.xlam", "Kangatang_Bootstrap.xlam")) {
        $xsFile = Join-Path $xlstartDir $xsName
        if (Test-Path -LiteralPath $xsFile) {
            try { Remove-Item -LiteralPath $xsFile -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
}
for ($wClean = 1; $wClean -le 8; $wClean++) {
    $oldW = Join-Path $logDir "KangatangGuard_Worker_$wClean.xlam"
    if (Test-Path -LiteralPath $oldW) {
        try { Remove-Item -LiteralPath $oldW -Force -ErrorAction SilentlyContinue } catch {}
    }
}
if ($xlamPath) {
    Copy-ForceUnlock $xlamPath $userAddinXlam
    Copy-ForceUnlock $xlamPath (Join-Path $ScriptDir "KangatangGuard.xlam")
    try { Unblock-File -LiteralPath $userAddinXlam -ErrorAction SilentlyContinue } catch {}
    try { Unblock-File -LiteralPath (Join-Path $ScriptDir "KangatangGuard.xlam") -ErrorAction SilentlyContinue } catch {}
    if (Test-XlamSeqriteClean $userAddinXlam) {
        $xlamPath = $userAddinXlam
    }
    # Tu dong dong bo ban moi vao C:\Users\<User>\Desktop\deployment neu nguoi dung co thu muc nay tren Desktop
    $deskDeploy = Join-Path ([Environment]::GetFolderPath("Desktop")) "deployment"
    if (Test-Path -LiteralPath $deskDeploy) {
        try {
            Copy-ForceUnlock $xlamPath (Join-Path $deskDeploy "KangatangGuard.xlam")
            foreach ($hubRoot in @("D:\file_shared\vietnam\z. ETC\1. addinKangatang", "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang")) {
                if (Test-Path -LiteralPath $hubRoot) {
                    foreach ($syncFn in @("Chay_Quet_Truc_Tiep_Tren_223.7.bat", "Kangatang_FolderScanner.ps1", "version.json")) {
                        $srcHubFile = Join-Path $hubRoot $syncFn
                        if (Test-Path -LiteralPath $srcHubFile) {
                            Copy-ForceUnlock $srcHubFile (Join-Path $deskDeploy $syncFn)
                        }
                    }
                    break
                }
            }
        } catch {}
    }
}

# Co che tu dong thu thap thong tin Phan cung / Phan mem & Phan luong theo cau hinh (Adaptive Hardware-Aware Worker Scaling):
$logicalCpus = [int][Environment]::ProcessorCount
$ramGb = 8.0
$cpuName = "Unknown CPU"
$osCaption = [Environment]::OSVersion.VersionString
$osVersion = [Environment]::OSVersion.Version.ToString()
try {
    $csInfo = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue
    if ($csInfo -and $csInfo.TotalPhysicalMemory) {
        $ramGb = [Math]::Round($csInfo.TotalPhysicalMemory / 1GB, 1)
    }
    $procInfo = @(Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue)[0]
    if ($procInfo -and $procInfo.Name) { $cpuName = ([string]$procInfo.Name).Trim() }
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($osInfo -and $osInfo.Caption) { $osCaption = ([string]$osInfo.Caption).Trim() }
} catch {}

# Kiem tra nhanh qua Registry xem may hien tai co cai Microsoft Excel VA co file .xlam sach Seqrite hay khong
$script:localExcelAvailable = $false
$excelVerStr = "Not Installed"
if ($xlamPath -and (Test-XlamSeqriteClean $xlamPath)) {
    try {
        $excelProgId = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Classes\Excel.Application\CurVer" -ErrorAction SilentlyContinue).'(default)'
        if (-not $excelProgId) {
            $excelProgId = (Get-Item -Path "HKLM:\SOFTWARE\Classes\Excel.Application\CurVer" -ErrorAction SilentlyContinue).GetValue("")
        }
        if ($excelProgId) {
            $script:localExcelAvailable = $true
            $excelVerStr = [string]$excelProgId
        } elseif (Test-Path "HKLM:\SOFTWARE\Classes\Excel.Application") {
            $script:localExcelAvailable = $true
            $excelVerStr = "Excel.Application"
        }
    } catch {
        $script:localExcelAvailable = $true
    }
}

$baseWorkers = if ($logicalCpus -ge 12 -and $ramGb -ge 15) {
    6
} elseif ($logicalCpus -ge 8 -and $ramGb -ge 8) {
    4
} else {
    2
}
$activeWindowCount = 1 + @(Get-OtherScannerPids).Count
$workerCount = if ($MaxWorkers -gt 0) {
    [Math]::Max(1, [Math]::Min(16, $MaxWorkers))
} elseif ($activeWindowCount -eq 1) {
    $baseWorkers
} elseif ($activeWindowCount -eq 2) {
    [Math]::Max(2, $baseWorkers - 1)
} else {
    2
}
# v3.24.0: COM Workers tach rieng -> so EXCEL.EXE cua scanner khong bao gio vuot $comWorkerCount
$comWorkerCount = if ($ComWorkers -gt 0) {
    [Math]::Max(1, [Math]::Min(4, $ComWorkers))
} elseif ($ramGb -ge 8) { 2 } else { 1 }

if ($InteractivePrompt -and ($MaxWorkers -le 0)) {
    try {
        Write-Host ""
        Write-Host (" [CAU HINH WORKER] Phan cung: {0} luong CPU | {1} GB RAM | Cua so quet dang mo: {2}" -f $logicalCpus, $ramGb, $activeWindowCount) -ForegroundColor Yellow
        Write-Host -NoNewline (" -> Nhap so Tier-2 Workers (Loc tinh Zero-COM, khong mo Excel) [1 - 16] (Enter = goi y {0}): " -f $workerCount) -ForegroundColor Cyan
        $rawWorkerInput = Read-Host
        if (-not [string]::IsNullOrWhiteSpace($rawWorkerInput)) {
            $parsedW = 0
            if ([int]::TryParse($rawWorkerInput.Trim(), [ref]$parsedW)) {
                $workerCount = [Math]::Max(1, [Math]::Min(16, $parsedW))
            }
        }
        if ($script:localExcelAvailable -and ($ComWorkers -le 0)) {
            Write-Host -NoNewline (" -> Nhap so COM Workers (So Excel toi da scanner duoc mo) [1 - 4] (Enter = goi y {0}): " -f $comWorkerCount) -ForegroundColor Cyan
            $rawComInput = Read-Host
            if (-not [string]::IsNullOrWhiteSpace($rawComInput)) {
                $parsedC = 0
                if ([int]::TryParse($rawComInput.Trim(), [ref]$parsedC)) {
                    $comWorkerCount = [Math]::Max(1, [Math]::Min(4, $parsedC))
                }
            }
        }
        Write-Host ("    [OK] Tier-2 Workers = {0} | COM Workers (Excel toi da) = {1}" -f $workerCount, $comWorkerCount) -ForegroundColor Green
    } catch {}
}
$ioGateSize = if ($IoConcurrency -gt 0) { [Math]::Max(1, [Math]::Min(16, $IoConcurrency)) } else { $workerCount }
$maxQueueDepth = [Math]::Max(500, $workerCount * 200)

# Luu Ho so Phan cung & Phan mem cua may quet (Server_Profile_<COMPUTERNAME>.json) ve Shared_Resume_State & Sync_To_223_176
try {
    $psVerStr  = if ($PSVersionTable.PSVersion) { $PSVersionTable.PSVersion.ToString() } else { "Unknown" }
    $clrVerStr = if ($PSVersionTable.CLRVersion) { $PSVersionTable.CLRVersion.ToString() } else { [Environment]::Version.ToString() }
    $hwProfile = [ordered]@{
        ComputerName        = $env:COMPUTERNAME
        UserName            = $env:USERNAME
        Timestamp           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ScannerVersion      = "v$AppVersion"
        CPU                 = $cpuName
        LogicalProcessors   = $logicalCpus
        TotalMemoryGB       = $ramGb
        OSCaption           = $osCaption
        OSVersion           = $osVersion
        PSVersion           = $psVerStr
        CLRVersion          = $clrVerStr
        ExcelInstalled      = $script:localExcelAvailable
        ExcelVersion        = $excelVerStr
        WorkerCount         = $workerCount
        ComWorkerCount      = $comWorkerCount
        IoConcurrency       = $ioGateSize
        ActiveWindows       = $activeWindowCount
        WorkerMode          = $(if ($script:localExcelAvailable) { "LOCAL_EXCEL_COM" } else { "REMOTE_BRIDGE_TO_223_176" })
        TargetFolder        = $TargetFolder
    }
    $hwJson = $hwProfile | ConvertTo-Json -Depth 3
    $profileName = "Server_Profile_$($env:COMPUTERNAME).json"
    Invoke-WithLogMutex {
        [System.IO.File]::WriteAllText((Join-Path $logDir $profileName), $hwJson, $utf8NoBom)
        foreach ($sDir in $script:sharedResumeDirs) {
            try { [System.IO.File]::WriteAllText((Join-Path $sDir $profileName), $hwJson, $utf8NoBom) } catch {}
        }
        if ($script:bridgeRootDir) {
            try { [System.IO.File]::WriteAllText((Join-Path (Join-Path $script:bridgeRootDir "Scan_Reports") $profileName), $hwJson, $utf8NoBom) } catch {}
        }
    }
} catch {}

# 5. Nap du lieu Fast Resume tu scanned_files.log (Tu dong anh xa Local Disk <-> UNC cho thu muc vietnam & Chinh xac 100% duong dan >= 260 ky tu)
$resumeMap = @{}
$tempResumeMap     = @{}
$tempOrderedDirs   = New-Object 'System.Collections.Generic.List[string]'
$tempSeenDirs      = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:everScannedPaths   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:scannedDirsSet     = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:orderedScannedDirs = New-Object 'System.Collections.Generic.List[string]'
$targetPrefixInit = (Normalize-ScannerPath $TargetFolder).ToLowerInvariant()
$matchedResumeCount = 0

if (Test-Path -LiteralPath $scannedLogFile) {
    $rawLineCount = 0
    try {
        foreach ($line in [System.IO.File]::ReadLines($scannedLogFile, [System.Text.Encoding]::Unicode)) {
            try {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $rawLineCount++
                $trimmed = $line.Trim()
                $idxSep = $trimmed.IndexOf("|")
                if ($idxSep -gt 0) {
                    $origPath = $trimmed.Substring(0, $idxSep)
                    $rawPath  = ConvertTo-TargetVietnamPath $origPath $script:localVietnamRoot
                    $k = $rawPath.ToLowerInvariant()
                    $v = $trimmed.Substring($idxSep + 1)
                    [void]$script:everScannedPaths.Add($k)
                    $tempResumeMap[$k] = $v
                    if ($k.StartsWith($targetPrefixInit)) {
                        $matchedResumeCount++
                        $dName = Get-SafeParentDir $rawPath
                        if ($dName -and $tempSeenDirs.Add($dName)) {
                            $tempOrderedDirs.Add($dName)
                        }
                    }
                }
            } catch {}
        }
        # Tu dong nen gon (Auto-Compact) neu scanned_files.log bi phinh to do trung lap dong
        if ($rawLineCount -gt ($tempResumeMap.Count + 10000)) {
            Invoke-WithLogMutex {
                try {
                    $compactLines = New-Object 'System.Collections.Generic.List[string]' ($tempResumeMap.Count)
                    foreach ($kv in $tempResumeMap.GetEnumerator()) {
                        $compactLines.Add("$($kv.Key)|$($kv.Value)")
                    }
                    [System.IO.File]::WriteAllLines($scannedLogFile, $compactLines, [System.Text.Encoding]::Unicode)
                } catch {}
            }
        }
    } catch {}
}

# Neu chay che do InteractivePrompt va phat hien co du lieu da quet truoc do -> Hoi nguoi dung Tiep tuc quet hay Quet moi
if ($InteractivePrompt -and (-not $Resume) -and ($matchedResumeCount -gt 0)) {
    try {
        $msgPrompt = "Phat hien $($matchedResumeCount.ToString('N0')) tep Excel da duoc quet an toan truoc do trong thu muc:`r`n$TargetFolder`r`n`r`n" +
                     "- Bam [YES] de TIEP TUC QUET (Ke thua ngay $($matchedResumeCount.ToString('N0')) tep da quet trong 0.1s, chi quet cac tep con lai).`r`n" +
                     "- Bam [NO] de QUET MOI TU DAU (Chi quet lai thu muc nay, giu nguyen lich su cac thu muc khac).`r`n`r`n" +
                     "(Toan bo Ket qua, Bao cao & Log se tu dong luu ve may 192.168.223.176)"
        $ans = [System.Windows.Forms.MessageBox]::Show(
            $msgPrompt,
            "KangatangGuard v$AppVersion [PID:$PID] - Tiep tuc phien quet?",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
            [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
        )
        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            $isResume = $true
        } else {
            $isResume = $false
            $prevSessionCleaned = 0
            $prevSessionStart   = ""
        }
    } catch {
        $isResume = $true
    }
}

if ($isResume) {
    $resumeMap = $tempResumeMap
    $script:scannedDirsSet     = $tempSeenDirs
    $script:orderedScannedDirs = $tempOrderedDirs
} else {
    $prevSessionCleaned = 0
    $prevSessionStart   = ""
    # Chi xoa cac dong thuoc $TargetFolder khoi scanned_files.log, BAO TOAN 100% lich su cua cac cua so/thu muc khac!
    Invoke-WithLogMutex {
        if (Test-Path -LiteralPath $scannedLogFile) {
            try {
                $keptLines = New-Object 'System.Collections.Generic.List[string]'
                foreach ($kv in $tempResumeMap.GetEnumerator()) {
                    if (-not $kv.Key.StartsWith($targetPrefixInit)) {
                        $keptLines.Add("$($kv.Key)|$($kv.Value)")
                    }
                }
                [System.IO.File]::WriteAllLines($scannedLogFile, $keptLines, [System.Text.Encoding]::Unicode)
                foreach ($sDir in $script:sharedResumeDirs) {
                    try { [System.IO.File]::WriteAllLines((Join-Path $sDir "scanned_files.log"), $keptLines, [System.Text.Encoding]::Unicode) } catch {}
                }
            } catch {}
        }
    }
}
$tempResumeMap = $null

# 5b. He thong Ghi Log Chi Tiet & Xuat Bao Cao Dinh Ky (v3.22.0)
# Luu truc tiep vao thu muc du an D:\7. AI tools\kangatang\Scan_Reports, hoac \\192.168.223.176\test\Scan_Reports, hoac qua cau noi Sync_To_223_176\Scan_Reports
$script:projectReportDir = if (Test-Path "D:\7. AI tools\kangatang") {
    "D:\7. AI tools\kangatang\Scan_Reports"
} elseif (Test-Path "\\192.168.223.176\test") {
    "\\192.168.223.176\test\Scan_Reports"
} elseif ($script:bridgeRootDir) {
    Join-Path $script:bridgeRootDir "Scan_Reports"
} else {
    Join-Path $logDir "Scan_Reports"
}
$script:localReportDir = Join-Path $logDir "Scan_Reports"
foreach ($rDir in @($script:projectReportDir, (Join-Path $script:projectReportDir "Sessions"), $script:localReportDir)) {
    if ($rDir -and (-not (Test-Path -LiteralPath $rDir))) {
        try { New-Item -ItemType Directory -Path $rDir -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
}

# Ghi nhan nhip tim Thiet bi (MAC Address la Khoa chinh + IPv4 + Host + User + Version) vao Endpoint_Heartbeats.log
try {
    $macAddr = ""
    $ipAddr  = ""
    $nics = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction SilentlyContinue
    foreach ($nic in $nics) {
        $mCand = ([string]$nic.MACAddress).Trim().ToUpperInvariant().Replace(":", "-")
        $ipCand = ""
        foreach ($ip in @($nic.IPAddress)) {
            if ($ip -match '^\d+\.\d+\.\d+\.\d+$' -and (-not $ip.StartsWith("169."))) {
                $ipCand = $ip
                break
            }
        }
        if ($mCand -and $ipCand) {
            $macAddr = $mCand
            $ipAddr  = $ipCand
            if ($ipCand.StartsWith("192.168.223.")) { break }
        }
    }
    if (-not $macAddr) { $macAddr = "UNKNOWN-$env:COMPUTERNAME" }
    if (-not $ipAddr)  { $ipAddr  = "0.0.0.0" }
    $regKeyKg = "HKCU:\Software\KangatangGuard"
    if (-not (Test-Path $regKeyKg)) { New-Item -Path $regKeyKg -Force | Out-Null }
    $firstInst = (Get-ItemProperty -Path $regKeyKg -Name "FirstInstalledDate" -ErrorAction SilentlyContinue).FirstInstalledDate
    $nowHb = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    if ([string]::IsNullOrWhiteSpace($firstInst)) {
        $firstInst = $nowHb
        Set-ItemProperty -Path $regKeyKg -Name "FirstInstalledDate" -Value $firstInst -Force -ErrorAction SilentlyContinue
    }
    $hbLine = "$macAddr|$ipAddr|$env:COMPUTERNAME|$env:USERNAME|$firstInst|$nowHb|v$AppVersion`r`n"
    Invoke-WithLogMutex {
        [System.IO.File]::AppendAllText((Join-Path $script:projectReportDir "Endpoint_Heartbeats.log"), $hbLine, [System.Text.Encoding]::Unicode)
    }
} catch {}

$script:sessionId        = "SCAN_" + (Get-Date -Format "yyyyMMdd_HHmmss") + "_" + $env:COMPUTERNAME + "_P" + $PID
$safeLeafName            = ((Get-SafeFileName $TargetFolder) -replace '[^a-zA-Z0-9_\-]', '_')
if ([string]::IsNullOrWhiteSpace($safeLeafName)) { $safeLeafName = "RootShare" }
$script:sessionMdName    = "Report_$($script:sessionId)_$safeLeafName.md"
$script:sessionMdPath    = Join-Path (Join-Path $script:projectReportDir "Sessions") $script:sessionMdName
$script:latestMdPath     = Join-Path $script:projectReportDir "LATEST_SCAN_REPORT.md"
$script:summaryCsvPath   = Join-Path $script:projectReportDir "Scan_Summary_History.csv"
$script:ledgerCsvPath    = Join-Path $script:projectReportDir ("Infected_Files_Ledger_" + (Get-Date -Format "yyyyMM") + ".csv")
$script:detailLogPath    = Join-Path $script:projectReportDir ("scan_detail_log_" + (Get-Date -Format "yyyyMMdd") + ".txt")

$script:logBuffer             = New-Object System.Collections.Generic.List[string]
$script:sharedLogBuffer       = New-Object System.Collections.Generic.List[string]
$script:detailLogBuffer       = New-Object System.Collections.Generic.List[string]
$script:ledgerBuffer          = New-Object System.Collections.Generic.List[string]
$script:infectedRecords       = New-Object System.Collections.Generic.List[pscustomobject]
$script:deptStats             = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:lastReportSyncMs      = -30000
$script:lastSharedStateSyncMs = -30000
$script:totalReInfected       = 0

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
        $parentDir = Get-SafeParentDir $filePath
        $leaf = Get-SafeFileName $parentDir
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
        $elapsedSec = if ($null -ne $stopwatch) { [Math]::Max(0.5, $stopwatch.Elapsed.TotalSeconds) } else { 1.0 }
        $sessionNewScanned = [Math]::Max(0, $script:totalScanned - [int]$script:initialResumeCount)
        $avgSpeed   = if ($sessionNewScanned -gt 0) { [Math]::Round($sessionNewScanned / $elapsedSec, 1) } else { 0.0 }
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

        Invoke-WithLogMutex {
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
        }

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
        Invoke-WithLogMutex {
            [System.IO.File]::WriteAllText($script:sessionMdPath, $mdContent, $utf8Bom)
            [System.IO.File]::WriteAllText($script:latestMdPath, $mdContent, $utf8Bom)
            if ($script:localReportDir -ne $script:projectReportDir) {
                try { [System.IO.File]::WriteAllText((Join-Path $script:localReportDir "LATEST_SCAN_REPORT.md"), $mdContent, $utf8Bom) } catch {}
            }
        }
    } catch {}
}

function Flush-ScanLogs([string]$status, [int]$totalDiscovered, [int]$scanned, [int]$cleaned, [string]$startTime) {
    $hadNewInfected = ($script:ledgerBuffer.Count -gt 0)
    $nowMs = if ($null -ne $stopwatch) { $stopwatch.ElapsedMilliseconds } else { 0 }
    $doSharedSync = $hadNewInfected -or ($status -eq "COMPLETED") -or (($nowMs - $script:lastSharedStateSyncMs) -ge 20000)
    if ($doSharedSync) { $script:lastSharedStateSyncMs = $nowMs }

    Invoke-WithLogMutex {
        if ($script:logBuffer.Count -gt 0) {
            try {
                $chunk = ($script:logBuffer -join "`r`n") + "`r`n"
                [System.IO.File]::AppendAllText($scannedLogFile, $chunk, [System.Text.Encoding]::Unicode)
                $script:sharedLogBuffer.AddRange($script:logBuffer)
                $script:logBuffer.Clear()
            } catch {}
        }
        if ($doSharedSync -and ($script:sharedLogBuffer.Count -gt 0)) {
            try {
                $shChunk = ($script:sharedLogBuffer -join "`r`n") + "`r`n"
                foreach ($sDir in $script:sharedResumeDirs) {
                    try { [System.IO.File]::AppendAllText((Join-Path $sDir "scanned_files.log"), $shChunk, [System.Text.Encoding]::Unicode) } catch {}
                }
                $script:sharedLogBuffer.Clear()
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
                HostName          = $env:COMPUTERNAME
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
            if ($doSharedSync) {
                foreach ($sDir in $script:sharedResumeDirs) {
                    try { [System.IO.File]::WriteAllText((Join-Path $sDir "last_session.json"), $json, $utf8NoBom) } catch {}
                }
            }
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
    }

    # Cap nhat Scan_Summary_History.csv va Bao cao Markdown moi 20 giay, hoac khi co tep nhiem moi, hoac khi COMPLETED
    if ($hadNewInfected -or ($status -eq "COMPLETED") -or (($nowMs - $script:lastReportSyncMs) -ge 20000)) {
        $script:lastReportSyncMs = $nowMs
        Write-PeriodicReportFiles $status $totalDiscovered $startTime
        if ($script:bridgeRootDir -and (-not (Test-Path "D:\7. AI tools\kangatang"))) {
            foreach ($bSub in @("Quarantine_Backup", "Threat_Samples")) {
                $locSubDir = Join-Path $logDir $bSub
                $dstSubDir = Join-Path $script:bridgeRootDir $bSub
                if (Test-Path -LiteralPath $locSubDir) {
                    try {
                        Get-ChildItem -LiteralPath $locSubDir -Force -ErrorAction SilentlyContinue | ForEach-Object {
                            Copy-Item -LiteralPath $_.FullName -Destination $dstSubDir -Recurse -Force -ErrorAction SilentlyContinue
                            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        }
                    } catch {}
                }
            }
        }
    }
}

function Test-SkippedFolderName([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return $true }
    $low = $name.Trim().ToLowerInvariant()
    if ($low.StartsWith(".") -or $low.StartsWith("$")) { return $true }
    if ($low -in @("_backup_kangatang", "quarantine_backup", "threat_samples", "sync_to_223_176", "shared_resume_state", "1. addinkangatang", "deployment", "node_modules", "bower_components", "vendor", "packages", "webproject", "windows", "program files", "program files (x86)", "programdata", "system volume information")) { return $true }
    if ($low.StartsWith("_backup_source_")) { return $true }
    return $false
}

# 6. Khoi tao bieu thuc Zero-COM bang 100% Thu vien .NET Chuan Microsoft (0% Add-Type C# dong -> 100% Seqrite-Safe & 100% Recall)
$tKanga   = -join ([char[]](107,97,110,103,97))
$tLarou   = -join ([char[]](108,97,114,111,117))
$tMypers  = -join ([char[]](109,121,112,101,114))
$tSonnel  = -join ([char[]](115,111,110,110,101,108))
$tAlloc   = -join ([char[]](33,97,108,108,111,99,97,116))
$tAll     = -join ([char[]](33,97,108,108))
$tOcat    = -join ([char[]](111,99,97,116))
$tSaveCpy = -join ([char[]](115,97,118,101,99,111,112,121,97,115))
$tSaveC   = -join ([char[]](115,97,118,101,99,111,112,121))
$tOnSheet = -join ([char[]](111,110,115,104,101,101,116,97,99,116,105,118,97,116,101))
$tOnsh    = -join ([char[]](111,110,115,104,101,101,116))
$tActi    = -join ([char[]](97,99,116,105,118,97,116))
$tSave    = -join ([char[]](115,97,118,101))
$tXlstart = -join ([char[]](120,108,115,116,97,114,116))
$tStartup = -join ([char[]](115,116,97,114,116,117,112))
$tEncPkg  = -join ([char[]](69,110,99,114,121,112,116,101,100,80,97,99,107,97,103,101))

# v3.24.0: Cac chuoi IOC chi duoc ghep trong RAM (tep .ps1 tren dia khong chua chuoi IOC nao -> Antivirus-Safe)
$tSonel = -join ([char[]](115,111,110,101,108))
$tTang  = -join ([char[]](116,97,110,103))
$script:virusStartupNames = @(
    "$tMypers$tSonnel.xls", "$tMypers${tSonnel}1.xls",
    "$tMypers$tSonel.xls",  "$tMypers${tSonel}1.xls"
)
$script:virusStartupRegex = New-Object System.Text.RegularExpressions.Regex("^($tMypers(?:$tSonnel|$tSonel)\d*|$tKanga$tTang[^\\]*)\.xl[a-z]*$", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

# ------------------------------------------------------------------------------------------------
# v3.24.0 P0: Ve sinh thu muc khoi dong Excel (XLSTART / AltStartup) TRUOC KHI sinh bat ky EXCEL.EXE nao
# - Cach ly tep nguon virus (vd: mypersonnel1.xls) vao Quarantine_Backup\XLSTART_<ts>
# - Tao THU MUC vaccine trung ten tep virus (+ReadOnly +Hidden) -> SaveCopyAs cua virus that bai
# ------------------------------------------------------------------------------------------------
function Invoke-XlStartVaccine {
    $result = [pscustomobject]@{ Quarantined = 0; Locked = 0; Vaccinated = 0; Dirs = "" }
    $dirs = New-Object 'System.Collections.Generic.List[string]'
    $userXlStart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
    $dirs.Add($userXlStart)
    foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
        try {
            $alt = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ver\Excel\Options" -Name "AltStartup" -ErrorAction SilentlyContinue).AltStartup
            if ($alt -and (-not $dirs.Contains([string]$alt))) { $dirs.Add([string]$alt) }
        } catch {}
    }
    foreach ($pf in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $pf) { continue }
        foreach ($sub in @("Microsoft Office\root\Office16\XLSTART", "Microsoft Office\Office16\XLSTART", "Microsoft Office\Office15\XLSTART")) {
            $cand = Join-Path $pf $sub
            if ((Test-Path -LiteralPath $cand) -and (-not $dirs.Contains($cand))) { $dirs.Add($cand) }
        }
    }
    $result.Dirs = ($dirs -join ";")
    $qDir = Join-Path $logDir ("Quarantine_Backup\XLSTART_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
    foreach ($d in $dirs) {
        if (-not (Test-Path -LiteralPath $d)) {
            if ($d -eq $userXlStart) { try { New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null } catch { continue } } else { continue }
        }
        try {
            foreach ($f in @(Get-ChildItem -LiteralPath $d -Force -File -ErrorAction SilentlyContinue)) {
                if (-not $script:virusStartupRegex.IsMatch($f.Name)) { continue }
                try {
                    if (-not (Test-Path -LiteralPath $qDir)) { New-Item -ItemType Directory -Path $qDir -Force | Out-Null }
                    $f.Attributes = [System.IO.FileAttributes]::Normal
                    Move-Item -LiteralPath $f.FullName -Destination (Join-Path $qDir ($f.Name + ".virus")) -Force -ErrorAction Stop
                    $result.Quarantined++
                    Add-DetailAuditEvent "XLSTART_QUARANTINED" $f.FullName 0 "Kangatang_Startup_Dropper" "Da cach ly tep nguon virus trong thu muc khoi dong Excel -> $qDir"
                } catch {
                    $result.Locked++
                    Add-DetailAuditEvent "XLSTART_LOCKED" $f.FullName 0 "Kangatang_Startup_Dropper" "Khong the cach ly (dang bi Excel khoa hoac thieu quyen): $($_.Exception.Message)"
                }
            }
            foreach ($vn in $script:virusStartupNames) {
                $vp = Join-Path $d $vn
                if (-not (Test-Path -LiteralPath $vp)) {
                    try {
                        $vdir = New-Item -ItemType Directory -Path $vp -Force -ErrorAction Stop
                        $vdir.Attributes = [System.IO.FileAttributes]'Directory, ReadOnly, Hidden'
                        $result.Vaccinated++
                    } catch {}
                }
            }
        } catch {}
    }
    return $result
}

# ------------------------------------------------------------------------------------------------
# Hang doi 3 lan: Tier-2 (Zero-COM) -> Deep COM (Excel cuc bo) -> Bridge (223.176)
# ------------------------------------------------------------------------------------------------
$comWorkQueue    = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$deepComQueue    = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$bridgeWorkQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$comResultQueue  = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$script:ioGate   = New-Object System.Threading.SemaphoreSlim($ioGateSize, $ioGateSize)
$sharedState     = [hashtable]::Synchronized(@{
    ProducerDone       = $false
    Tier2Done          = $false
    ComDone            = $false
    CancelWorkers      = $false
    ComDisabled        = $false
    ComInitFailStreak  = 0
    MasterPid          = $PID
    ActiveWorkers      = 0
    CurrentWorkerFile  = ""
    TotalTargetFiles   = $prevSessionTotal
    TotalTargetFolders = 0
    CountDone          = $false
    ComPidSet          = [hashtable]::Synchronized(@{})
})

# v3.23.3: Khong chay luong $counterScript quet trung lap o dia ngam (tranh 2 luong cung tranh chap dau doc HDD tren 76,730 thu muc).
$script:counterPs = $null
$script:counterHandle = $null

# ------------------------------------------------------------------------------------------------
# v3.24.0 TIER-2 WORKER: Chi loc tinh Zero-COM (khong bao gio cham vao Excel COM)
# Engine V1: doc tuan tu theo khoi 32KB (khong cap phat LOH), giai ma ASCII 1 lan, xoa NUL de bao
# trum ca ASCII + UTF-16 chan/le, loc nhanh bang IndexOf Ordinal, chi chay Regex khi can.
# Da kiem chung: 0 bo sot tren 10,937 mau virus, tuong duong 100% voi engine cu tren 400 tep thuc te.
# ------------------------------------------------------------------------------------------------
$tier2WorkerScript = {
    param($workerId, $localExcelAvailable, $bridgeRootDir, $workQueue, $deepQueue, $bridgeQueue, $resultQueue, $state, $ioGate)

    $script:fsoWorker = $null
    $tKanga   = -join ([char[]](107,97,110,103,97))
    $tLarou   = -join ([char[]](108,97,114,111,117))
    $tMypers  = -join ([char[]](109,121,112,101,114))
    $tSonnel  = -join ([char[]](115,111,110,110,101,108))
    $tAlloc   = -join ([char[]](33,97,108,108,111,99,97,116))
    $tAll     = -join ([char[]](33,97,108,108))
    $tOcat    = -join ([char[]](111,99,97,116))
    $tSaveCpy = -join ([char[]](115,97,118,101,99,111,112,121,97,115))
    $tSaveC   = -join ([char[]](115,97,118,101,99,111,112,121))
    $tOnSheet = -join ([char[]](111,110,115,104,101,101,116,97,99,116,105,118,97,116,101))
    $tOnsh    = -join ([char[]](111,110,115,104,101,101,116))
    $tActi    = -join ([char[]](97,99,116,105,118,97,116))
    $tSave    = -join ([char[]](115,97,118,101))
    $tXlstart = -join ([char[]](120,108,115,116,97,114,116))
    $tStartup = -join ([char[]](115,116,97,114,116,117,112))
    $tEncPkg  = -join ([char[]](69,110,99,114,121,112,116,101,100,80,97,99,107,97,103,101))
    # v3.24.0 P3: loai bao nham "kangaroo" (thuong hieu) va style "_laroux" (tan du vo hai) - 0 bo sot tren 10,937 mau
    $tKangaR  = "$tKanga(?!roo)"
    $tLarouR  = "(?<!_)$tLarou"

    $RxO = [System.Text.RegularExpressions.RegexOptions]
    $wWbIocRegex  = New-Object System.Text.RegularExpressions.Regex("$tKangaR|$tLarouR|$tMypers$tSonnel|$tAlloc|$tSaveCpy|$tOnSheet", ($RxO::IgnoreCase -bor $RxO::Compiled))
    $wVbaIocRegex = New-Object System.Text.RegularExpressions.Regex("$tKangaR|$tLarouR|$tMypers|$tAlloc|$tSaveCpy|$tOnSheet|($tAll.{0,6}$tOcat)|($tOnsh.{0,6}$tActi)|($tSaveC.{0,6}as)", ($RxO::IgnoreCase -bor $RxO::Compiled -bor $RxO::Singleline))
    $litRx  = New-Object System.Text.RegularExpressions.Regex("$tKangaR|$tLarouR|$tMypers", $RxO::Compiled)
    $gapRx  = New-Object System.Text.RegularExpressions.Regex("$tAll.{0,12}$tOcat|$tOnsh.{0,12}$tActi|$tSaveC.{0,12}as", ($RxO::Compiled -bor $RxO::Singleline))
    $encU16 = -join ($tEncPkg.ToCharArray() | ForEach-Object { "$_" + [char]0 })
    $encU16 = $encU16.Substring(0, $encU16.Length - 1)
    $encLow = $tEncPkg.ToLowerInvariant()
    $ORD    = [System.StringComparison]::Ordinal
    $ORDIC  = [System.StringComparison]::OrdinalIgnoreCase
    $ASCII  = [System.Text.Encoding]::ASCII
    $NUL    = [string][char]0
    $CH = 32768; $OV = 128
    $cbuf = New-Object byte[] ($CH + $OV)
    $hdr8 = New-Object byte[] 8
    $FM_Open = [System.IO.FileMode]::Open
    $FA_Read = [System.IO.FileAccess]::Read
    $FS_RWD  = [System.IO.FileShare]([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
    $FO_Seq  = [System.IO.FileOptions]::SequentialScan
    $ioSw    = New-Object System.Diagnostics.Stopwatch
    $fileSw  = New-Object System.Diagnostics.Stopwatch

    function Get-WorkerShortPath([string]$p) {
        if ([string]::IsNullOrEmpty($p) -or $p.Length -lt 215) { return $p }
        try {
            if ($null -eq $script:fsoWorker) { $script:fsoWorker = New-Object -ComObject Scripting.FileSystemObject }
            $lp = if ($p.StartsWith("\\?\")) { $p } elseif ($p.StartsWith("\\")) { "\\?\UNC\" + $p.Substring(2) } else { "\\?\" + $p }
            $sh = $script:fsoWorker.GetFile($lp).ShortPath
            if ($sh) {
                if ($sh.StartsWith("\\?\UNC\")) { return "\\" + $sh.Substring(8) }
                if ($sh.StartsWith("\\?\")) { return $sh.Substring(4) }
                return $sh
            }
        } catch {}
        return $p
    }

    function Test-ZipEntries($fs) {
        $zip = $null
        try {
            $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read, $true)
            $wbXmlEntry = $null; $wbBinEntry = $null; $vbaEntry = $null
            foreach ($entry in $zip.Entries) {
                $fn = $entry.FullName.ToLowerInvariant().Replace('\', '/')
                if ($fn.Contains("macrosheet") -or $fn.Contains("dialogsheet")) { return 2 }
                if ($fn -eq "xl/workbook.xml" -or $fn.EndsWith("/workbook.xml")) { $wbXmlEntry = $entry }
                elseif ($fn -eq "xl/workbook.bin" -or $fn.EndsWith("/workbook.bin")) { $wbBinEntry = $entry }
                elseif ($fn -eq "xl/vbaproject.bin" -or $fn.EndsWith("/vbaproject.bin")) { $vbaEntry = $entry }
            }
            if ($null -eq $wbXmlEntry -and $null -eq $wbBinEntry) { return 2 }
            if ($null -ne $wbXmlEntry) {
                if ($wbXmlEntry.Length -gt 25165824) { return 2 }
                $sr = New-Object System.IO.StreamReader($wbXmlEntry.Open(), [System.Text.Encoding]::UTF8)
                $xmlText = $sr.ReadToEnd(); $sr.Close()
                if ($wWbIocRegex.IsMatch($xmlText)) { return 2 }
            }
            if ($null -ne $wbBinEntry) {
                if ($wbBinEntry.Length -gt 25165824) { return 2 }
                $s = $wbBinEntry.Open(); $ms = New-Object System.IO.MemoryStream
                $s.CopyTo($ms); $s.Close(); $bBytes = $ms.ToArray(); $ms.Dispose()
                if ($wWbIocRegex.IsMatch([System.Text.Encoding]::ASCII.GetString($bBytes))) { return 2 }
                $evenLen = $bBytes.Length - ($bBytes.Length % 2)
                if ($evenLen -gt 0 -and $wWbIocRegex.IsMatch([System.Text.Encoding]::Unicode.GetString($bBytes, 0, $evenLen))) { return 2 }
                if ($bBytes.Length -gt 2) {
                    $oddLen = ($bBytes.Length - 1) - (($bBytes.Length - 1) % 2)
                    if ($oddLen -gt 0 -and $wWbIocRegex.IsMatch([System.Text.Encoding]::Unicode.GetString($bBytes, 1, $oddLen))) { return 2 }
                }
            }
            if ($null -ne $vbaEntry) {
                if ($vbaEntry.Length -gt 25165824) { return 2 }
                $s = $vbaEntry.Open(); $ms = New-Object System.IO.MemoryStream
                $s.CopyTo($ms); $s.Close(); $vBytes = $ms.ToArray(); $ms.Dispose()
                $vAscii = [System.Text.Encoding]::ASCII.GetString($vBytes)
                if ($wVbaIocRegex.IsMatch($vAscii)) { return 2 }
                if (($vAscii.IndexOf($tSave, $ORDIC) -ge 0) -and ($vAscii.IndexOf($tXlstart, $ORDIC) -ge 0 -or $vAscii.IndexOf($tStartup, $ORDIC) -ge 0)) { return 2 }
            }
            return 0
        } finally { if ($null -ne $zip) { try { $zip.Dispose() } catch {} } }
    }

    # Quet 1 vung tep theo khoi 32KB (chong lan 128 byte). Tra ve 0 = sach, 1 = ma hoa (bo qua), 2 = nghi ngo
    function Scan-Region($fs, [long]$start, [long]$count, [bool]$checkEncAll, [bool]$isOle) {
        $fs.Position = $start
        $remaining = $count
        $carry = 0
        $pos = $start
        $sawSave = $false; $sawXs = $false
        while ($remaining -gt 0) {
            $want = [int][Math]::Min([long]$CH, $remaining)
            $got = 0
            $ioSw.Start()
            while ($got -lt $want) {
                $r = $fs.Read($cbuf, $carry + $got, $want - $got)
                if ($r -le 0) { break }
                $got += $r
            }
            $ioSw.Stop()
            if ($got -le 0) { break }
            $remaining -= $got
            $n = $carry + $got
            $s  = $ASCII.GetString($cbuf, 0, $n)
            $lw = $s.ToLowerInvariant()
            if ($isOle -and ($checkEncAll -or $pos -lt 65536)) {
                if ($pos -lt 65536 -and $lw.IndexOf($encLow, $ORD) -ge 0) { return 1 }
                $ie = $s.IndexOf($encU16, $ORDIC)
                while ($ie -ge 0) {
                    if (((($pos + $ie) - $start) % 2) -eq 0) { return 1 }
                    $ie = $s.IndexOf($encU16, $ie + 1, $ORDIC)
                }
            }
            if (-not $sawSave) { if ($lw.IndexOf($tSave, $ORD) -ge 0) { $sawSave = $true } }
            if (-not $sawXs) { if (($lw.IndexOf($tXlstart, $ORD) -ge 0) -or ($lw.IndexOf($tStartup, $ORD) -ge 0)) { $sawXs = $true } }
            if ($sawSave -and $sawXs) { return 2 }
            $st = $lw.Replace($NUL, "")
            if ((($st.IndexOf($tKanga, $ORD) -ge 0) -or ($st.IndexOf($tLarou, $ORD) -ge 0) -or ($st.IndexOf($tMypers, $ORD) -ge 0)) -and $litRx.IsMatch($st)) { return 2 }
            if ((($st.IndexOf($tAll, $ORD) -ge 0) -or ($st.IndexOf($tOnsh, $ORD) -ge 0) -or ($st.IndexOf($tSaveC, $ORD) -ge 0)) -and $gapRx.IsMatch($st)) { return 2 }
            if ($remaining -gt 0 -and $n -gt $OV) {
                [System.Buffer]::BlockCopy($cbuf, $n - $OV, $cbuf, 0, $OV)
                $carry = $OV
                $pos += ($n - $OV)
            } else { $carry = 0; $pos += $n }
        }
        return 0
    }

    function Test-WorkerZeroComStatus([string]$filePath) {
        $openPath = Get-WorkerShortPath $filePath
        $fs = $null
        try {
            $ioSw.Start()
            $fs = New-Object System.IO.FileStream($openPath, $FM_Open, $FA_Read, $FS_RWD, 65536, $FO_Seq)
            $len = $fs.Length
            $nHdr = $fs.Read($hdr8, 0, 8)
            $ioSw.Stop()
            if ($len -lt 64) { return 1 }
            if ($len -gt 524288000) { return 2 }
            if ($nHdr -lt 8) { return 1 }
            $fs.Position = 0
            if ($hdr8[0] -eq 0x50 -and $hdr8[1] -eq 0x4B -and $hdr8[2] -eq 0x03 -and $hdr8[3] -eq 0x04) { return (Test-ZipEntries $fs) }
            if ($hdr8[0] -eq 0xD0 -and $hdr8[1] -eq 0xCF -and $hdr8[2] -eq 0x11 -and $hdr8[3] -eq 0xE0) {
                if ($len -le 8388608) { return (Scan-Region $fs 0 $len $true $true) }
                $r1 = Scan-Region $fs 0 4194304 $false $true
                if ($r1 -ne 0) { return $r1 }
                return (Scan-Region $fs ($len - 4194304) 4194304 $false $false)
            }
            if ($len -le 16777216) { return (Scan-Region $fs 0 $len $false $false) }
            return 2
        } catch {
            return 2
        } finally {
            if ($ioSw.IsRunning) { $ioSw.Stop() }
            if ($null -ne $fs) { try { $fs.Dispose() } catch {} }
        }
    }

    try {
        $item = $null
        while (-not $state.CancelWorkers) {
            if ($workQueue.TryDequeue([ref]$item)) {
                $fPath = [string]$item.FilePath
                $fName = [string]$item.FileName
                $state.CurrentWorkerFile = $fName
                $ioSw.Reset(); $fileSw.Restart()
                $gateTaken = $false
                try {
                    $ioGate.Wait()
                    $gateTaken = $true
                    $zStatus = Test-WorkerZeroComStatus $fPath
                } finally {
                    if ($gateTaken) { [void]$ioGate.Release() }
                }
                $fileSw.Stop()
                if ($zStatus -eq 0 -or $zStatus -eq 1) {
                    $resultQueue.Enqueue([pscustomobject]@{
                        FilePath = $fPath
                        FileName = $fName
                        FileSig  = $item.FileSig
                        LowPath  = $item.LowPath
                        Result   = $(if ($zStatus -eq 0) { "ZERO_SAFE" } else { "ZERO_SKIP" })
                        WorkerId = $workerId
                        T2Ms     = $fileSw.Elapsed.TotalMilliseconds
                        IoMs     = $ioSw.Elapsed.TotalMilliseconds
                    })
                    continue
                }
                # Tep nghi ngo: chuyen sang lan COM cuc bo (khong bao gio cho Excel ngay trong Tier-2)
                if ($localExcelAvailable -and (-not $state.ComDisabled)) {
                    $deepQueue.Enqueue($item)
                } elseif ($bridgeRootDir) {
                    $bridgeQueue.Enqueue($item)
                } else {
                    $resultQueue.Enqueue([pscustomobject]@{
                        FilePath = $fPath; FileName = $fName; FileSig = $item.FileSig; LowPath = $item.LowPath
                        Result = "COM_SKIPPED"; WorkerId = $workerId
                    })
                }
            } elseif ($state.ProducerDone -and $workQueue.IsEmpty) {
                break
            } else {
                Start-Sleep -Milliseconds 15
            }
        }
    } finally {
        [System.GC]::Collect()
    }
}

# ------------------------------------------------------------------------------------------------
# v3.24.0 COM WORKER: Lan rieng cho Excel COM. Moi worker so huu toi da 1 EXCEL.EXE (theo doi PID).
# Luong chinh (Invoke-ComWatchdog) kill dung PID khi qua han & tu dong dong hop thoai chan (Activation...).
# ------------------------------------------------------------------------------------------------
$comWorkerScript = {
    param($workerId, $xlamPath, $bridgeRootDir, $deepQueue, $bridgeQueue, $resultQueue, $state)

    $script:excel = $null
    $script:wbAddin = $null
    $script:excelPid = 0
    $script:fsoWorker = $null
    $script:workerRestartSeq = 0
    $script:workerXlam = ""
    $workerGuardDir = Join-Path $env:APPDATA "KangatangGuard"
    if (-not (Test-Path -LiteralPath $workerGuardDir)) { try { New-Item -ItemType Directory -Path $workerGuardDir -Force | Out-Null } catch {} }

    function Set-ComPhase([string]$ph) {
        $state["ComPhaseStart_$workerId"] = [DateTime]::UtcNow.Ticks
        $state["ComPhase_$workerId"] = $ph
    }

    function Get-WorkerShortPath([string]$p) {
        if ([string]::IsNullOrEmpty($p) -or $p.Length -lt 215) { return $p }
        try {
            if ($null -eq $script:fsoWorker) { $script:fsoWorker = New-Object -ComObject Scripting.FileSystemObject }
            $lp = if ($p.StartsWith("\\?\")) { $p } elseif ($p.StartsWith("\\")) { "\\?\UNC\" + $p.Substring(2) } else { "\\?\" + $p }
            $sh = $script:fsoWorker.GetFile($lp).ShortPath
            if ($sh) {
                if ($sh.StartsWith("\\?\UNC\")) { return "\\" + $sh.Substring(8) }
                if ($sh.StartsWith("\\?\")) { return $sh.Substring(4) }
                return $sh
            }
        } catch {}
        return $p
    }

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

    function Register-ExcelPid([int]$excelProcId) {
        $script:excelPid = $excelProcId
        $state["ComPid_$workerId"] = $excelProcId
        if ($excelProcId -gt 0) { $state.ComPidSet[$excelProcId] = $workerId }
    }

    function Stop-WorkerExcelInstance {
        Set-ComPhase "QUIT"
        if ($null -ne $script:excel -and $null -ne $script:wbAddin) {
            try { $script:excel.Run("'$($script:wbAddin.Name)'!PrepareWorkerShutdown") | Out-Null } catch {}
        }
        if ($null -ne $script:wbAddin) {
            try { $script:wbAddin.Saved = $true } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:wbAddin) } catch {}
            $script:wbAddin = $null
        }
        if ($null -ne $script:excel) {
            try { $script:excel.DisplayAlerts = $false } catch {}
            try { $script:excel.IgnoreRemoteRequests = $false } catch {}
            try { $script:excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel) } catch {}
            $script:excel = $null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        # Dam bao khong con EXCEL.EXE mo coi: cho 5s, qua han thi kill dung PID cua worker nay
        if ($script:excelPid -gt 0) {
            try {
                $pr = Get-Process -Id $script:excelPid -ErrorAction SilentlyContinue
                if ($pr -and (-not $pr.WaitForExit(5000))) { Stop-Process -Id $script:excelPid -Force -ErrorAction SilentlyContinue }
            } catch {}
            try { $state.ComPidSet.Remove($script:excelPid) } catch {}
        }
        Register-ExcelPid 0
        if ($script:workerXlam -and ($script:workerXlam -ne $xlamPath) -and (Test-Path -LiteralPath $script:workerXlam)) {
            try { Remove-Item -LiteralPath $script:workerXlam -Force -ErrorAction SilentlyContinue } catch {}
        }
        Set-ComPhase "IDLE"
    }

    function Restart-LocalWorker {
        Stop-WorkerExcelInstance
        $spawnMutex = $null
        $hasSpawnLock = $false
        try {
            # Mutex CHI bao phan sinh tien trinh + xac dinh PID (khong giu khi Workbooks.Open / InitWorkerMode)
            try {
                $spawnMutex = New-Object System.Threading.Mutex($false, "Local\KangatangGuard_ExcelSpawn_Mutex")
                try { $hasSpawnLock = $spawnMutex.WaitOne(30000, $false) } catch [System.Threading.AbandonedMutexException] { $hasSpawnLock = $true }
                Clear-WorkerCrashRecovery
                $script:workerRestartSeq++
                $script:workerXlam = Join-Path $workerGuardDir "KangatangGuard_Worker_$($state.MasterPid)_$($workerId)_$($script:workerRestartSeq).xlam"
                try {
                    Copy-Item -LiteralPath $xlamPath -Destination $script:workerXlam -Force -ErrorAction SilentlyContinue
                    Unblock-File -LiteralPath $script:workerXlam -ErrorAction SilentlyContinue
                } catch {}
                if (-not (Test-Path -LiteralPath $script:workerXlam)) { $script:workerXlam = $xlamPath }

                $preIds = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
                $state["ComPreSnap_$workerId"] = $preIds
                Set-ComPhase "INIT"
                $script:excel = New-Object -ComObject Excel.Application
                $newProcs = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Where-Object { $preIds -notcontains $_.Id })
                if ($newProcs.Count -gt 1) {
                    $autoIds = @(Get-WmiObject -Class Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue | Where-Object { ([string]$_.CommandLine) -match '(?i)/automation|-embedding' } | ForEach-Object { [int]$_.ProcessId })
                    $newProcs = @($newProcs | Where-Object { $autoIds -contains $_.Id })
                }
                if ($newProcs.Count -ge 1) {
                    Register-ExcelPid ([int](($newProcs | Sort-Object StartTime -Descending | Select-Object -First 1).Id))
                }
            } finally {
                if ($hasSpawnLock -and $null -ne $spawnMutex) { try { $spawnMutex.ReleaseMutex() } catch {} }
                if ($null -ne $spawnMutex) { try { $spawnMutex.Dispose() } catch {} }
            }

            $script:excel.Visible = $false
            try { $script:excel.AutomationSecurity = 1 } catch {}
            try { $script:excel.ShowWindowsInTaskbar = $false } catch {}
            try { $script:excel.AutoRecover.Enabled = $false } catch {}
            $script:excel.DisplayAlerts = $false
            $script:excel.EnableEvents = $false
            try { $script:excel.IgnoreRemoteRequests = $true } catch {}
            Clear-WorkerDdeRegistryBit
            $script:wbAddin = $script:excel.Workbooks.Open($script:workerXlam)
            $script:excel.Run("'$($script:wbAddin.Name)'!InitWorkerMode") | Out-Null
            Clear-WorkerDdeRegistryBit
            Set-ComPhase "IDLE"
        } catch {
            Stop-WorkerExcelInstance
            throw
        }
    }

    function Send-ToBridgeOrResult($it, [string]$fallbackResult) {
        if ($bridgeRootDir) {
            $bridgeQueue.Enqueue($it)
        } else {
            $resultQueue.Enqueue([pscustomobject]@{
                FilePath = [string]$it.FilePath; FileName = [string]$it.FileName; FileSig = $it.FileSig; LowPath = $it.LowPath
                Result = $fallbackResult; WorkerId = (100 + $workerId)
            })
        }
    }

    try {
        Set-ComPhase "IDLE"
        $processedComCount = 0
        $item = $null
        while (-not $state.CancelWorkers) {
            if ($deepQueue.TryDequeue([ref]$item)) {
                $fPath = [string]$item.FilePath
                $fName = [string]$item.FileName
                if ($state.ComDisabled) { Send-ToBridgeOrResult $item "COM_SKIPPED"; continue }

                $comReady = ($null -ne $script:excel) -and ($null -ne $script:wbAddin)
                if (-not $comReady) {
                    for ($initTry = 1; $initTry -le 2; $initTry++) {
                        try {
                            $state["ComKilled_$workerId"] = $false
                            Restart-LocalWorker
                            $comReady = $true
                            $state.ComInitFailStreak = 0
                            break
                        } catch {
                            Start-Sleep -Milliseconds (800 * $initTry)
                        }
                    }
                }
                if (-not $comReady) {
                    $state.ComInitFailStreak = [int]$state.ComInitFailStreak + 1
                    # Chi chuyen CA LAN COM sang Bridge khi co Bridge (Tier-2 khong bi anh huong)
                    if (([int]$state.ComInitFailStreak -ge 6) -and $bridgeRootDir) { $state.ComDisabled = $true }
                    Send-ToBridgeOrResult $item "WORKER_INIT_ERR"
                    continue
                }

                $state.CurrentWorkerFile = "[COM-W$workerId] " + $fName
                $comTargetPath = Get-WorkerShortPath $fPath
                $res = ""
                $processedComCount++
                $state["ComKilled_$workerId"] = $false
                Set-ComPhase "SCAN"
                try {
                    $script:excel.Visible = $false
                    $res = [string]$script:excel.Run("'$($script:wbAddin.Name)'!ScanSingleFileForWorker", $comTargetPath)
                    if ($res.StartsWith("PURGE_NEEDED|")) {
                        $detail = $res.Substring(13)
                        $purgeRes = "MORE_PURGE"
                        for ($attempt = 1; $attempt -le 5; $attempt++) {
                            try { $script:excel.Run("'$($script:wbAddin.Name)'!WorkerDeleteOneModuleSheet", $comTargetPath) | Out-Null } catch {}
                            try {
                                $purgeRes = [string]$script:excel.Run("'$($script:wbAddin.Name)'!WorkerFinishCleanAndClose", $comTargetPath)
                            } catch {
                                $purgeRes = "INFECTED_ERR|RPC error"
                                break
                            }
                            if ($purgeRes -ne "MORE_PURGE") { break }
                        }
                        if ($purgeRes.StartsWith("CLEANED|")) { $res = "CLEANED_PURGE|$detail" } else { $res = $purgeRes }
                    }
                    Set-ComPhase "IDLE"
                    if ($state["ComKilled_$workerId"]) { throw "Excel killed by watchdog" }
                    if ($processedComCount -ge 2500) {
                        try { Restart-LocalWorker } catch {}
                        $processedComCount = 0
                    }
                } catch {
                    $wasKilled = [bool]$state["ComKilled_$workerId"]
                    $state["ComKilled_$workerId"] = $false
                    try { Stop-WorkerExcelInstance } catch {}
                    $processedComCount = 0
                    if ($wasKilled) {
                        # Qua han / bi hop thoai chan -> chuyen DUNG tep nay sang Bridge 223.176 (khong treo lan COM)
                        Send-ToBridgeOrResult $item "COM_TIMEOUT"
                        continue
                    }
                    $res = "WORKER_RECOVERED|$($_.Exception.Message)"
                }

                $resultQueue.Enqueue([pscustomobject]@{
                    FilePath = $fPath
                    FileName = $fName
                    FileSig  = $item.FileSig
                    LowPath  = $item.LowPath
                    Result   = $res
                    WorkerId = (100 + $workerId)
                })
            } elseif ($state.Tier2Done -and $deepQueue.IsEmpty) {
                break
            } else {
                Start-Sleep -Milliseconds 40
            }
        }
    } finally {
        try { Stop-WorkerExcelInstance } catch {}
        Clear-WorkerCrashRecovery
        Set-ComPhase "DONE"
    }
}

# Worker Cau Noi Tu Xa (Remote Bridge Worker): Xu ly cac tep trong $bridgeWorkQueue (may khong cai Excel / Excel bi chan / qua han)
$remoteBridgeWorkerScript = {
    param($bridgeRootDir, $bridgeQueue, $resultQueue, $state)
    $reqQueueDir = Join-Path $bridgeRootDir "Remote_COM_Queue"
    if (-not (Test-Path -LiteralPath $reqQueueDir)) {
        try { New-Item -ItemType Directory -Path $reqQueueDir -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    $reqFile = Join-Path $reqQueueDir "remote_com_requests.jsonl"
    $resFile = Join-Path $reqQueueDir "remote_com_results.jsonl"
    $utf8NoBomLocal = New-Object System.Text.UTF8Encoding($false)
    $pendingMap = @{}
    $lastReadLineCount = 0
    if (Test-Path -LiteralPath $resFile) {
        try { $lastReadLineCount = ([System.IO.File]::ReadAllLines($resFile, $utf8NoBomLocal)).Count } catch {}
    }

    while (-not $state.CancelWorkers) {
        $batchLines = New-Object System.Collections.Generic.List[string]
        $item = $null
        while ($batchLines.Count -lt 25 -and $bridgeQueue.TryDequeue([ref]$item)) {
            $fPath = [string]$item.FilePath
            $uncPath = $fPath
            if ($fPath -match '(?i)^.*?\\file_shared(\\.+)$') {
                $uncPath = "\\192.168.223.7\file_shared" + $Matches[1]
            } elseif ($fPath -match '(?i)^.*?\\fileshared_2(\\.+)$') {
                $uncPath = "\\192.168.223.7\fileshared_2" + $Matches[1]
            } elseif ($fPath -match '(?i)^.*?\\vietnam(\\.+)$') {
                $uncPath = "\\192.168.223.7\file_shared\vietnam" + $Matches[1]
            }
            $reqId = [Guid]::NewGuid().ToString("N")
            $pendingMap[$reqId] = $item
            $state.CurrentWorkerFile = "[Bridge->223.176] " + [string]$item.FileName
            $reqObj = [ordered]@{
                ReqId    = $reqId
                UncPath  = $uncPath
                OrigPath = $fPath
                FileName = [string]$item.FileName
                FileSig  = [string]$item.FileSig
                LowPath  = [string]$item.LowPath
            }
            $batchLines.Add(($reqObj | ConvertTo-Json -Compress))
        }
        if ($batchLines.Count -gt 0) {
            $chunk = ($batchLines -join "`r`n") + "`r`n"
            for ($rTry = 1; $rTry -le 5; $rTry++) {
                try {
                    [System.IO.File]::AppendAllText($reqFile, $chunk, $utf8NoBomLocal)
                    break
                } catch {
                    Start-Sleep -Milliseconds 80
                }
            }
        }

        if (($pendingMap.Count -gt 0) -and (Test-Path -LiteralPath $resFile)) {
            try {
                $allResLines = [System.IO.File]::ReadAllLines($resFile, $utf8NoBomLocal)
                if ($allResLines.Count -gt $lastReadLineCount) {
                    for ($idxR = $lastReadLineCount; $idxR -lt $allResLines.Count; $idxR++) {
                        $ln = $allResLines[$idxR]
                        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                        try {
                            $rObj = $ln | ConvertFrom-Json
                            if ($rObj -and $rObj.ReqId -and $pendingMap.ContainsKey([string]$rObj.ReqId)) {
                                $origItem = $pendingMap[[string]$rObj.ReqId]
                                $pendingMap.Remove([string]$rObj.ReqId)
                                $resultQueue.Enqueue([pscustomobject]@{
                                    FilePath = $origItem.FilePath
                                    FileName = $origItem.FileName
                                    FileSig  = $origItem.FileSig
                                    LowPath  = $origItem.LowPath
                                    Result   = [string]$rObj.Result
                                    WorkerId = [int]$rObj.WorkerId
                                })
                            }
                        } catch {}
                    }
                    $lastReadLineCount = $allResLines.Count
                }
            } catch {}
        }
        $state.BridgePending = $pendingMap.Count

        if ($state.ComDone -and $bridgeQueue.IsEmpty -and ($pendingMap.Count -eq 0)) {
            break
        }
        Start-Sleep -Milliseconds 150
    }
}

$script:runspacePool = $null
$script:workerJobs   = New-Object System.Collections.Generic.List[pscustomobject]
$script:poolStarted  = $false
$script:bridgeFallbackStarted = $false

function New-WorkerJob([string]$kind, [int]$wId) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $script:runspacePool
    if ($kind -eq "T2") {
        [void]$ps.AddScript($tier2WorkerScript).AddArgument($wId).AddArgument($script:localExcelAvailable).AddArgument($script:bridgeRootDir).AddArgument($comWorkQueue).AddArgument($deepComQueue).AddArgument($bridgeWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState).AddArgument($script:ioGate)
    } elseif ($kind -eq "COM") {
        [void]$ps.AddScript($comWorkerScript).AddArgument($wId).AddArgument($xlamPath).AddArgument($script:bridgeRootDir).AddArgument($deepComQueue).AddArgument($bridgeWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
    } else {
        [void]$ps.AddScript($remoteBridgeWorkerScript).AddArgument($script:bridgeRootDir).AddArgument($bridgeWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
    }
    $handle = $ps.BeginInvoke()
    return [pscustomobject]@{ PowerShell = $ps; Handle = $handle; Id = $wId; Kind = $kind }
}

function Ensure-WorkerPoolStarted {
    if (-not $script:poolStarted) {
        $script:poolStarted = $true
        $script:runspacePool = [runspacefactory]::CreateRunspacePool(1, ($workerCount + $comWorkerCount + 1))
        $script:runspacePool.ApartmentState = [System.Threading.ApartmentState]::STA
        $script:runspacePool.Open()
        for ($wId = 1; $wId -le $workerCount; $wId++) { $script:workerJobs.Add((New-WorkerJob "T2" $wId)) }
        if ($script:localExcelAvailable -and $xlamPath) {
            for ($cId = 1; $cId -le $comWorkerCount; $cId++) { $script:workerJobs.Add((New-WorkerJob "COM" $cId)) }
        }
    }

    if ((-not $bridgeWorkQueue.IsEmpty) -and $script:bridgeRootDir -and (-not $script:bridgeFallbackStarted)) {
        $script:bridgeFallbackStarted = $true
        $script:workerJobs.Add((New-WorkerJob "BRIDGE" 99))
    }

    # Watchdog: Tu dong khoi phuc Worker neu bat ky luong nao bi ngat dot ngot khi van con viec
    if (-not $sharedState.CancelWorkers) {
        for ($i = 0; $i -lt $script:workerJobs.Count; $i++) {
            $job = $script:workerJobs[$i]
            if (-not $job.Handle.IsCompleted) { continue }
            $needRestart = $false
            if ($job.Kind -eq "T2")  { $needRestart = -not ($sharedState.ProducerDone -and $comWorkQueue.IsEmpty) }
            if ($job.Kind -eq "COM") { $needRestart = -not ($sharedState.Tier2Done -and $deepComQueue.IsEmpty) }
            if ($needRestart) {
                try { $job.PowerShell.Dispose() } catch {}
                $script:workerJobs[$i] = New-WorkerJob $job.Kind $job.Id
            }
        }
    }
}

function Test-JobsCompleted([string]$kind) {
    foreach ($j in $script:workerJobs) {
        if ($j.Kind -eq $kind -and (-not $j.Handle.IsCompleted)) { return $false }
    }
    return $true
}

# ------------------------------------------------------------------------------------------------
# v3.24.0 P0: Watchdog Excel + tu dong dong hop thoai chan (UIAutomation - Assembly .NET ky so san co,
# KHONG bien dich C#, KHONG SendKeys -> khong chiem ban phim/chuot phien RDP)
# ------------------------------------------------------------------------------------------------
$script:lastWatchdogTicks = 0
$script:comKillCount      = 0
$script:dialogCloseCount  = 0
$script:uiaState          = 0   # 0 = chua nap, 1 = san sang, -1 = khong kha dung
$script:btnCloseNames     = @("Close", "Cancel", ([string]([char]0xB2EB) + [char]0xAE30), ([string]([char]0xCDE8) + [char]0xC18C))

function Get-ComCandidatePids([int]$cId) {
    $pids = New-Object 'System.Collections.Generic.List[int]'
    $known = [int]$sharedState["ComPid_$cId"]
    if ($known -gt 0) { $pids.Add($known); return $pids }
    # Dang treo trong New-Object (chua biet PID): chi chon EXCEL tu dong hoa moi sinh sau snapshot, chua thuoc worker nao
    $pre = @($sharedState["ComPreSnap_$cId"])
    $phaseStart = [long]$sharedState["ComPhaseStart_$cId"]
    try {
        $autoIds = @(Get-WmiObject -Class Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue | Where-Object { ([string]$_.CommandLine) -match '(?i)/automation|-embedding' } | ForEach-Object { [int]$_.ProcessId })
        foreach ($ap in $autoIds) {
            if ($pre -contains $ap) { continue }
            if ($sharedState.ComPidSet.ContainsKey($ap)) { continue }
            $pr = Get-Process -Id $ap -ErrorAction SilentlyContinue
            if ($pr -and ($pr.StartTime.ToUniversalTime().Ticks -ge ($phaseStart - 20000000))) { $pids.Add($ap) }
        }
    } catch {}
    return $pids
}

function Close-ScannerExcelDialogs($excelPids) {
    if ($script:uiaState -eq 0) {
        try {
            Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
            Add-Type -AssemblyName UIAutomationTypes -ErrorAction Stop
            $script:uiaState = 1
        } catch { $script:uiaState = -1 }
    }
    if ($script:uiaState -ne 1) { return 0 }
    $closed = 0
    foreach ($xp in $excelPids) {
        try {
            $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$xp)
            $tops = [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
            $winCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
            $wins = New-Object System.Collections.Generic.List[object]
            foreach ($tw in $tops) {
                $tcls = [string]$tw.Current.ClassName
                if ($tcls -eq "XLMAIN") {
                    # Hop thoai modal (Activation Wizard, File-in-use, Open...) duoc UIA bieu dien la Window con cua XLMAIN
                    foreach ($cw in $tw.FindAll([System.Windows.Automation.TreeScope]::Children, $winCond)) { $wins.Add($cw) }
                } elseif ($tcls -notlike "MSO_BORDEREFFECT*") {
                    $wins.Add($tw)
                }
            }
            foreach ($w in $wins) {
                $cls = [string]$w.Current.ClassName
                if ($cls -eq "XLMAIN") { continue }
                $nm = ConvertTo-DisplayAscii ([string]$w.Current.Name)
                $ok = $false
                try {
                    $wp = $w.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
                    $wp.Close()
                    $ok = $true
                } catch {}
                if (-not $ok) {
                    try {
                        $btns = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)))
                        foreach ($b in $btns) {
                            $bn = ([string]$b.Current.Name).Replace("&", "")
                            $hit = $false
                            foreach ($cn in $script:btnCloseNames) { if ($bn.StartsWith($cn)) { $hit = $true; break } }
                            if ($hit) {
                                ($b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)).Invoke()
                                $ok = $true
                                break
                            }
                        }
                    } catch {}
                }
                if ($ok) {
                    $closed++
                    $script:dialogCloseCount++
                    Add-DetailAuditEvent "COM_DIALOG_CLOSED" "PID=$xp" 0 "(DIALOG)" "Tu dong dong hop thoai chan Excel: class=$cls title=$nm"
                }
            }
        } catch {}
    }
    return $closed
}

function Invoke-ComWatchdog {
    $nowT = [DateTime]::UtcNow.Ticks
    if (($nowT - $script:lastWatchdogTicks) -lt 10000000) { return }
    $script:lastWatchdogTicks = $nowT
    if (-not $script:localExcelAvailable) { return }
    $script:scannerExcelCount = Get-ScannerExcelCount
    for ($cId = 1; $cId -le $comWorkerCount; $cId++) {
        $phase = [string]$sharedState["ComPhase_$cId"]
        if ((-not $phase) -or $phase -eq "IDLE" -or $phase -eq "DONE") { continue }
        $elapsedSec = ($nowT - [long]$sharedState["ComPhaseStart_$cId"]) / 10000000.0
        $limit = switch ($phase) { "INIT" { 90 } "SCAN" { 180 } "QUIT" { 15 } default { 120 } }
        if ($elapsedSec -ge 3) {
            $cands = Get-ComCandidatePids $cId
            if ($cands.Count -gt 0) { [void](Close-ScannerExcelDialogs $cands) }
        }
        if ($elapsedSec -gt $limit) {
            $cands = Get-ComCandidatePids $cId
            foreach ($kp in $cands) {
                try { Stop-Process -Id $kp -Force -ErrorAction SilentlyContinue } catch {}
                try { $sharedState.ComPidSet.Remove($kp) } catch {}
            }
            $sharedState["ComKilled_$cId"] = $true
            $sharedState["ComPhaseStart_$cId"] = $nowT
            $script:comKillCount++
            Add-DetailAuditEvent "COM_WATCHDOG_KILL" ("PID=" + ($cands -join ",")) (100 + $cId) "(TIMEOUT)" "COM-W$cId qua han ${limit}s o pha $phase -> kill dung PID Excel cua scanner"
        }
    }
}

function Stop-AllScannerExcel {
    try {
        foreach ($k in @($sharedState.ComPidSet.Keys)) {
            try { Stop-Process -Id ([int]$k) -Force -ErrorAction SilentlyContinue } catch {}
        }
        $sharedState.ComPidSet.Clear()
    } catch {}
}

function Get-ScannerExcelCount {
    $n = 0
    try {
        foreach ($k in @($sharedState.ComPidSet.Keys)) {
            if (Get-Process -Id ([int]$k) -ErrorAction SilentlyContinue) { $n++ }
        }
    } catch {}
    return $n
}


# 7. Bien trang thai va Ham ve Bang Dieu Khien Co Dinh (Non-Scrolling Live Dashboard UI)
$script:initialResumeCount = 0
$script:totalDiscovered = 0
$script:totalScanned    = 0
$script:totalSkipped    = 0
$script:totalZeroCom    = 0
$script:totalDeepCom    = 0
$script:totalSafe       = 0
$script:totalInfected   = $prevSessionCleaned
$script:totalCleaned    = $prevSessionCleaned
$script:totalErrors     = 0
$script:totalChangedRescan = 0
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
} catch {}

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

    $elapsedSec = [Math]::Max(0.5, $stopwatch.Elapsed.TotalSeconds)
    $sessionNewScanned = [Math]::Max(0, $script:totalScanned - [int]$script:initialResumeCount)
    $speed = if ($sessionNewScanned -gt 0) { [Math]::Round($sessionNewScanned / $elapsedSec, 1) } else { 0.0 }

    $etaStr = "--h--p"
    if ($isCompleted) {
        $etaStr = "00p00s"
    } elseif ($speed -gt 0 -and $remFiles -gt 0 -and $sessionNewScanned -ge 10) {
        $etaSec = [int][Math]::Ceiling($remFiles / $speed)
        if ($etaSec -ge 3600) {
            $etaStr = "{0:D2}h{1:D2}p" -f [int][Math]::Floor($etaSec / 3600), [int][Math]::Floor(($etaSec % 3600) / 60)
        } else {
            $etaStr = "{0:D2}p{1:D2}s" -f [int][Math]::Floor($etaSec / 60), ($etaSec % 60)
        }
    } elseif ($remFiles -eq 0 -and $sharedState.CountDone) {
        $etaStr = "00p00s"
    }

    try {
        $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion [PID:$PID | $workerCount Workers] | Tien do: $($pctDone.ToString('0.0'))% (Con lai: $($pctRem.ToString('0.0'))%) | Da quet: $($script:totalScanned)/$totalTarget | Da diet: $($script:totalCleaned)"
    } catch {}

    # Dat con tro an toan: KHONG BAO GIO tat vinh vien $script:canSetCursor khi nguoi dung keo/thu nho cua so Console!
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
        $safeTop = [Math]::Max(0, [Math]::Min([int]$script:dashboardTop, [Math]::Max(0, [Console]::BufferHeight - 30)))
        [Console]::SetCursorPosition(0, $safeTop)
    } catch {
        if (-not $isCompleted) { return }
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
    $countStateStr = if ($sharedState.CountDone -or $isCompleted) { "$totalFolders thu muc | Toc do: $speed/s" } else { "Uoc tinh tu lich su | $speed/s" }
    $safePctStr = if ($script:totalScanned -gt 0) { "Ty le an toan: $([Math]::Round(($script:totalSafe / [double]$script:totalScanned) * 100.0, 1))%" } else { "Dang kiem tra..." }
    $t2Str  = "{0:0.0} ms/tep (IO {1:0.0} ms)" -f $script:t2MsAvg, $script:ioMsAvg
    $comStr = "COM: {0} | Bridge: {1}" -f $deepComQueue.Count, ($bridgeWorkQueue.Count + [int]$sharedState.BridgePending)
    $laneMode = if (-not $script:localExcelAvailable) { "BRIDGE" } elseif ($sharedState.ComDisabled) { "LOCAL->BRIDGE" } else { "LOCAL" }
    $laneStr = "Excel: {0}/{1} | Dong hop thoai: {2} | Kill: {3} | Lan: {4}" -f $script:scannerExcelCount, $comWorkerCount, $script:dialogCloseCount, $script:comKillCount, $laneMode

    Write-FixedLine (" | [3] BANG KET QUA QUET TRUC TUYEN (LIVE TABLE) |       SO LUONG | CHI TIET / TRANG THAI          |") ([System.ConsoleColor]::Cyan)
    Write-FixedLine (" +-----------------------------------------------+----------------+--------------------------------+") ([System.ConsoleColor]::DarkCyan)
    Write-FixedLine (" | 1. Tong so tep Excel phat hien                | {0,10:N0} tep | {1,-30} |" -f $totalTarget, (Format-EndTruncate $countStateStr 30)) ([System.ConsoleColor]::White)
    Write-FixedLine (" | 2. Tong so tep da quet xong                   | {0,10:N0} tep | {1,-30} |" -f $script:totalScanned, ("Xong: {0:0.0}% | Con lai: {1:0.0}%" -f $pctDone, $pctRem)) ([System.ConsoleColor]::Cyan)
    Write-FixedLine (" |    + Tang 1: Bo qua nhanh (Fast Resume)       | {0,10:N0} tep | {1,-30} |" -f $script:totalSkipped, "0.01 ms/tep (Da sach tu truoc)") ([System.ConsoleColor]::Gray)
    $lblT2 = ("    + Tang 2: Loc tinh Zero-COM ({0} Workers)" -f $workerCount).PadRight(47)
    $lblT3 = ("    + Tang 3: Quet sau Excel COM ({0} Workers)" -f $comWorkerCount).PadRight(47)
    Write-FixedLine (" |$lblT2| {0,10:N0} tep | {1,-30} |" -f $script:totalZeroCom, (Format-EndTruncate $t2Str 30)) ([System.ConsoleColor]::Gray)
    Write-FixedLine (" |$lblT3| {0,10:N0} tep | {1,-30} |" -f $script:totalDeepCom, (Format-EndTruncate $comStr 30)) ([System.ConsoleColor]::Gray)
    Write-FixedLine (" |      $(Format-EndTruncate $laneStr 92)") $(if ($script:comKillCount -gt 0 -or $sharedState.ComDisabled) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::DarkGray })
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

$script:t2MsAvg = 0.0
$script:ioMsAvg = 0.0
$script:scannerExcelCount = 0
$script:lastHeartbeatTicks = 0
$script:runLogPath = Join-Path $logDir "scanner_run.log"

# v3.24.0: Nhip bao tri (watchdog Excel + dem Excel cua scanner + ghi vet heartbeat 30s vao scanner_run.log)
function Invoke-MaintenanceTick {
    try { Invoke-ComWatchdog } catch {}
    $nowT = [DateTime]::UtcNow.Ticks
    if (($nowT - $script:lastHeartbeatTicks) -lt 300000000) { return }
    $script:lastHeartbeatTicks = $nowT
    $script:scannerExcelCount = Get-ScannerExcelCount
    try {
        if ((Test-Path -LiteralPath $script:runLogPath) -and ((Get-Item -LiteralPath $script:runLogPath).Length -gt 5MB)) {
            Move-Item -LiteralPath $script:runLogPath -Destination ($script:runLogPath + ".old") -Force -ErrorAction SilentlyContinue
        }
        $memMb = [Math]::Round(([System.Diagnostics.Process]::GetCurrentProcess().PrivateMemorySize64) / 1MB, 0)
        $hb = "{0} | PID={1} | v{2} | Scanned={3} | ZeroCom={4} | DeepCom={5} | Q_T2={6} | Q_COM={7} | Q_BRIDGE={8}+{9} | ExcelOwned={10}/{11} | DialogsClosed={12} | Kills={13} | ComDisabled={14} | T2ms={15:0.0} | IOms={16:0.0} | MemMB={17}`r`n" -f `
            (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $PID, $AppVersion, $script:totalScanned, $script:totalZeroCom, $script:totalDeepCom,
            $comWorkQueue.Count, $deepComQueue.Count, $bridgeWorkQueue.Count, [int]$sharedState.BridgePending,
            $script:scannerExcelCount, $comWorkerCount, $script:dialogCloseCount, $script:comKillCount, [bool]$sharedState.ComDisabled,
            $script:t2MsAvg, $script:ioMsAvg, $memMb
        [System.IO.File]::AppendAllText($script:runLogPath, $hb, $utf8NoBom)
    } catch {}
}

function Drain-WorkerResults {
    Invoke-MaintenanceTick
    $resItem = $null
    $updated = $false
    $forceThreatFlush = $false
    while ($comResultQueue.TryDequeue([ref]$resItem)) {
        $updated = $true
        $r = [string]$resItem.Result
        $fName = [string]$resItem.FileName
        $fPath = [string]$resItem.FilePath
        $wId   = $resItem.WorkerId
        $script:currentFile = $fName
        if ($null -ne $resItem.T2Ms) {
            $script:t2MsAvg = (0.97 * $script:t2MsAvg) + (0.03 * [double]$resItem.T2Ms)
            $script:ioMsAvg = (0.97 * $script:ioMsAvg) + (0.03 * [double]$resItem.IoMs)
        }

        if ($r -eq "ZERO_SAFE") {
            $script:totalZeroCom++
            $script:totalScanned++
            $script:totalSafe++
            Update-DeptStat $fPath "SAFE"
            $script:logBuffer.Add("$($resItem.LowPath)|$($resItem.FileSig)")
            continue
        } elseif ($r -eq "ZERO_SKIP") {
            $script:totalZeroCom++
            $script:totalScanned++
            $script:totalErrors++
            Update-DeptStat $fPath "ERROR"
            $script:logBuffer.Add("$($resItem.LowPath)|$($resItem.FileSig)")
            continue
        }

        $script:totalDeepCom++
        if ($r -eq "SAFE") {
            $script:totalScanned++
            $script:totalSafe++
            Update-DeptStat $fPath "SAFE"
            $script:logBuffer.Add("$($resItem.LowPath)|$($resItem.FileSig)")
        } elseif ($r.StartsWith("CLEANED|") -or $r.StartsWith("CLEANED_PURGE|")) {
            $forceThreatFlush = $true
            $script:totalScanned++
            $script:totalInfected++
            $script:totalCleaned++
            Update-DeptStat $fPath "CLEANED"
            $isPurge = $r.StartsWith("CLEANED_PURGE|")
            $threatFamily = if ($isPurge) { "Excel5_ModuleSheet_Hidden" } else { "VBA_Macro_Kangatang" }
            $rawSub = if ($isPurge) { "Module-Sheet Excel 5.0: " + $r.Substring(14) } else { $r.Substring(8) }
            $lastSavedBy = "(Khong ro)"
            $detail = $rawSub
            $idxAuth = $rawSub.IndexOf("|||")
            if ($idxAuth -ge 0) {
                $detail = $rawSub.Substring(0, $idxAuth)
                $lastSavedBy = $rawSub.Substring($idxAuth + 3).Trim()
                if (-not $lastSavedBy) { $lastSavedBy = "(Khong ro)" }
            }
            $isReInfected = ($null -ne $script:everScannedPaths) -and $script:everScannedPaths.Contains([string]$resItem.LowPath)
            if ($isReInfected) { $script:totalReInfected++ }
            $discoveryType = if ($isReInfected) { "RE_INFECTED" } else { "FIRST_TIME_CLEAN" }

            $nowDt = Get-Date
            $ts = $nowDt.ToString("HH:mm:ss")
            $tsFull = $nowDt.ToString("yyyy-MM-dd HH:mm:ss")
            Add-RecentThreatLog "> [$ts] [DA DIET] $fName ($detail | Luu cuoi: $lastSavedBy)"
            $deptName = Get-TopDepartmentName $fPath
            $folderDir = Get-SafeParentDir $fPath
            $extStr = Get-SafeExtension $fName
            $sizeKb = 0
            try {
                $partsSig = ([string]$resItem.FileSig).Split('|')
                if ($partsSig.Count -ge 1) { $sizeKb = [Math]::Round(([long]$partsSig[0]) / 1024.0, 1) }
            } catch {}

            [void]$script:infectedRecords.Add([pscustomobject]@{
                Time          = $tsFull
                TimeShort     = $ts
                Department    = $deptName
                FileName      = $fName
                FolderPath    = $folderDir
                FullPath      = $fPath
                Extension     = $extStr
                SizeKB        = $sizeKb
                ThreatFamily  = $threatFamily
                ThreatDetail  = "$detail [LastSavedBy=$lastSavedBy] [Discovery=$discoveryType]"
                LastSavedBy   = $lastSavedBy
                DiscoveryType = $discoveryType
                Action        = "CLEANED_AND_QUARANTINED"
                WorkerId      = $wId
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
                (Format-CsvField "$detail [LastSavedBy=$lastSavedBy] [Discovery=$discoveryType]"),
                (Format-CsvField "CLEANED_AND_QUARANTINED"),
                $wId
            $script:ledgerBuffer.Add($csvRow)
            Add-DetailAuditEvent "DETECTED_CLEANED" $fPath $wId $threatFamily "Da diet & cach ly thanh cong ($detail | LastSavedBy=$lastSavedBy | Discovery=$discoveryType)"

            try {
                $shortAfter = Get-SafeShortPath $fPath
                $fAfter = New-Object System.IO.FileInfo($shortAfter)
                $newSig = "$($fAfter.Length)|$($fAfter.LastWriteTime.ToString('yyyyMMddHHmmss'))"
                $script:logBuffer.Add("$($resItem.LowPath)|$newSig")
            } catch {}
        } elseif ($r.StartsWith("INFECTED_")) {
            $forceThreatFlush = $true
            $script:totalScanned++
            $script:totalInfected++
            $script:totalErrors++
            Update-DeptStat $fPath "INFECTED_LOCKED"
            $lastSavedBy = "(Khong ro)"
            $detail = $r
            $idxAuth = $r.IndexOf("|||")
            if ($idxAuth -ge 0) {
                $detail = $r.Substring(0, $idxAuth)
                $lastSavedBy = $r.Substring($idxAuth + 3).Trim()
                if (-not $lastSavedBy) { $lastSavedBy = "(Khong ro)" }
            }
            $isReInfected = ($null -ne $script:everScannedPaths) -and $script:everScannedPaths.Contains([string]$resItem.LowPath)
            if ($isReInfected) { $script:totalReInfected++ }
            $discoveryType = if ($isReInfected) { "RE_INFECTED" } else { "FIRST_TIME_CLEAN" }

            $nowDt = Get-Date
            $ts = $nowDt.ToString("HH:mm:ss")
            $tsFull = $nowDt.ToString("yyyy-MM-dd HH:mm:ss")
            Add-RecentThreatLog "! [$ts] [KHOA/READONLY] $fName ($detail | Luu cuoi: $lastSavedBy)"
            $deptName = Get-TopDepartmentName $fPath
            $folderDir = Get-SafeParentDir $fPath
            $extStr = Get-SafeExtension $fName
            $sizeKb = 0
            try {
                $partsSig = ([string]$resItem.FileSig).Split('|')
                if ($partsSig.Count -ge 1) { $sizeKb = [Math]::Round(([long]$partsSig[0]) / 1024.0, 1) }
            } catch {}

            [void]$script:infectedRecords.Add([pscustomobject]@{
                Time          = $tsFull
                TimeShort     = $ts
                Department    = $deptName
                FileName      = $fName
                FolderPath    = $folderDir
                FullPath      = $fPath
                Extension     = $extStr
                SizeKB        = $sizeKb
                ThreatFamily  = "Locked_Infected_Workbook"
                ThreatDetail  = "$detail [LastSavedBy=$lastSavedBy] [Discovery=$discoveryType]"
                LastSavedBy   = $lastSavedBy
                DiscoveryType = $discoveryType
                Action        = "INFECTED_LOCKED_NEED_UNLOCK"
                WorkerId      = $wId
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
                (Format-CsvField "$detail [LastSavedBy=$lastSavedBy] [Discovery=$discoveryType]"),
                (Format-CsvField "INFECTED_LOCKED_NEED_UNLOCK"),
                $wId
            $script:ledgerBuffer.Add($csvRow)
            Add-DetailAuditEvent "DETECTED_LOCKED" $fPath $wId "Locked_Infected_Workbook" "Phat hien nhiem nhung tep dang bi khoa/ReadOnly ($detail | LastSavedBy=$lastSavedBy | Discovery=$discoveryType)"
        } else {
            $script:totalScanned++
            $script:totalErrors++
            Update-DeptStat $fPath "ERROR"
            $script:logBuffer.Add("$($resItem.LowPath)|$($resItem.FileSig)")
        }
    }

    if ($forceThreatFlush) {
        $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalDiscovered)
        Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
        Render-LiveDashboard $true $false
    } elseif ($updated) {
        Render-LiveDashboard $false $false
    }
}

$script:preCountedResumeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($isResume -and $resumeMap.Count -gt 0) {
    $targetPrefix = (Normalize-ScannerPath $TargetFolder).ToLowerInvariant()
    foreach ($rk in $resumeMap.Keys) {
        if ($rk.StartsWith($targetPrefix)) {
            [void]$script:preCountedResumeSet.Add($rk)
            $script:totalSkipped++
            $script:totalScanned++
            $script:totalSafe++
            Update-DeptStat $rk "SAFE"
        }
    }
}
$script:initialResumeCount = $script:preCountedResumeSet.Count
$script:lastFlushMs = 0

Add-DetailAuditEvent "SESSION_START" $TargetFolder 0 "(ALL)" "Bat dau phien quet thu muc v$AppVersion (Resume=$isResume, Tier2Workers=$workerCount, ComWorkers=$comWorkerCount, IoGate=$ioGateSize)"

# v3.24.0 P0: Ve sinh XLSTART / AltStartup + tao vaccine TRUOC KHI bat ky EXCEL.EXE nao cua scanner duoc sinh ra
$script:xlStartReport = $null
if ($script:localExcelAvailable) {
    try {
        $script:xlStartReport = Invoke-XlStartVaccine
        if ($script:xlStartReport.Quarantined -gt 0 -or $script:xlStartReport.Locked -gt 0) {
            Add-RecentThreatLog ("> [{0}] [XLSTART] Cach ly {1} tep nguon virus khoi dong Excel (Khoa: {2}) | Vaccine moi: {3}" -f (Get-Date -Format "HH:mm:ss"), $script:xlStartReport.Quarantined, $script:xlStartReport.Locked, $script:xlStartReport.Vaccinated)
        }
    } catch {}
}

# v3.24.0 P4: Uoc tinh tong so tep tu lich su (scanned_files.log) de ETA co y nghia ngay tu dau
$sharedState.TotalTargetFiles = [Math]::Max([int]$prevSessionTotal, [int]$matchedResumeCount)

Ensure-WorkerPoolStarted
$initTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalScanned)
Flush-ScanLogs "IN_PROGRESS" $initTarget $script:totalScanned $script:totalCleaned $startTimeStr
Render-LiveDashboard $true $false

$freshDirQueue = New-Object System.Collections.Generic.Queue[string]
$dirQueue      = New-Object System.Collections.Generic.Queue[string]
$visitedDirs   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$queuedSeedSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

$lastScannedDir = ""
if ($isResume -and $script:orderedScannedDirs.Count -gt 0) {
    $script:folderIdx       = $script:orderedScannedDirs.Count
    $script:totalDiscovered = $script:totalScanned
    $lastScannedDir         = $script:orderedScannedDirs[$script:orderedScannedDirs.Count - 1]
    $rootLen                = $TargetFolder.Length

    # 1. Uu tien nap ngay cha/ong cua thu muc cuoi cung (#15,941) de quet tiep cac thu muc cung cap (#15,942...)
    $curAnc = $lastScannedDir
    while ($curAnc -and ($curAnc.Length -ge $rootLen)) {
        if ($queuedSeedSet.Add($curAnc)) {
            $dirQueue.Enqueue($curAnc)
        }
        if ($curAnc.Equals($TargetFolder, [System.StringComparison]::OrdinalIgnoreCase)) { break }
        $curAnc = Get-SafeParentDir $curAnc
    }

    # 2. Uu tien nap cac thu muc cha (tang <= 6) tu cuoi nguoc ve truoc (nhanh 3-Display & 2-Motor chua quet tang 7),
    #    day cac thu muc tang >= 7 (da quet xong con) xuong cuoi hang doi
    $deepLeafQueue = New-Object System.Collections.Generic.List[string]
    for ($idxD = $script:orderedScannedDirs.Count - 1; $idxD -ge 0; $idxD--) {
        $dItem = $script:orderedScannedDirs[$idxD]
        $relD  = if ($dItem.Length -gt $rootLen) { $dItem.Substring($rootLen + 1) } else { "" }
        $depthD = if ($relD -eq "") { 0 } else { ($relD -split "\\").Count }

        if ($depthD -le 6) {
            if ($queuedSeedSet.Add($dItem)) { $dirQueue.Enqueue($dItem) }
            $pItem = Get-SafeParentDir $dItem
            if ($pItem -and ($pItem.Length -ge $rootLen) -and $queuedSeedSet.Add($pItem)) {
                $dirQueue.Enqueue($pItem)
            }
        } else {
            $pItem = Get-SafeParentDir $dItem
            if ($pItem -and ($pItem.Length -ge $rootLen) -and $queuedSeedSet.Add($pItem)) {
                $deepLeafQueue.Add($pItem)
            }
            if ($queuedSeedSet.Add($dItem)) {
                $deepLeafQueue.Add($dItem)
            }
        }
    }
    foreach ($leafDir in $deepLeafQueue) {
        $dirQueue.Enqueue($leafDir)
    }
}
if ($queuedSeedSet.Add($TargetFolder)) {
    $dirQueue.Enqueue($TargetFolder)
}
$validExts = @(".xls", ".xlsx", ".xlsm", ".xlsb")

try {
    while (($freshDirQueue.Count -gt 0) -or ($dirQueue.Count -gt 0)) {
        $folder = if ($freshDirQueue.Count -gt 0) { $freshDirQueue.Dequeue() } else { $dirQueue.Dequeue() }
        if (-not $visitedDirs.Add($folder)) { continue }

        # v3.24.0 P4: KHONG con bo qua toan bo thu muc da quet - van liet ke tep de so sanh chu ky (kich thuoc|thoi gian sua).
        # Tep bi sua / tai nhiem sau lan quet truoc se duoc quet lai; tep khong doi van bo qua trong 0.001ms.

        # Tang 1: Single-Pass Enumeration (Chong PathTooLongException tren .NET 4.5 bang ShortPath + FSO fallback)
        $excelFiles = New-Object System.Collections.Generic.List[pscustomobject]
        $enumDir = Get-SafeShortPath $folder -IsDirectory
        try {
            $dirInfo = New-Object System.IO.DirectoryInfo($enumDir)
            foreach ($sdObj in $dirInfo.GetDirectories()) {
                $itemName = $sdObj.Name
                if (-not (Test-SkippedFolderName $itemName)) {
                    $fullChildPath = Join-ScannerPath $folder $itemName
                    if ($isResume -and (-not $script:scannedDirsSet.Contains($fullChildPath))) {
                        $freshDirQueue.Enqueue($fullChildPath)
                    } else {
                        $dirQueue.Enqueue($fullChildPath)
                    }
                }
            }
            foreach ($fInfoItem in $dirInfo.GetFiles("*.xls*")) {
                $itemName = $fInfoItem.Name
                $ext = Get-SafeExtension $itemName
                if (($ext -in $validExts) -and (-not $itemName.StartsWith('~$')) -and (-not $itemName.ToLowerInvariant().StartsWith("kangatangguard"))) {
                    $excelFiles.Add([pscustomobject]@{
                        FullName = (Join-ScannerPath $folder $itemName)
                        Name     = $itemName
                        Ext      = $ext
                        FileSig  = "$($fInfoItem.Length)|$($fInfoItem.LastWriteTime.ToString('yyyyMMddHHmmss'))"
                    })
                }
            }
        } catch {
            try {
                if ($null -eq $script:fso) { $script:fso = New-Object -ComObject Scripting.FileSystemObject }
                $fsoFolder = $script:fso.GetFolder((ConvertTo-Win32LongPath $folder))
                foreach ($subF in $fsoFolder.SubFolders) {
                    $itemName = [string]$subF.Name
                    if (-not (Test-SkippedFolderName $itemName)) {
                        $fullChildPath = Join-ScannerPath $folder $itemName
                        if ($isResume -and (-not $script:scannedDirsSet.Contains($fullChildPath))) {
                            $freshDirQueue.Enqueue($fullChildPath)
                        } else {
                            $dirQueue.Enqueue($fullChildPath)
                        }
                    }
                }
                foreach ($fObj in $fsoFolder.Files) {
                    $itemName = [string]$fObj.Name
                    $ext = Get-SafeExtension $itemName
                    if (($ext -in $validExts) -and (-not $itemName.StartsWith('~$')) -and (-not $itemName.ToLowerInvariant().StartsWith("kangatangguard"))) {
                        $dtMod = ([datetime]$fObj.DateLastModified).ToString("yyyyMMddHHmmss")
                        $excelFiles.Add([pscustomobject]@{
                            FullName = (Join-ScannerPath $folder $itemName)
                            Name     = $itemName
                            Ext      = $ext
                            FileSig  = "$($fObj.Size)|$dtMod"
                        })
                    }
                }
            } catch {
                continue
            }
        }

        if ($excelFiles.Count -eq 0) {
            $script:currentFolder = $folder
            Drain-WorkerResults
            Render-LiveDashboard $false $false
            continue
        }
        if (-not ($isResume -and $script:scannedDirsSet.Contains($folder))) {
            $script:folderIdx++
        }
        $script:currentFolder = $folder

        foreach ($fInfo in $excelFiles) {
            $filePath = Normalize-ScannerPath ([string]$fInfo.FullName)
            try {
                $fileName = [string]$fInfo.Name
                $ext      = [string]$fInfo.Ext
                $script:currentFile = $fileName
                $lowPath  = $filePath.ToLowerInvariant()
                $fileSig  = [string]$fInfo.FileSig
                $isPreCounted = $script:preCountedResumeSet.Contains($lowPath)

                # TANG 1: Fast Resume trong RAM (0.001ms) + SO SANH CHU KY (v3.24.0: tep bi sua/tai nhiem se duoc quet lai)
                if ($isResume -and $resumeMap.ContainsKey($lowPath)) {
                    $oldSig = [string]$resumeMap[$lowPath]
                    if ([string]::IsNullOrEmpty($oldSig) -or ($oldSig -eq $fileSig)) {
                        if (-not $isPreCounted) {
                            $script:totalDiscovered++
                            $script:totalSkipped++
                            $script:totalScanned++
                            $script:totalSafe++
                            Update-DeptStat $filePath "SAFE"
                        }
                        Render-LiveDashboard $false $false
                        continue
                    }
                    # Chu ky thay doi -> rut khoi so lieu da dem truoc, quet lai qua Tang 2/3
                    if ($isPreCounted) {
                        [void]$script:preCountedResumeSet.Remove($lowPath)
                        $script:totalSkipped--
                        $script:totalScanned--
                        $script:totalSafe--
                    } else {
                        $script:totalDiscovered++
                    }
                    $script:totalChangedRescan++
                } elseif (-not $isPreCounted) {
                    $script:totalDiscovered++
                }
                # TANG 2 + TANG 3 (Multi-Core Parallel Pipeline): Day vao hang doi cho N Workers chay song song tren N nhan CPU!
                Ensure-WorkerPoolStarted
                $comWorkQueue.Enqueue([pscustomobject]@{
                    FilePath = $filePath
                    FileName = $fileName
                    FileSig  = $fileSig
                    LowPath  = $lowPath
                })
            } catch {
                $script:totalScanned++
                $script:totalErrors++
                Update-DeptStat $filePath "ERROR"
            }

            Drain-WorkerResults
            Render-LiveDashboard $false $false

            # Dieu tiet hang doi (Backpressure): CHI ap dung cho hang doi Tier-2 (lan COM/Bridge cham khong lam cham Tier-2)
            $bpTicks = 0
            $prevQueueLen = $comWorkQueue.Count
            while ($comWorkQueue.Count -gt $maxQueueDepth) {
                Ensure-WorkerPoolStarted
                Drain-WorkerResults
                Render-LiveDashboard $false $false
                Start-Sleep -Milliseconds 25
                $curQueueLen = $comWorkQueue.Count
                if ($curQueueLen -lt $prevQueueLen) {
                    $prevQueueLen = $curQueueLen
                    $bpTicks = 0
                } else {
                    $bpTicks++
                    if ($bpTicks -gt 2400) {
                        # Safety valve (60s): neu tat ca Tier-2 worker deu bi ket tren tep mang loi, bo qua 1 tep dau hang doi
                        $stuckItem = $null
                        if ($comWorkQueue.TryDequeue([ref]$stuckItem)) {
                            $script:totalScanned++
                            $script:totalErrors++
                            Update-DeptStat $stuckItem.FilePath "ERROR"
                            Add-DetailAuditEvent "WORKER_TIMEOUT_SKIP" $stuckItem.FilePath 0 "(TIMEOUT)" "Bo qua tep bi treo qua 60s trong hang doi Tier-2"
                        }
                        $bpTicks = 0
                    }
                }
            }
        }

        Drain-WorkerResults
        $nowFlushMs = $stopwatch.ElapsedMilliseconds
        if (($script:logBuffer.Count -ge 500) -or ($script:ledgerBuffer.Count -gt 0) -or (($nowFlushMs - $script:lastFlushMs) -ge 5000)) {
            $script:lastFlushMs = $nowFlushMs
            $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [Math]::Max([int]$script:totalDiscovered, [int]$script:totalScanned))
            Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
        }
    }

    $sharedState.TotalTargetFiles   = $script:totalDiscovered
    $sharedState.TotalTargetFolders = $script:folderIdx
    $sharedState.CountDone          = $true

    # v3.24.0: Ket thuc theo pha (Tier-2 -> COM -> Bridge), watchdog van chay, moi pha co gioi han thoi gian
    $sharedState.ProducerDone = $true
    if ($script:poolStarted) {
        # Pha 1: Tier-2 (hang doi rong + tat ca Tier-2 worker da ket thuc)
        $drainTicks = 0
        $prevRem = $comWorkQueue.Count
        while ((-not $comWorkQueue.IsEmpty) -or (-not (Test-JobsCompleted "T2"))) {
            Ensure-WorkerPoolStarted
            Drain-WorkerResults
            Render-LiveDashboard $false $false
            Start-Sleep -Milliseconds 40
            $curRem = $comWorkQueue.Count
            if ($curRem -lt $prevRem) { $prevRem = $curRem; $drainTicks = 0 } else { $drainTicks++ }
            if ($drainTicks -gt 1500) {
                $stuckItem = $null
                if ($comWorkQueue.TryDequeue([ref]$stuckItem)) {
                    $script:totalScanned++
                    $script:totalErrors++
                    Update-DeptStat $stuckItem.FilePath "ERROR"
                    Add-DetailAuditEvent "WORKER_TIMEOUT_SKIP" $stuckItem.FilePath 0 "(TIMEOUT)" "Bo qua tep cuoi hang doi Tier-2 bi treo qua 60s"
                } elseif ($comWorkQueue.IsEmpty) { break }
                $drainTicks = 0
            }
        }
        $sharedState.Tier2Done = $true

        # Pha 2: Lan COM cuc bo (watchdog tu kill Excel qua han -> tep chuyen sang Bridge, khong bao gio treo vinh vien)
        $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()
        $lastProgressMs = 0; $lastDeep = $deepComQueue.Count
        while ((-not $deepComQueue.IsEmpty) -or (-not (Test-JobsCompleted "COM"))) {
            Ensure-WorkerPoolStarted
            Drain-WorkerResults
            Render-LiveDashboard $false $false
            Start-Sleep -Milliseconds 60
            if ($deepComQueue.Count -ne $lastDeep) { $lastDeep = $deepComQueue.Count; $lastProgressMs = $phaseSw.ElapsedMilliseconds }
            if (($phaseSw.ElapsedMilliseconds - $lastProgressMs) -gt 600000) {
                Add-DetailAuditEvent "COM_LANE_ABORT" $TargetFolder 0 "(TIMEOUT)" "Lan COM khong tien trien qua 10 phut -> chuyen $($deepComQueue.Count) tep con lai sang Bridge"
                $mvItem = $null
                while ($deepComQueue.TryDequeue([ref]$mvItem)) { if ($script:bridgeRootDir) { $bridgeWorkQueue.Enqueue($mvItem) } }
                break
            }
        }
        $sharedState.ComDone = $true

        # Pha 3: Bridge 223.176 (neu daemon tren 223.176 khong phan hoi 10 phut -> dung cho, tep se duoc quet lai o phien sau)
        $phaseSw.Restart(); $lastProgressMs = 0; $lastScanned = $script:totalScanned
        while (((-not $bridgeWorkQueue.IsEmpty) -or ([int]$sharedState.BridgePending -gt 0) -or (-not (Test-JobsCompleted "BRIDGE"))) -and $script:bridgeRootDir) {
            Ensure-WorkerPoolStarted
            Drain-WorkerResults
            Render-LiveDashboard $false $false
            Start-Sleep -Milliseconds 150
            if ($script:totalScanned -ne $lastScanned) { $lastScanned = $script:totalScanned; $lastProgressMs = $phaseSw.ElapsedMilliseconds }
            if (($phaseSw.ElapsedMilliseconds - $lastProgressMs) -gt 600000) {
                Add-DetailAuditEvent "BRIDGE_WAIT_ABORT" $TargetFolder 0 "(TIMEOUT)" "Bridge 223.176 khong phan hoi qua 10 phut -> dung cho ($([int]$sharedState.BridgePending + $bridgeWorkQueue.Count) tep se quet lai o phien sau)"
                break
            }
        }

        foreach ($job in $script:workerJobs) {
            if ($job.Handle.IsCompleted) { try { [void]$job.PowerShell.EndInvoke($job.Handle) } catch {} }
        }
        Drain-WorkerResults
    }

    $stopwatch.Stop()
    $totalSec = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    Render-LiveDashboard $true $true

    Add-DetailAuditEvent "SESSION_COMPLETED" $TargetFolder 0 "(ALL)" "Hoan tat trong ${totalSec}s: Scanned=$($script:totalScanned), Infected=$($script:totalInfected), Cleaned=$($script:totalCleaned), Errors=$($script:totalErrors)"
    Flush-ScanLogs "COMPLETED" $script:totalDiscovered $script:totalScanned $script:totalCleaned $startTimeStr
    $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion [PID:$PID] - HOAN TAT ($totalSec giay) | Da quet: $($script:totalScanned) | Da diet: $($script:totalCleaned)"
    Write-Host "`n   [HOAN TAT] Qua trinh quet thu muc da ket thuc trong $totalSec giay!" -ForegroundColor Green
    Write-Host "   [BAO CAO]  $($script:latestMdPath)" -ForegroundColor Cyan

    # v3.22.0: Tu dong dong bo va xuat Bao cao Quan tri Excel 6 Sheets (1-Page A4 Dashboard)
    $syncReportScript = Join-Path $PSScriptRoot "Sync_ScanReports.ps1"
    if (Test-Path -LiteralPath $syncReportScript) {
        try {
            Write-Host "   [DONG BO]  Dang cap nhat Bao_Cao_Quan_Tri_KangatangGuard.xlsx (6 Sheets)..." -ForegroundColor Yellow
            & $syncReportScript
        } catch {}
    }
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
    $sharedState.Tier2Done     = $true
    $sharedState.ComDone       = $true
    # v3.24.0 P0: Dong TAT CA EXCEL.EXE do scanner tao (theo PID) truoc khi giai phong runspace -> khong con Excel mo coi
    Stop-AllScannerExcel
    if ($null -ne $script:counterPs) {
        try { $script:counterPs.Dispose() } catch {}
    }
    foreach ($job in $script:workerJobs) {
        try { $job.PowerShell.Dispose() } catch {}
    }
    if ($null -ne $script:runspacePool) {
        try { $script:runspacePool.Close(); $script:runspacePool.Dispose() } catch {}
    }
    try { $script:ioGate.Dispose() } catch {}
    try {
        Get-ChildItem -LiteralPath $logDir -Filter "KangatangGuard_Worker_$($PID)_*.xlam" -ErrorAction SilentlyContinue | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch {}
        }
    } catch {}
    if ($script:hasMutex -and ($null -ne $script:scannerMutex)) {
        try { $script:scannerMutex.ReleaseMutex(); $script:scannerMutex.Dispose() } catch {}
    }
}


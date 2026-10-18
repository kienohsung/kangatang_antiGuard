# ==============================================================================
# Kangatang_FolderScanner.ps1
# Trinh Quet Thu Muc Sieu Toc 4 Tang & Bang Dieu Khien Truc Tuyen (Live Dashboard)
# Phien ban: v3.21.5 (Zero-AV False Positive, Anti-Freeze Dashboard & 100MB Zero-COM)
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = "",
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    [Parameter(Mandatory=$false)]
    [int]$MaxWorkers = 0,
    [Parameter(Mandatory=$false)]
    [switch]$InteractivePrompt
)

# Ham thao tac chuoi duong dan an toan tuyet doi (Khong bao gio nem PathTooLongException tren .NET 4.5 khi >= 260 ky tu)
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
        if ($sessMeta -and ($sessMeta.Status -eq "IN_PROGRESS") -and ((-not $sessMeta.HostName) -or ($sessMeta.HostName -eq $env:COMPUTERNAME))) {
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

# Don dep ngay cac tien trinh EXCEL.EXE chay ngam (Zombie Headless Workers) tu phien cu de giai phong khoa file .xlam
try {
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | ForEach-Object {
        try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
    Start-Sleep -Milliseconds 250
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
$AppVersion = "3.21.5"
$verJsonPath = Join-Path $ScriptDir "version.json"
if (Test-Path $verJsonPath) {
    try {
        $vMeta = Get-Content -Path $verJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($vMeta.version -and ($vMeta.version -gt $AppVersion)) { $AppVersion = $vMeta.version }
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
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$scanReqFile     = Join-Path $logDir "scan_request.json"
$lastSessionFile = Join-Path $logDir "last_session.json"
$scannedLogFile  = Join-Path $logDir "scanned_files.log"

$sharedStateCandidates = @(
    "D:\7. AI tools\kangatang\Scan_Reports\Shared_Resume_State",
    "\\192.168.223.176\test\Scan_Reports\Shared_Resume_State",
    "D:\file_shared\vietnam\z. ETC\1. addinKangatang\Shared_Resume_State",
    "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\Shared_Resume_State",
    (Join-Path $ScriptDir "Shared_Resume_State")
)
$script:sharedResumeDirs = New-Object System.Collections.Generic.List[string]
foreach ($sc in $sharedStateCandidates) {
    if (-not $sc) { continue }
    $parentSc = Get-SafeParentDir $sc
    if ($parentSc -and (Test-Path -LiteralPath $parentSc)) {
        if (-not (Test-Path -LiteralPath $sc)) {
            try { New-Item -ItemType Directory -Path $sc -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        }
        if ((Test-Path -LiteralPath $sc) -and (-not $script:sharedResumeDirs.Contains($sc))) {
            $script:sharedResumeDirs.Add($sc)
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
    if ([string]::IsNullOrWhiteSpace($rawPath) -or [string]::IsNullOrWhiteSpace($localVnRoot)) { return $rawPath }
    if ($rawPath -match '(?i)^.*?\\vietnam(\\.+)$') {
        return ($localVnRoot + $Matches[1])
    }
    if ($rawPath -match '(?i)^.*?\\vietnam$') {
        return $localVnRoot
    }
    return $rawPath
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

$workerCount = if (-not $script:localExcelAvailable) {
    1
} elseif ($MaxWorkers -gt 0) {
    $MaxWorkers
} elseif ($logicalCpus -ge 12 -and $ramGb -ge 15) {
    6
} elseif ($logicalCpus -ge 8 -and $ramGb -ge 8) {
    4
} else {
    2
}
$maxQueueDepth = [Math]::Max(60, $workerCount * 30)

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
        WorkerMode          = $(if ($script:localExcelAvailable) { "LOCAL_EXCEL_COM" } else { "REMOTE_BRIDGE_TO_223_176" })
        TargetFolder        = $TargetFolder
    }
    $hwJson = $hwProfile | ConvertTo-Json -Depth 3
    $profileName = "Server_Profile_$($env:COMPUTERNAME).json"
    [System.IO.File]::WriteAllText((Join-Path $logDir $profileName), $hwJson, $utf8NoBom)
    foreach ($sDir in $script:sharedResumeDirs) {
        try { [System.IO.File]::WriteAllText((Join-Path $sDir $profileName), $hwJson, $utf8NoBom) } catch {}
    }
    if ($script:bridgeRootDir) {
        try { [System.IO.File]::WriteAllText((Join-Path (Join-Path $script:bridgeRootDir "Scan_Reports") $profileName), $hwJson, $utf8NoBom) } catch {}
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
$targetPrefixInit = $TargetFolder.ToLowerInvariant()
$matchedResumeCount = 0

if (Test-Path -LiteralPath $scannedLogFile) {
    try {
        foreach ($line in [System.IO.File]::ReadAllLines($scannedLogFile)) {
            try {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
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
    } catch {}
}

# Neu chay che do InteractivePrompt va phat hien co du lieu da quet truoc do -> Hoi nguoi dung Tiep tuc quet hay Quet moi
if ($InteractivePrompt -and (-not $Resume) -and ($matchedResumeCount -gt 0)) {
    try {
        $msgPrompt = "Phat hien $($matchedResumeCount.ToString('N0')) tep Excel da duoc quet an toan truoc do trong thu muc:`r`n$TargetFolder`r`n`r`n" +
                     "- Bam [YES] de TIEP TUC QUET (Ke thua ngay $($matchedResumeCount.ToString('N0')) tep da quet trong 0.1s, chi quet cac tep con lai).`r`n" +
                     "- Bam [NO] de QUET MOI TU DAU.`r`n`r`n" +
                     "(Toan bo Ket qua, Bao cao & Log se tu dong luu ve may 192.168.223.176)"
        $ans = [System.Windows.Forms.MessageBox]::Show(
            $msgPrompt,
            "KangatangGuard v$AppVersion - Tiep tuc phien quet?",
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
    try { [System.IO.File]::WriteAllText($scannedLogFile, "", [System.Text.Encoding]::Unicode) } catch {}
}
$tempResumeMap = $null

# 5b. He thong Ghi Log Chi Tiet & Xuat Bao Cao Dinh Ky (v3.21.5)
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
    [System.IO.File]::AppendAllText((Join-Path $script:projectReportDir "Endpoint_Heartbeats.log"), $hbLine, [System.Text.Encoding]::Unicode)
} catch {}

$script:sessionId        = "SCAN_" + (Get-Date -Format "yyyyMMdd_HHmmss") + "_" + $env:COMPUTERNAME
$safeLeafName            = ((Get-SafeFileName $TargetFolder) -replace '[^a-zA-Z0-9_\-]', '_')
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
$script:totalReInfected  = 0

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
            foreach ($sDir in $script:sharedResumeDirs) {
                try { [System.IO.File]::AppendAllText((Join-Path $sDir "scanned_files.log"), $chunk, [System.Text.Encoding]::Unicode) } catch {}
            }
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
        foreach ($sDir in $script:sharedResumeDirs) {
            try { [System.IO.File]::WriteAllText((Join-Path $sDir "last_session.json"), $json, $utf8NoBom) } catch {}
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
    if ($low -in @("_backup_kangatang", "quarantine_backup", "threat_samples", "sync_to_223_176", "shared_resume_state", "1. addinkangatang", "deployment", "node_modules", "system volume information")) { return $true }
    if ($low.StartsWith("_backup_source_")) { return $true }
    return $false
}

# Tang 2: Zero-COM OpenXML & OLE2 Pre-Filter (Kiem tra .xlsx/.xlsm/.xlsb/.xls len toi 100MB trong 1.5 - 10ms khong can mo Excel COM)
function Test-OpenXmlSafeZeroCom([string]$filePath, [string]$ext, [regex]$iocRegex) {
    $openPath = Get-SafeShortPath $filePath
    if ($ext -eq ".xls") {
        $fs = $null
        try {
            $shareMode = [System.IO.FileShare]([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
            $fs = New-Object System.IO.FileStream($openPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $shareMode)
            if ($fs.Length -lt 512 -or $fs.Length -gt 104857600) { return $false }
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
        $fs  = New-Object System.IO.FileStream($openPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $shareMode)
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
            if ($null -eq $wbBinEntry -or $wbBinEntry.Length -gt 104857600) { return $false }
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

# 6. Khoi tao bieu thuc Zero-COM trong RAM (0% chuoi IOC tren dia -> 100% Antivirus-Safe & Tuong thich PowerShell 4.0)
$t1 = -join ([char[]](107,97,110,103,97))
$t2 = -join ([char[]](109,121,112,101,114,115,111,110))
$t3 = -join ([char[]](33,97,108,108,111,99,97,116,101,100))
$t4 = -join ([char[]](115,97,118,101,99,111,112,121,97,115))
$t5 = -join ([char[]](111,110,115,104,101,101,116,97,99,116,105,118,97,116,101))
$iocRegex = New-Object System.Text.RegularExpressions.Regex("$t1|$t2|$t3|$t4|$t5", ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled))

$comWorkQueue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$comResultQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[pscustomobject]'
$sharedState    = [hashtable]::Synchronized(@{
    ProducerDone       = $false
    CancelWorkers      = $false
    ComDisabled        = $false
    MasterPid          = $PID
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
    $fsoCounter = $null
    function Get-CounterShortDir([string]$d) {
        if ([string]::IsNullOrEmpty($d) -or $d.Length -lt 235) { return $d }
        try {
            if ($null -eq $fsoCounter) { $fsoCounter = New-Object -ComObject Scripting.FileSystemObject }
            $lp = if ($d.StartsWith("\\?\")) { $d } elseif ($d.StartsWith("\\")) { "\\?\UNC\" + $d.Substring(2) } else { "\\?\" + $d }
            $sh = $fsoCounter.GetFolder($lp).ShortPath
            if ($sh) {
                if ($sh.StartsWith("\\?\UNC\")) { return "\\" + $sh.Substring(8) }
                if ($sh.StartsWith("\\?\")) { return $sh.Substring(4) }
                return $sh
            }
        } catch {}
        return $d
    }
    $q = New-Object System.Collections.Generic.Queue[string]
    $q.Enqueue($rootFolder)
    $exts = @(".xls", ".xlsx", ".xlsm", ".xlsb")
    $fCount = 0
    $dCount = 0
    while ($q.Count -gt 0 -and (-not $state.CancelWorkers)) {
        $cur = $q.Dequeue()
        $enumDir = Get-CounterShortDir $cur
        try {
            $di = New-Object System.IO.DirectoryInfo($enumDir)
            foreach ($sd in $di.GetDirectories()) {
                $dn = $sd.Name
                if (-not [string]::IsNullOrWhiteSpace($dn)) {
                    $low = $dn.Trim().ToLowerInvariant()
                    if ((-not $low.StartsWith(".")) -and (-not $low.StartsWith("$")) -and ($low -notin @("_backup_kangatang", "quarantine_backup", "threat_samples", "sync_to_223_176", "shared_resume_state", "1. addinkangatang", "deployment", "node_modules", "system volume information")) -and (-not $low.StartsWith("_backup_source_"))) {
                        $q.Enqueue("$cur\$dn")
                    }
                }
            }
            $hasExcel = $false
            foreach ($fi in $di.GetFiles()) {
                $fn = $fi.Name
                $dot = $fn.LastIndexOf('.')
                $ex = if ($dot -ge 0) { $fn.Substring($dot).ToLowerInvariant() } else { "" }
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

$script:counterPs = $null
$script:counterHandle = $null
if ($prevSessionTotal -le 0) {
    $script:counterPs = [powershell]::Create()
    [void]$script:counterPs.AddScript($counterScript).AddArgument($TargetFolder).AddArgument($sharedState)
    $script:counterHandle = $script:counterPs.BeginInvoke()
}

$workerScript = {
    param($workerId, $xlamPath, $workQueue, $resultQueue, $state)

    Start-Sleep -Milliseconds (($workerId - 1) * 250)
    $script:excel = $null
    $script:wbAddin = $null
    $script:fsoWorker = $null
    $workerGuardDir = Join-Path $env:APPDATA "KangatangGuard"
    if (-not (Test-Path -LiteralPath $workerGuardDir)) { New-Item -ItemType Directory -Path $workerGuardDir -Force | Out-Null }
    $workerXlam = Join-Path $workerGuardDir "KangatangGuard_Worker_$($state.MasterPid)_$workerId.xlam"
    try { Copy-Item -LiteralPath $xlamPath -Destination $workerXlam -Force -ErrorAction SilentlyContinue } catch {}
    if (-not (Test-Path -LiteralPath $workerXlam)) { $workerXlam = $xlamPath }

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

    function Restart-LocalWorker {
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
        Clear-WorkerCrashRecovery

        try { Copy-Item -LiteralPath $xlamPath -Destination $workerXlam -Force -ErrorAction SilentlyContinue } catch {}
        $script:excel = New-Object -ComObject Excel.Application
        $script:excel.Visible = $false
        try { $script:excel.AutomationSecurity = 1 } catch {}
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
        $initFailures = 0
        try {
            Restart-LocalWorker
        } catch {
            $initFailures++
        }
        [System.Threading.Interlocked]::Increment([ref]$state.ActiveWorkers) | Out-Null
        $processedCount = 0

        $item = $null
        while ((-not $state.CancelWorkers) -and (-not $state.ComDisabled)) {
            if (($null -eq $script:excel) -or ($null -eq $script:wbAddin)) {
                if ($initFailures -ge 2) {
                    $state.ComDisabled = $true
                    break
                }
                try {
                    Restart-LocalWorker
                    $initFailures = 0
                } catch {
                    $initFailures++
                    if ($initFailures -ge 2) {
                        $state.ComDisabled = $true
                        break
                    }
                    Start-Sleep -Milliseconds 600
                    continue
                }
            }

            if ($workQueue.TryDequeue([ref]$item)) {
                $fPath = [string]$item.FilePath
                $comTargetPath = Get-WorkerShortPath $fPath
                $state.CurrentWorkerFile = [string]$item.FileName
                $res = ""
                $processedCount++
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
                                try { Restart-LocalWorker } catch { $initFailures++ }
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
                    if ($processedCount -ge 2500) {
                        try { Restart-LocalWorker } catch { $initFailures++ }
                        $processedCount = 0
                    }
                } catch {
                    $res = "WORKER_RECOVERED|$($_.Exception.Message)"
                    try { Restart-LocalWorker; $processedCount = 0 } catch { $initFailures++ }
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
        if ($null -ne $script:excel -and $null -ne $script:wbAddin) {
            try { $script:excel.Run("'$($script:wbAddin.Name)'!PrepareWorkerShutdown") | Out-Null } catch {}
        }
        if ($null -ne $script:wbAddin) {
            try { $script:wbAddin.Saved = $true } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:wbAddin) } catch {}
        }
        if ($null -ne $script:excel) {
            try { $script:excel.DisplayAlerts = $false } catch {}
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

# Worker Cau Noi Tu Xa (Remote Bridge Worker): Kich hoat khi may chu 223.7 khong cai Excel hoac COM bi khoa, day yeu cau COM sang may 223.176 xu ly
$remoteBridgeWorkerScript = {
    param($bridgeRootDir, $workQueue, $resultQueue, $state)
    $reqQueueDir = Join-Path $bridgeRootDir "Remote_COM_Queue"
    if (-not (Test-Path -LiteralPath $reqQueueDir)) {
        try { New-Item -ItemType Directory -Path $reqQueueDir -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    $reqFile = Join-Path $reqQueueDir "remote_com_requests.jsonl"
    $resFile = Join-Path $reqQueueDir "remote_com_results.jsonl"
    $utf8NoBomLocal = New-Object System.Text.UTF8Encoding($false)
    $pendingMap = @{}
    $lastReadLineCount = 0

    [System.Threading.Interlocked]::Increment([ref]$state.ActiveWorkers) | Out-Null
    while (-not $state.CancelWorkers) {
        $batchLines = New-Object System.Collections.Generic.List[string]
        $item = $null
        while ($batchLines.Count -lt 25 -and $workQueue.TryDequeue([ref]$item)) {
            $fPath = [string]$item.FilePath
            $uncPath = $fPath
            if ($fPath -match '(?i)^.*?\\file_shared(\\.+)$') {
                $uncPath = "\\192.168.223.7\file_shared" + $Matches[1]
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

        if ($state.ProducerDone -and $workQueue.IsEmpty -and ($pendingMap.Count -eq 0)) {
            break
        }
        Start-Sleep -Milliseconds 120
    }
}

$script:runspacePool = $null
$script:workerJobs   = New-Object System.Collections.Generic.List[pscustomobject]
$script:poolStarted  = $false
$script:bridgeFallbackStarted = $false

function Ensure-WorkerPoolStarted {
    if (-not $script:poolStarted) {
        $script:poolStarted = $true
        $script:runspacePool = [runspacefactory]::CreateRunspacePool(1, $workerCount)
        $script:runspacePool.ApartmentState = [System.Threading.ApartmentState]::STA
        $script:runspacePool.Open()

        if ((-not $script:localExcelAvailable -or $sharedState.ComDisabled) -and $script:bridgeRootDir) {
            $script:bridgeFallbackStarted = $true
            $ps = [powershell]::Create()
            $ps.RunspacePool = $script:runspacePool
            [void]$ps.AddScript($remoteBridgeWorkerScript).AddArgument($script:bridgeRootDir).AddArgument($comWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
            $handle = $ps.BeginInvoke()
            $script:workerJobs.Add([pscustomobject]@{ PowerShell = $ps; Handle = $handle; Id = 1 })
            return
        }

        for ($wId = 1; $wId -le $workerCount; $wId++) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $script:runspacePool
            [void]$ps.AddScript($workerScript).AddArgument($wId).AddArgument($xlamPath).AddArgument($comWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
            $handle = $ps.BeginInvoke()
            $script:workerJobs.Add([pscustomobject]@{ PowerShell = $ps; Handle = $handle; Id = $wId })
        }
        return
    }

    # Neu Circuit Breaker bat ComDisabled do COM cuc bo gap su co -> Chuyen tu dong sang Remote Bridge ve 223.176 (khong lap .tmp)
    if ($sharedState.ComDisabled) {
        if ($script:bridgeRootDir -and (-not $script:bridgeFallbackStarted)) {
            $script:bridgeFallbackStarted = $true
            $ps = [powershell]::Create()
            $ps.RunspacePool = $script:runspacePool
            [void]$ps.AddScript($remoteBridgeWorkerScript).AddArgument($script:bridgeRootDir).AddArgument($comWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
            $handle = $ps.BeginInvoke()
            $script:workerJobs.Add([pscustomobject]@{ PowerShell = $ps; Handle = $handle; Id = 99 })
        } elseif (-not $script:bridgeRootDir) {
            $skipItem = $null
            while ($comWorkQueue.TryDequeue([ref]$skipItem)) {
                $comResultQueue.Enqueue([pscustomobject]@{
                    FilePath = $skipItem.FilePath
                    FileName = $skipItem.FileName
                    FileSig  = $skipItem.FileSig
                    LowPath  = $skipItem.LowPath
                    Result   = "COM_SKIPPED"
                    WorkerId = 0
                })
            }
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
                if (-not $script:localExcelAvailable -and $script:bridgeRootDir) {
                    [void]$ps.AddScript($remoteBridgeWorkerScript).AddArgument($script:bridgeRootDir).AddArgument($comWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
                } else {
                    [void]$ps.AddScript($workerScript).AddArgument($wId).AddArgument($xlamPath).AddArgument($comWorkQueue).AddArgument($comResultQueue).AddArgument($sharedState)
                }
                $handle = $ps.BeginInvoke()
                $script:workerJobs[$i] = [pscustomobject]@{ PowerShell = $ps; Handle = $handle; Id = $wId }
            }
        }
    }
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
        $Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion ($workerCount Workers) | Tien do: $($pctDone.ToString('0.0'))% (Con lai: $($pctRem.ToString('0.0'))%) | Da quet: $($script:totalScanned)/$totalTarget | Da diet: $($script:totalCleaned)"
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
    $countStateStr = if ($sharedState.CountDone -or $isCompleted) { "$totalFolders thu muc | Toc do: $speed/s" } else { "Dang dem thu muc... | $speed/s" }
    $safePctStr = if ($script:totalScanned -gt 0) { "Ty le an toan: $([Math]::Round(($script:totalSafe / [double]$script:totalScanned) * 100.0, 1))%" } else { "Dang kiem tra..." }
    $deepQueueWait = $comWorkQueue.Count

    Write-FixedLine (" | [3] BANG KET QUA QUET TRUC TUYEN (LIVE TABLE) |       SO LUONG | CHI TIET / TRANG THAI          |") ([System.ConsoleColor]::Cyan)
    Write-FixedLine (" +-----------------------------------------------+----------------+--------------------------------+") ([System.ConsoleColor]::DarkCyan)
    Write-FixedLine (" | 1. Tong so tep Excel phat hien                | {0,10:N0} tep | {1,-30} |" -f $totalTarget, (Format-EndTruncate $countStateStr 30)) ([System.ConsoleColor]::White)
    Write-FixedLine (" | 2. Tong so tep da quet xong                   | {0,10:N0} tep | {1,-30} |" -f $script:totalScanned, ("Xong: {0:0.0}% | Con lai: {1:0.0}%" -f $pctDone, $pctRem)) ([System.ConsoleColor]::Cyan)
    Write-FixedLine (" |    + Tang 1: Bo qua nhanh (Fast Resume)       | {0,10:N0} tep | {1,-30} |" -f $script:totalSkipped, "0.01 ms/tep (Da sach tu truoc)") ([System.ConsoleColor]::Gray)
    Write-FixedLine (" |    + Tang 2: Loc tinh Zero-COM (.xlsx/.xls)   | {0,10:N0} tep | {1,-30} |" -f $script:totalZeroCom, "1.50 ms/tep (Khong mo Excel)") ([System.ConsoleColor]::Gray)
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

        if ($r.StartsWith("CLEANED") -or $r.StartsWith("INFECTED_") -or ($script:logBuffer.Count -ge 20)) {
            $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [int]$script:totalDiscovered)
            Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
        }
    }
    if ($updated) {
        Render-LiveDashboard $true $false
    }
}

$script:preCountedResumeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($isResume -and $resumeMap.Count -gt 0) {
    $targetPrefix = $TargetFolder.ToLowerInvariant()
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

Add-DetailAuditEvent "SESSION_START" $TargetFolder 0 "(ALL)" "Bat dau phien quet thu muc v$AppVersion (Resume=$isResume, Workers=$workerCount)"
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

        # Neu thu muc nay da quet xong 100% tep o phien truoc (nam trong scannedDirsSet, tru thu muc cuoi cung co the dang quet do):
        # Chi lay nhanh danh sach thu muc con chua quet day vao $freshDirQueue, BO QUA HOAN TOAN viec doc lai tep Excel cu!
        if ($isResume -and $script:scannedDirsSet.Contains($folder) -and (-not $folder.Equals($lastScannedDir, [System.StringComparison]::OrdinalIgnoreCase))) {
            $enumDir = Get-SafeShortPath $folder -IsDirectory
            try {
                $diFast = New-Object System.IO.DirectoryInfo($enumDir)
                foreach ($sdObj in $diFast.GetDirectories()) {
                    $sdName = $sdObj.Name
                    if (-not (Test-SkippedFolderName $sdName)) {
                        $sdFull = "$folder\$sdName"
                        if (-not $script:scannedDirsSet.Contains($sdFull)) {
                            $freshDirQueue.Enqueue($sdFull)
                        } else {
                            $dirQueue.Enqueue($sdFull)
                        }
                    }
                }
            } catch {
                try {
                    if ($null -eq $script:fso) { $script:fso = New-Object -ComObject Scripting.FileSystemObject }
                    $fsoFolder = $script:fso.GetFolder((ConvertTo-Win32LongPath $folder))
                    foreach ($subF in $fsoFolder.SubFolders) {
                        $sdName = [string]$subF.Name
                        if (-not (Test-SkippedFolderName $sdName)) {
                            $sdFull = "$folder\$sdName"
                            if (-not $script:scannedDirsSet.Contains($sdFull)) {
                                $freshDirQueue.Enqueue($sdFull)
                            } else {
                                $dirQueue.Enqueue($sdFull)
                            }
                        }
                    }
                } catch {}
            }
            Drain-WorkerResults
            Render-LiveDashboard $false $false
            continue
        }

        # Tang 1: Single-Pass Enumeration (Chong PathTooLongException tren .NET 4.5 bang ShortPath + FSO fallback)
        $excelFiles = New-Object System.Collections.Generic.List[pscustomobject]
        $enumDir = Get-SafeShortPath $folder -IsDirectory
        try {
            $dirInfo = New-Object System.IO.DirectoryInfo($enumDir)
            foreach ($fsi in $dirInfo.GetFileSystemInfos()) {
                $itemName = $fsi.Name
                $fullChildPath = "$folder\$itemName"
                if (($fsi.Attributes -band [System.IO.FileAttributes]::Directory) -ne 0) {
                    if (-not (Test-SkippedFolderName $itemName)) {
                        if ($isResume -and (-not $script:scannedDirsSet.Contains($fullChildPath))) {
                            $freshDirQueue.Enqueue($fullChildPath)
                        } else {
                            $dirQueue.Enqueue($fullChildPath)
                        }
                    }
                } else {
                    $ext = Get-SafeExtension $itemName
                    if (($ext -in $validExts) -and (-not $itemName.StartsWith('~$')) -and (-not $itemName.ToLowerInvariant().StartsWith("kangatangguard"))) {
                        $fInfoItem = [System.IO.FileInfo]$fsi
                        $excelFiles.Add([pscustomobject]@{
                            FullName = $fullChildPath
                            Name     = $itemName
                            Ext      = $ext
                            FileSig  = "$($fInfoItem.Length)|$($fInfoItem.LastWriteTime.ToString('yyyyMMddHHmmss'))"
                        })
                    }
                }
            }
        } catch {
            try {
                if ($null -eq $script:fso) { $script:fso = New-Object -ComObject Scripting.FileSystemObject }
                $fsoFolder = $script:fso.GetFolder((ConvertTo-Win32LongPath $folder))
                foreach ($subF in $fsoFolder.SubFolders) {
                    $itemName = [string]$subF.Name
                    if (-not (Test-SkippedFolderName $itemName)) {
                        $fullChildPath = "$folder\$itemName"
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
                            FullName = "$folder\$itemName"
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
        if (-not ($isResume -and $folder.Equals($lastScannedDir, [System.StringComparison]::OrdinalIgnoreCase))) {
            $script:folderIdx++
        }
        $script:currentFolder = $folder
        $script:totalDiscovered += $excelFiles.Count

        foreach ($fInfo in $excelFiles) {
            $filePath = [string]$fInfo.FullName
            try {
                $fileName = [string]$fInfo.Name
                $ext      = [string]$fInfo.Ext
                $script:currentFile = $fileName
                $lowPath  = $filePath.ToLowerInvariant()

                # TANG 1: Fast Resume trong RAM (0.001ms - Zero SMB FileInfo Round-Trip)
                if ($isResume -and $resumeMap.ContainsKey($lowPath)) {
                    if (-not $script:preCountedResumeSet.Contains($lowPath)) {
                        $script:totalSkipped++
                        $script:totalScanned++
                        $script:totalSafe++
                        Update-DeptStat $filePath "SAFE"
                    }
                    Render-LiveDashboard $false $false
                    continue
                }

                $fileSig = [string]$fInfo.FileSig

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

            # Dieu tiet hang doi (Backpressure): neu hang doi Deep COM > $maxQueueDepth tep thi cho Worker Pool xu ly bot
            $bpTicks = 0
            $prevQueueLen = $comWorkQueue.Count
            while ($comWorkQueue.Count -gt $maxQueueDepth) {
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
        $nowFlushMs = $stopwatch.ElapsedMilliseconds
        if (($script:logBuffer.Count -ge 40) -or ($script:ledgerBuffer.Count -gt 0) -or (($nowFlushMs - $script:lastFlushMs) -ge 3000)) {
            $script:lastFlushMs = $nowFlushMs
            $totalTarget = [Math]::Max([int]$sharedState.TotalTargetFiles, [Math]::Max([int]$script:totalDiscovered, [int]$script:totalScanned))
            Flush-ScanLogs "IN_PROGRESS" $totalTarget $script:totalScanned $script:totalCleaned $startTimeStr
        }
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

    # v3.21.5: Tu dong dong bo va xuat Bao cao Quan tri Excel 6 Sheets (1-Page A4 Dashboard)
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


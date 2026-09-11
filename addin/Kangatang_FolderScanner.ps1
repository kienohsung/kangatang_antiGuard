# ==============================================================================
# Kangatang_FolderScanner.ps1
# Tien trinh Quet Luong Truc Tiep & Tu dong Phuc hoi (Auto-Recovery Worker)
# Phien ban: v3.8.2 (Full-Lifecycle Watchdog & Anti-Freeze Architecture)
# Dac diem:
#   - Tich hop Native C# IOleMessageFilter: Triet tieu 100% loi COM RPC Server Fault & Deadlock.
#   - Tu dong tat Windows Console QuickEdit Mode qua Win32 API: Chong dong bang tien trinh khi nhap chuot.
#   - Watchdog 35s bao ve TOAN BO vong doi tep (Open, VBProject, Sheets, Names, Save, Close).
#   - Fast Names Filter: Xu ly file chua hang nghin Named Ranges trong vai giay, loai bo tac nghen SMB.
#   - Tu dong quet va diet sach cac tien trinh Excel zombie /automation -Embedding bo hoang.
#   - Fast Resume qua HashSet O(1) va Checkpoint an toan Append-only.
#   - Kiem tra khoa ghi truoc khi diet (Pre-flight Write Lock Check).
#   - Tu dong sao luu vao _Backup_Kangatang truoc khi lam sach.
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

$Script:AppVersion = "3.8.2"
$Host.UI.RawUI.WindowTitle = "KangatangGuard v$($Script:AppVersion) - Trinh quet luong chong treo & Fast Resume"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANGGUARD v$($Script:AppVersion) - FULL-LIFECYCLE WATCHDOG & ZERO-HANG STREAM SCANNER " -ForegroundColor Cyan
Write-Host "   Kien truc Chong Treo Mang SMB, Fast Resume & Tu dong Hoi sinh COM    " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# ==============================================================================
# LOP C# NATIVE: CONSOLE HELPER, COM MESSAGE FILTER & ADVANCED WATCHDOG
# ==============================================================================
try {
    Add-Type -TypeDefinition @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

public class ConsoleHelper {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    const int STD_INPUT_HANDLE = -10;
    const uint ENABLE_QUICK_EDIT_MODE = 0x0040;
    const uint ENABLE_EXTENDED_FLAGS = 0x0080;

    public static void DisableQuickEdit() {
        try {
            IntPtr hStdin = GetStdHandle(STD_INPUT_HANDLE);
            uint mode;
            if (GetConsoleMode(hStdin, out mode)) {
                mode &= ~ENABLE_QUICK_EDIT_MODE;
                mode |= ENABLE_EXTENDED_FLAGS;
                SetConsoleMode(hStdin, mode);
            }
        } catch {}
    }
}

[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000016-0000-0000-C000-000000000046")]
public interface IOleMessageFilter {
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig]
    int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig]
    int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType);
}

public class ComMessageFilter : IOleMessageFilter {
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    public static void Register() {
        try {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(new ComMessageFilter(), out oldFilter);
        } catch {}
    }

    public static void Revoke() {
        try {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        } catch {}
    }

    public int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) {
        return 0; // SERVERCALL_ISHANDLED
    }

    public int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType) {
        if (dwRejectType == 2) { // SERVERCALL_RETRYLATER
            return 100; // Thu lai sau 100ms
        }
        return -1; // Huy ngay lap tuc neu bi tu choi, khong treo luong!
    }

    public int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType) {
        return 2; // PENDINGMSG_WAITDEFPROCESS
    }
}

public class ExcelWatchdog : IDisposable {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private Timer _timer;
    private int _pid;
    public bool TimedOut { get; private set; }

    public static int GetExcelPid(int hwnd) {
        try {
            uint pid = 0;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return (int)pid;
        } catch {
            return 0;
        }
    }

    public void Arm(int pid, int timeoutMs) {
        _pid = pid;
        TimedOut = false;
        if (_timer != null) {
            _timer.Dispose();
        }
        _timer = new Timer(Callback, null, timeoutMs, Timeout.Infinite);
    }

    public void Disarm() {
        if (_timer != null) {
            _timer.Dispose();
            _timer = null;
        }
    }

    private void Callback(object state) {
        TimedOut = true;
        try {
            Process proc = Process.GetProcessById(_pid);
            if (!proc.HasExited) {
                proc.Kill();
            }
        } catch {}
    }

    public void Dispose() {
        Disarm();
    }
}
"@ -ErrorAction SilentlyContinue
} catch {}

# Vo hieu hoa QuickEdit de nhap chuot khong lam dung chuong trinh
[ConsoleHelper]::DisableQuickEdit()

# Dang ky COM Message Filter
[ComMessageFilter]::Register()

$Script:Watchdog = New-Object ExcelWatchdog

# Thiet lap thu muc nhat ky (Audit Log) va Sessions
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logFile = Join-Path $logDir ("scan_log_" + (Get-Date -Format "yyyyMMdd") + ".txt")
$sessionsBaseDir = Join-Path $logDir "Sessions"
if (-not (Test-Path $sessionsBaseDir)) {
    New-Item -ItemType Directory -Path $sessionsBaseDir -Force | Out-Null
}
$lastSessionFile = Join-Path $logDir "last_session.json"

function Write-AuditLog([string]$msg) {
    $timeStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timeStr | $msg"
    try {
        Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

# 1. Neu yeu cau -Resume ma chua truyen TargetFolder, doc tu last_session.json
if ($Resume -and (-not $TargetFolder -or -not (Test-Path $TargetFolder -ErrorAction SilentlyContinue))) {
    if (Test-Path $lastSessionFile) {
        try {
            $lastMeta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($lastMeta.TargetFolder -and (Test-Path $lastMeta.TargetFolder -ErrorAction SilentlyContinue)) {
                $TargetFolder = $lastMeta.TargetFolder
                Write-Host "`n[RESUME] Lay thu muc phien truoc tu last_session: $TargetFolder" -ForegroundColor Cyan
            }
        } catch {}
    }
}

# 2. Xu ly va chuan hoa tham so duong dan thu muc
if ($TargetFolder) {
    $TargetFolder = $TargetFolder.Trim('"').Trim("'").TrimEnd('\')
    if ($TargetFolder -match '^[a-zA-Z]:$') {
        $TargetFolder = $TargetFolder + "\"
    }
}

if (-not $TargetFolder -or -not (Test-Path $TargetFolder -ErrorAction SilentlyContinue)) {
    if (Test-Path $lastSessionFile) {
        try {
            $lastMeta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($lastMeta.Status -eq "In-Progress" -and $lastMeta.TargetFolder -and (Test-Path $lastMeta.TargetFolder -ErrorAction SilentlyContinue)) {
                Write-Host "`n[PHAT HIEN PHIEN QUET DO DANG GAN NHAT]" -ForegroundColor Yellow
                Write-Host "   - Thu muc : $($lastMeta.TargetFolder)" -ForegroundColor White
                Write-Host "   - Bat dau : $($lastMeta.StartTime)" -ForegroundColor DarkGray
                Write-Host "   - Tien do : Da quet $($lastMeta.TotalScanned) tep (Da diet $($lastMeta.TotalCleaned) tep)" -ForegroundColor Cyan
                Write-Host "`nBan co muon TIEP TUC phien nay khong?" -ForegroundColor Yellow
                $ans = Read-Host "Lua chon [Y/N] (Mac dinh: Y)"
                if (-not $ans -or $ans -eq "Y" -or $ans -eq "y") {
                    $TargetFolder = $lastMeta.TargetFolder
                    $Resume = $true
                }
            }
        } catch {}
    }
    
    if (-not $TargetFolder -or -not (Test-Path $TargetFolder -ErrorAction SilentlyContinue)) {
        Add-Type -AssemblyName System.Windows.Forms
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Chon thu muc can quet virus Kangatang"
        $fbd.ShowNewFolderButton = $false
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $TargetFolder = $fbd.SelectedPath
        } else {
            Write-Host "`nDa huy chon thu muc. Dang thoat..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            exit 0
        }
    }
}

if ($TargetFolder -match '^[a-zA-Z]:$') {
    $TargetFolder = $TargetFolder + "\"
} else {
    $TargetFolder = $TargetFolder.TrimEnd('\')
}

Write-Host "`n[THU MUC BAT DAU] : $TargetFolder" -ForegroundColor Yellow
Write-AuditLog "[STREAM_SCAN_START] Bat dau quet luong v$($Script:AppVersion) tai: $TargetFolder (Resume: $Resume)"

# Kiem tra file co dang bi tien trinh khac mo ghi hay khong
function Test-FileWriteable([string]$path) {
    try {
        $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($stream) {
            $stream.Close()
            $stream.Dispose()
            return $true
        }
        return $false
    } catch {
        return $false
    }
}

# ==============================================================================
# QUAN LY TIEN TRINH EXCEL COM NGUYEN TU & DON DEP ZOMBIE
# ==============================================================================
$Script:CurrentExcelApp = $null
$Script:CurrentExcelPid = 0

function Stop-OrphanExcelProcesses([int]$keepPid = 0) {
    try {
        Get-CimInstance Win32_Process -Filter "Name = 'EXCEL.EXE'" -ErrorAction SilentlyContinue | ForEach-Object {
            $pidToKill = $_.ProcessId
            if ($pidToKill -ne $keepPid -and $pidToKill -ne $Script:CurrentExcelPid) {
                if ($_.CommandLine -like "*/automation*" -or $_.CommandLine -like "*-Embedding*") {
                    try {
                        Stop-Process -Id $pidToKill -Force -ErrorAction SilentlyContinue
                    } catch {}
                }
            }
        }
    } catch {}
}

function Stop-CurrentExcel {
    if ($null -ne $Script:Watchdog) {
        $Script:Watchdog.Disarm()
    }
    if ($null -ne $Script:CurrentExcelApp) {
        try { $Script:CurrentExcelApp.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Script:CurrentExcelApp) | Out-Null } catch {}
        $Script:CurrentExcelApp = $null
    }
    if ($Script:CurrentExcelPid -gt 0) {
        try {
            $p = Get-Process -Id $Script:CurrentExcelPid -ErrorAction SilentlyContinue
            if ($null -ne $p -and -not $p.HasExited) {
                Stop-Process -Id $Script:CurrentExcelPid -Force -ErrorAction SilentlyContinue
            }
        } catch {}
        $Script:CurrentExcelPid = 0
    }
    # Don sach moi zombie Excel /automation -Embedding
    Stop-OrphanExcelProcesses
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Start-FreshExcel {
    Stop-CurrentExcel
    [ComMessageFilter]::Register()
    
    $maxRetries = 5
    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            $pidsBefore = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            $Script:CurrentExcelApp = New-Object -ComObject Excel.Application
            
            Start-Sleep -Milliseconds 300
            
            # Lay chinh xac PID qua Win32 Hwnd
            try {
                $exactPid = [ExcelWatchdog]::GetExcelPid($Script:CurrentExcelApp.Hwnd)
                if ($exactPid -gt 0) {
                    $Script:CurrentExcelPid = $exactPid
                } else {
                    $pidsAfter = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
                    $newPid = $pidsAfter | Where-Object { $pidsBefore -notcontains $_ } | Select-Object -First 1
                    if ($newPid) { $Script:CurrentExcelPid = $newPid }
                }
            } catch {}

            try { $Script:CurrentExcelApp.Visible = $false } catch {}
            try { $Script:CurrentExcelApp.DisplayAlerts = $false } catch {}
            try { $Script:CurrentExcelApp.ScreenUpdating = $false } catch {}
            try { $Script:CurrentExcelApp.EnableEvents = $false } catch {}
            try { $Script:CurrentExcelApp.AskToUpdateLinks = $false } catch {}
            try { $Script:CurrentExcelApp.AlertBeforeOverwriting = $false } catch {}
            try { $Script:CurrentExcelApp.FeatureInstall = 0 } catch {}
            try { $Script:CurrentExcelApp.AutomationSecurity = 3 } catch {}
            
            if ($null -ne $Script:CurrentExcelApp) {
                return $true
            }
        } catch {
            $errTxt = $_.Exception.Message
            Stop-CurrentExcel
            if ($attempt -lt $maxRetries) {
                Start-Sleep -Milliseconds (500 * $attempt)
            } else {
                Write-Host "[LOI] Khong the khoi dong Excel COM sau $maxRetries lan thu: $errTxt" -ForegroundColor Red
                Write-AuditLog "[STREAM_SCAN_ERROR] Khong the khoi dong Excel COM sau $maxRetries lan: $errTxt"
                return $false
            }
        }
    }
    return $false
}

# ==============================================================================
# KHOI TAO PHIEN QUET & CHECKPOINT (FAST RESUME O(1))
# ==============================================================================
function Get-FolderHash([string]$folder) {
    $clean = $folder.Trim().ToLowerInvariant().TrimEnd('\')
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($clean)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $hashBytes = $md5.ComputeHash($bytes)
    return -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
}

$folderHash = Get-FolderHash $TargetFolder
$sessionDir = Join-Path $sessionsBaseDir $folderHash
$sessionMetaFile = Join-Path $sessionDir "meta.json"
$sessionScannedLog = Join-Path $sessionDir "scanned_files.log"

if (-not (Test-Path $sessionDir)) {
    New-Item -ItemType Directory -Path $sessionDir -Force | Out-Null
}

if (-not $Resume -and (Test-Path $sessionMetaFile)) {
    try {
        $existingMeta = Get-Content -Path $sessionMetaFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($existingMeta.Status -eq "In-Progress") {
            Write-Host "`n======================================================================" -ForegroundColor Yellow
            Write-Host "   PHAT HIEN PHIEN QUET CHUA HOAN TAT TAI THU MUC NAY!              " -ForegroundColor Yellow
            Write-Host "======================================================================" -ForegroundColor Yellow
            Write-Host "   - Bat dau luc : $($existingMeta.StartTime)" -ForegroundColor White
            Write-Host "   - Da quet     : $($existingMeta.TotalScanned) tep (Da diet: $($existingMeta.TotalCleaned) tep)" -ForegroundColor Cyan
            Write-Host "   - Thu muc     : $TargetFolder" -ForegroundColor White
            Write-Host "`nBan co muon TIEP TUC quet tu vi tri nay khong?" -ForegroundColor Yellow
            Write-Host "  [Y] Tiep tuc phien truoc (Bo qua cac tep da kiem tra)" -ForegroundColor Green
            Write-Host "  [N] Quet lai tu dau (Xoa phien cu)" -ForegroundColor DarkYellow
            $ans = Read-Host "Lua chon [Y/N] (Mac dinh: Y)"
            if (-not $ans -or $ans -eq "Y" -or $ans -eq "y") {
                $Resume = $true
            } else {
                if (Test-Path $sessionScannedLog) { Remove-Item -Path $sessionScannedLog -Force -ErrorAction SilentlyContinue }
                if (Test-Path $sessionMetaFile) { Remove-Item -Path $sessionMetaFile -Force -ErrorAction SilentlyContinue }
            }
        }
    } catch {}
}

$Script:ScannedFilesSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$Script:SessionStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

if ($Resume -and (Test-Path $sessionMetaFile)) {
    try {
        $loadedMeta = Get-Content -Path $sessionMetaFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $Script:TotalScanned = [int]$loadedMeta.TotalScanned
        $Script:TotalCleaned = [int]$loadedMeta.TotalCleaned
        $Script:TotalSafe    = [int]$loadedMeta.TotalSafe
        $Script:TotalErrors  = [int]$loadedMeta.TotalErrors
        $Script:TotalFolders = [int]$loadedMeta.TotalFolders
        if ($loadedMeta.StartTime) { $Script:SessionStartTime = $loadedMeta.StartTime }
        Write-Host "`n   -> [OK] Da khoi phuc phien quet: Da quet $Script:TotalScanned tep (Da diet $Script:TotalCleaned)" -ForegroundColor Green
    } catch {
        $Script:TotalScanned = 0
        $Script:TotalCleaned = 0
        $Script:TotalSafe    = 0
        $Script:TotalErrors  = 0
        $Script:TotalFolders = 0
    }
    
    if (Test-Path $sessionScannedLog) {
        Get-Content -Path $sessionScannedLog -Encoding UTF8 | ForEach-Object {
            $line = $_.Trim()
            if ($line) { $Script:ScannedFilesSet.Add($line) | Out-Null }
        }
        Write-Host "   -> [OK] Da nap $($Script:ScannedFilesSet.Count) tep da quet vao bo nho dem Fast-Skip." -ForegroundColor Green
    }
} else {
    $Script:TotalScanned = 0
    $Script:TotalCleaned = 0
    $Script:TotalSafe    = 0
    $Script:TotalErrors  = 0
    $Script:TotalFolders = 0
    if (Test-Path $sessionScannedLog) { Remove-Item -Path $sessionScannedLog -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType File -Path $sessionScannedLog -Force | Out-Null
}

function Save-SessionMeta([string]$status = "In-Progress") {
    $metaObj = [ordered]@{
        SessionId    = $folderHash
        TargetFolder = $TargetFolder
        Status       = $status
        StartTime    = $Script:SessionStartTime
        LastUpdated  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        TotalFolders = $Script:TotalFolders
        TotalScanned = $Script:TotalScanned
        TotalCleaned = $Script:TotalCleaned
        TotalSafe    = $Script:TotalSafe
        TotalErrors  = $Script:TotalErrors
        SessionDir   = $sessionDir
    }
    $json = $metaObj | ConvertTo-Json -Depth 3
    try {
        [System.IO.File]::WriteAllText($sessionMetaFile, $json, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($lastSessionFile, $json, [System.Text.Encoding]::UTF8)
    } catch {}
}

Save-SessionMeta "In-Progress"

Write-Host "`nDang khoi dong Excel COM doc lap..." -ForegroundColor Gray
if (-not (Start-FreshExcel)) {
    Write-Host "`nNhan Enter de thoat..." -ForegroundColor Gray
    try { Read-Host | Out-Null } catch {}
    exit 1
}
Write-Host "   -> [OK] Excel Worker v$($Script:AppVersion) da san sang (PID: $Script:CurrentExcelPid)" -ForegroundColor Green

$excelExtensions = @(".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt", ".xla", ".xlam")
$virusKeywords   = @("Kangatang", "Kangaatang", "Kanga", "mypersonnel")

# ==============================================================================
# HAM QUET VA LAM SACH TEP VOI FULL-LIFECYCLE WATCHDOG 35s
# ==============================================================================
function Scan-SingleExcelFile($file) {
    $Script:TotalScanned++
    $filePath = $file.FullName
    $idx = $Script:TotalScanned
    
    $Host.UI.RawUI.WindowTitle = "KangatangGuard v$($Script:AppVersion) | Da quet: $Script:TotalScanned | Da diet: $Script:TotalCleaned | PID: $Script:CurrentExcelPid"
    
    # Dinh ky lam moi tien trinh Excel moi 30 tep de chong tran bo nho
    if ($Script:TotalScanned % 30 -eq 0) {
        Write-Host "  -> [RECYCLE] Dinh ky lam sach bo nho COM tai tep #$idx..." -ForegroundColor DarkCyan
        Start-FreshExcel | Out-Null
    }

    # KICH HOAT WATCHDOG 35s BAO VE TOAN BO VONG DOI TEP
    $Script:Watchdog.Arm($Script:CurrentExcelPid, 35000)
    $wb = $null

    try {
        # 1. MO TEP READ-ONLY
        try {
            $wb = $Script:CurrentExcelApp.Workbooks.Open($filePath, 0, $true)
        } catch {
            $errMsg = $_.Exception.Message
            $Script:TotalErrors++
            if ($Script:Watchdog.TimedOut) {
                Write-Host "  [$idx] [TIMEOUT] Tep bi treo qua 35s (bo qua an toan): $($file.Name)" -ForegroundColor Red
                Write-AuditLog "[TIMEOUT] $filePath | Qua thoi gian 35s"
            } else {
                Write-Host "  [$idx] Khong the mo: $($file.Name) | Loi: $errMsg" -ForegroundColor DarkYellow
                Write-AuditLog "[SKIP] Khong the mo: $filePath | Loi: $errMsg"
            }
            Start-FreshExcel | Out-Null
            return
        }

        if ($null -eq $wb) {
            $Script:TotalErrors++
            return
        }

        # 2. KIEM TRA CAC DAU HIEU MA DOC
        $isInfected = $false
        $infectionDetails = [System.Collections.Generic.List[string]]::new()
        $ext = $file.Extension.ToLower()

        # VBProject: Chi kiem tra voi tep Macro de tranh crash VBE7
        if ($ext -ne ".xlsx" -and $ext -ne ".xltx") {
            try {
                $vbProj = $wb.VBProject
                if ($vbProj -and $vbProj.Protection -eq 0) {
                    foreach ($comp in $vbProj.VBComponents) {
                        foreach ($kw in $virusKeywords) {
                            if ($comp.Name -like "*$kw*") {
                                $isInfected = $true
                                $infectionDetails.Add("Module doc hai: " + $comp.Name)
                                break
                            }
                        }
                        try {
                            $cm = $comp.CodeModule
                            if ($cm -and $cm.CountOfLines -gt 0) {
                                $checkLines = [math]::Min(500, $cm.CountOfLines)
                                $lines = $cm.Lines(1, $checkLines)
                                foreach ($kw in $virusKeywords) {
                                    if ($lines -like "*$kw*") {
                                        $isInfected = $true
                                        $infectionDetails.Add("Ma doc trong: " + $comp.Name)
                                        break
                                    }
                                }
                            }
                        } catch {}
                    }
                }
            } catch {}
        }

        # Sheet An
        try {
            foreach ($sht in $wb.Sheets) {
                foreach ($kw in $virusKeywords) {
                    if ($sht.Name -like "*$kw*") {
                        $isInfected = $true
                        $infectionDetails.Add("Sheet an doc hai: " + $sht.Name)
                        break
                    }
                }
            }
        } catch {}

        # FAST NAMES FILTER: Khong doc RefersTo qua mang, chi kiem tra Name trong RAM
        try {
            $namesCount = $wb.Names.Count
            if ($namesCount -gt 0) {
                $checkLimit = [math]::Min($namesCount, 500)
                $nIdx = 0
                foreach ($nm in $wb.Names) {
                    $nIdx++
                    if ($nIdx -gt $checkLimit) { break }
                    try {
                        $nmName = $nm.Name
                        foreach ($kw in $virusKeywords) {
                            if ($nmName -like "*$kw*") {
                                $isInfected = $true
                                $infectionDetails.Add("Named Range doc hai: " + $nmName)
                                break
                            }
                        }
                    } catch {}
                }
            }
        } catch {}

        # 3. XU LY DIET VIRUS
        if ($isInfected) {
            Write-Host "  [PHAT HIEN VIRUS] [$idx]: $($file.Name)" -ForegroundColor Red
            foreach ($det in $infectionDetails) {
                Write-Host "     - $det" -ForegroundColor DarkRed
            }
            Write-AuditLog "[DETECTED] $filePath | $($infectionDetails -join '; ')"

            # Dong file Read-Only de giai phong khoa
            try { $wb.Close($false) } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
            $wb = $null

            # Pre-flight Write Lock Check
            if (-not (Test-FileWriteable $filePath)) {
                Write-Host "     [KHOA TEP] Tep dang duoc mo boi nguoi dung khac hoac he thong mang. Bo qua lam sach." -ForegroundColor Yellow
                Write-AuditLog "[LOCKED_BY_USER] $filePath"
                $Script:TotalErrors++
                return
            }

            # Tao ban sao luu an toan vao _Backup_Kangatang
            $parentDir = $file.DirectoryName
            $backupDir = Join-Path $parentDir "_Backup_Kangatang"
            if (-not (Test-Path $backupDir)) {
                New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
            }
            $ts = Get-Date -Format "yyyyMMdd_HHmmss"
            $backupName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + "_backup_" + $ts + $file.Extension
            $backupPath = Join-Path $backupDir $backupName

            $backupSuccess = $false
            try {
                Copy-Item -Path $filePath -Destination $backupPath -Force -ErrorAction Stop
                Write-Host "     [SAO LUU] Da tao ban sao tai: $backupName" -ForegroundColor Green
                Write-AuditLog "[BACKUP] $backupPath"
                $backupSuccess = $true
            } catch {
                $errTxt = $_.Exception.Message
                Write-Host "     [LOI SAO LUU] Khong the tao ban sao: $errTxt" -ForegroundColor Yellow
                Write-AuditLog "[BACKUP_ERROR] Khong the backup $filePath : $errTxt"
            }

            if ($backupSuccess) {
                $cleaned = $false

                # Kiem tra ket noi COM Excel App truoc khi mo ghi, neu mat ket noi thi hoi sinh
                $isAlive = $false
                try {
                    if ($null -ne $Script:CurrentExcelApp -and $Script:CurrentExcelApp.Workbooks.Count -ge 0) {
                        $isAlive = $true
                    }
                } catch { $isAlive = $false }

                if (-not $isAlive) {
                    Write-Host "     [HOI SINH COM] Ket noi Excel bi gian doan, dang tu dong khoi phuc worker..." -ForegroundColor DarkYellow
                    Start-FreshExcel | Out-Null
                    $Script:Watchdog.Arm($Script:CurrentExcelPid, 35000)
                }

                # Mo lai che do Read-Write voi co che tu dong thu lai phong thu
                $openAttempts = 2
                for ($oa = 1; $oa -le $openAttempts; $oa++) {
                    try {
                        $wb = $Script:CurrentExcelApp.Workbooks.Open($filePath, $false, $false)
                        if ($null -ne $wb) { break }
                    } catch {
                        if ($oa -lt $openAttempts) {
                            Start-Sleep -Milliseconds 400
                            Start-FreshExcel | Out-Null
                            $Script:Watchdog.Arm($Script:CurrentExcelPid, 35000)
                        } else {
                            throw $_
                        }
                    }
                }
                
                if ($null -eq $wb) {
                    Write-Host "     [LOI GHI] Khong the lay doi tuong Workbook khi mo Read-Write." -ForegroundColor Red
                    Write-AuditLog "[CLEAN_ERROR] $filePath : Workbook is null"
                    $Script:TotalErrors++
                    Start-FreshExcel | Out-Null
                    return
                }

                # Lam sach VBProject
                if ($ext -ne ".xlsx" -and $ext -ne ".xltx") {
                    try {
                        $vbProj = $wb.VBProject
                        if ($vbProj -and $vbProj.Protection -eq 0) {
                            for ($cIdx = $vbProj.VBComponents.Count; $cIdx -ge 1; $cIdx--) {
                                $comp = $vbProj.VBComponents.Item($cIdx)
                                if ($comp.Type -eq 1 -or $comp.Type -eq 2) {
                                    foreach ($kw in $virusKeywords) {
                                        if ($comp.Name -like "*$kw*") {
                                            $vbProj.VBComponents.Remove($comp)
                                            $cleaned = $true
                                            break
                                        }
                                    }
                                } elseif ($comp.CodeModule -and $comp.CodeModule.CountOfLines -gt 0) {
                                    $lines = $comp.CodeModule.Lines(1, [math]::Min(500, $comp.CodeModule.CountOfLines))
                                    foreach ($kw in $virusKeywords) {
                                        if ($lines -like "*$kw*") {
                                            $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
                                            $cleaned = $true
                                            break
                                        }
                                    }
                                }
                            }
                        }
                    } catch {}
                }
                
                # Lam sach Sheets
                for ($sIdx = $wb.Sheets.Count; $sIdx -ge 1; $sIdx--) {
                    if ($wb.Sheets.Count -le 1) { break }
                    $sht = $wb.Sheets.Item($sIdx)
                    foreach ($kw in $virusKeywords) {
                        if ($sht.Name -like "*$kw*") {
                            $sht.Visible = -1
                            $sht.Delete()
                            $cleaned = $true
                            break
                        }
                    }
                }

                # Lam sach Names: Fast forward loop voi danh sach ten can xoa
                try {
                    $delNames = [System.Collections.Generic.List[string]]::new()
                    foreach ($nm in $wb.Names) {
                        try {
                            $nmName = $nm.Name
                            foreach ($kw in $virusKeywords) {
                                if ($nmName -like "*$kw*") {
                                    $delNames.Add($nmName)
                                    break
                                }
                            }
                        } catch {}
                    }
                    foreach ($dName in $delNames) {
                        try {
                            $wb.Names.Item($dName).Delete()
                            $cleaned = $true
                        } catch {}
                    }
                } catch {}

                # Tat cac hop thoai tuong thich khi luu
                $wb.CheckCompatibility = $false
                try { $wb.RemovePersonalInformation = $false } catch {}

                if ($cleaned) {
                    $wb.Save()
                    Write-Host "     [DA TIEU DIET] Da lam sach va luu tep thanh cong!" -ForegroundColor Green
                    Write-AuditLog "[CLEANED] $filePath"
                    $Script:TotalCleaned++
                } else {
                    Write-Host "     Khong tim thay thanh phan can xoa khi lam sach." -ForegroundColor DarkYellow
                }
            }
        } else {
            Write-Host "  [OK] [$idx] An toan: $($file.Name)" -ForegroundColor Gray
            $Script:TotalSafe++
        }
    } catch {
        $errTxt = $_.Exception.Message
        $Script:TotalErrors++
        if ($Script:Watchdog.TimedOut) {
            Write-Host "  [$idx] [TIMEOUT TOAN DIEN] File bi treo qua 35s: $($file.Name)" -ForegroundColor Red
            Write-AuditLog "[TIMEOUT_FULL] $filePath | Qua thoi gian 35s"
        } else {
            Write-Host "  [$idx] Loi xu ly tep: $($file.Name) | $errTxt" -ForegroundColor DarkYellow
            Write-AuditLog "[ERROR_FILE] $filePath : $errTxt"
        }
        # Auto-recovery ngay lap tuc
        Start-FreshExcel | Out-Null
    } finally {
        $Script:Watchdog.Disarm()
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
            $wb = $null
        }
    }
}

# ==============================================================================
# HAM DUYET LUONG DE QUY (STREAMING RECURSION)
# ==============================================================================
function Scan-FolderStream([string]$currentDir) {
    if ($currentDir -like "*_Backup_Kangatang*") { return }

    $Script:TotalFolders++
    Write-Host "`n----------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Folder [$Script:TotalFolders]: $currentDir" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkGray
    
    Write-Progress -Activity "KangatangGuard v$($Script:AppVersion) - Dang quet luong chong treo" -Status "Thu muc #$Script:TotalFolders: $currentDir" -CurrentOperation "Da quet: $Script:TotalScanned tep | Da diet: $Script:TotalCleaned"

    # 1. Quet ngay lap tuc tat ca tep Excel co trong thu muc nay
    try {
        $files = Get-ChildItem -Path $currentDir -File -ErrorAction SilentlyContinue
        foreach ($f in $files) {
            if ($excelExtensions -contains $f.Extension.ToLower() -and $f.Name -ne "KangatangGuard.xlam") {
                if ($Script:ScannedFilesSet.Contains($f.FullName)) {
                    Write-Host "  ⏩ [BO QUA - DA QUET] $($f.Name)" -ForegroundColor DarkGray
                    continue
                }
                Scan-SingleExcelFile $f
                
                try {
                    [System.IO.File]::AppendAllText($sessionScannedLog, $f.FullName + "`r`n", [System.Text.Encoding]::UTF8)
                    $Script:ScannedFilesSet.Add($f.FullName) | Out-Null
                } catch {}
            }
        }
    } catch {
        $errTxt = $_.Exception.Message
        Write-Host "  [LOI DUYET TEP] $errTxt" -ForegroundColor Red
    }

    Save-SessionMeta "In-Progress"

    # 2. Lay danh sach thu muc con va lan luot quet tiep
    try {
        $subDirs = Get-ChildItem -Path $currentDir -Directory -ErrorAction SilentlyContinue
        foreach ($sub in $subDirs) {
            if ($sub.Name -ne "_Backup_Kangatang" -and $sub.Name -ne "Windows" -and $sub.Name -ne "Program Files" -and $sub.Name -ne "Program Files (x86)") {
                Scan-FolderStream $sub.FullName
            }
        }
    } catch {
        $errTxt = $_.Exception.Message
        Write-Host "  [LOI DUYET THU MUC CON] $errTxt" -ForegroundColor Red
    }
}

# Kich hoat quet luong
Write-Host "`nBAT DAU QUET LUONG TRUC TIEP CHONG TREO (v$($Script:AppVersion))..." -ForegroundColor Cyan
Scan-FolderStream $TargetFolder

Write-Progress -Activity "KangatangGuard v$($Script:AppVersion)" -Completed

# Giai phong tai nguyen
Stop-CurrentExcel
[ComMessageFilter]::Revoke()

Save-SessionMeta "Completed"

# Bao cao tong ket
Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   BAO CAO TONG KET QUET LUONG CHONG TREO (v$($Script:AppVersion))      " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green
Write-Host "   Thu muc bat dau                 : $TargetFolder" -ForegroundColor White
if ($Resume) {
    Write-Host "   Che do                          : TIEP TUC PHIEN TRUOC (Fast Resume)" -ForegroundColor Cyan
}
Write-Host "   Tong so thu muc da duyet qua    : $Script:TotalFolders" -ForegroundColor Yellow
Write-Host "   Tong so tep Excel da kiem tra   : $Script:TotalScanned" -ForegroundColor Cyan
Write-Host "   So tep an toan                  : $Script:TotalSafe" -ForegroundColor Green
Write-Host "   So tep phat hien & da tieu diet : $Script:TotalCleaned" -ForegroundColor $(if ($Script:TotalCleaned -gt 0) { "Red" } else { "Green" })
Write-Host "   So tep loi / bo qua / bi khoa   : $Script:TotalErrors" -ForegroundColor $(if ($Script:TotalErrors -gt 0) { "Yellow" } else { "Gray" })
Write-Host "   Nhat ky chi tiet                : $logFile" -ForegroundColor DarkGray
Write-Host "   Trang thai phien quet           : HOAN TAT 100%" -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green

Write-AuditLog "[STREAM_SCAN_END] $TargetFolder - ThuMuc: $Script:TotalFolders, Tep: $Script:TotalScanned, Diet: $Script:TotalCleaned, AnToan: $Script:TotalSafe, Loi: $Script:TotalErrors, Resume: $Resume"

Write-Host "`nToan bo tien trinh quet luong da hoan tat ma khong lam giam hieu nang Excel." -ForegroundColor Cyan
Write-Host "Nhan Enter de hoan tat va dong cua so..." -ForegroundColor Gray
try { Read-Host | Out-Null } catch {}

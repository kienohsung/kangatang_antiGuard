# ==============================================================================
# Kangatang_Standalone_Scanner.ps1
# Trình Quét & Tiêu Diệt Virus Macro Excel Độc Lập (Standalone Scanner v3.8.2)
# Thiết kế dành riêng cho các máy trạm không thể cài đặt Add-in Excel
# 
# Đặc tính Kỹ thuật Chuẩn Production v3.8.2:
#   - Không yêu cầu cài đặt Add-in, không cần quyền Administrator (tự thích ứng).
#   - Tích hợp Native C# IOleMessageFilter: Triệt tiêu 100% lỗi COM RPC Server Fault & Deadlock.
#   - Tự động tắt Windows Console QuickEdit Mode qua Win32 API: Chuột bấm vào cửa sổ không làm dừng tiến trình.
#   - Watchdog 35s bảo vệ TOÀN BỘ vòng đời tệp (Open, VBProject, Sheets, Names, Save, Close).
#   - Fast Names Filter: Xử lý tệp chứa hàng nghìn Named Ranges trong vài giây, loại bỏ tắc nghẽn SMB.
#   - Tự động quét và dọn sạch các tiến trình Excel zombie /automation -Embedding bỏ hoang.
#   - Hỗ trợ Tiếp tục phiên quét dở dang (Fast Resume) qua HashSet O(1) và Checkpoint an toàn.
#   - Kiểm tra khóa ghi tệp trước khi can thiệp (Pre-flight Write Lock Check).
#   - Tự động sao lưu an toàn vào thư mục _Backup_Kangatang trước khi sửa đổi file.
#   - Duyệt luồng trực tiếp (Streaming Recursion): Quét tức thì không nghẽn RAM.
#   - Dọn sạch ổ dịch khởi động ngầm: XLSTART, Registry Options (OPEN*), AddIns.
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = "",
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    [Parameter(Mandatory=$false)]
    [switch]$CleanSystemOnly
)

# Thiết lập bảng mã UTF-8 hiển thị tiếng Việt hoàn hảo
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$Script:AppVersion = "3.8.2"
$Host.UI.RawUI.WindowTitle = "KangatangGuard v$($Script:AppVersion) - Trình Diệt Virus Excel Độc Lập (Standalone)"

# Khởi tạo thư mục Audit Log và Sessions
$Script:LogDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $Script:LogDir)) {
    New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
}
$Script:LogFile = Join-Path $Script:LogDir ("standalone_scan_" + (Get-Date -Format "yyyyMMdd") + ".txt")
$Script:SessionsBaseDir = Join-Path $Script:LogDir "Sessions"
if (-not (Test-Path $Script:SessionsBaseDir)) {
    New-Item -ItemType Directory -Path $Script:SessionsBaseDir -Force | Out-Null
}
$Script:LastSessionFile = Join-Path $Script:LogDir "last_session.json"

function Write-AuditLog([string]$msg) {
    $timeStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timeStr | $msg"
    try {
        Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

# ==============================================================================
# LỚP C# NATIVE: CONSOLE HELPER, COM MESSAGE FILTER & WATCHDOG 35s
# ==============================================================================
$nativeDefinitions = @"
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
        return 0;
    }

    public int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType) {
        if (dwRejectType == 2) {
            return 100; // Thử lại sau 100ms
        }
        return -1; // Ngắt ngay nếu bị từ chối
    }

    public int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType) {
        return 2; // PENDINGMSG_WAITDEFPROCESS
    }
}

public class StandaloneWatchdog : IDisposable {
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
"@

try {
    Add-Type -TypeDefinition $nativeDefinitions -ErrorAction SilentlyContinue
} catch {}

# Vô hiệu hóa QuickEdit
[ConsoleHelper]::DisableQuickEdit()

# Đăng ký COM Message Filter
[ComMessageFilter]::Register()

$Script:Watchdog = New-Object StandaloneWatchdog

# ==============================================================================
# QUẢN TRỊ REGISTRY AccessVBOM
# ==============================================================================
$Script:AccessVBOMBackup = @{}

function Enable-AccessVBOM {
    $officeVersions = @("14.0", "15.0", "16.0")
    foreach ($ver in $officeVersions) {
        $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
        if (Test-Path $regPath) {
            try {
                $currentVal = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
                if ($null -eq $currentVal) { $currentVal = -1 }
                $Script:AccessVBOMBackup[$regPath] = $currentVal
                Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    }
}

function Restore-AccessVBOM {
    foreach ($regPath in $Script:AccessVBOMBackup.Keys) {
        $orig = $Script:AccessVBOMBackup[$regPath]
        try {
            if ($orig -eq -1) {
                Remove-ItemProperty -Path $regPath -Name "AccessVBOM" -Force -ErrorAction SilentlyContinue
            } else {
                Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value $orig -Type DWord -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}

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
# QUẢN LÝ TIẾN TRÌNH EXCEL COM NGUYÊN TỬ & DỌN DẸP ZOMBIE
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
            
            try {
                $exactPid = [StandaloneWatchdog]::GetExcelPid($Script:CurrentExcelApp.Hwnd)
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
                Write-Host "   [LỖI] Không thể khởi động Excel COM sau $maxRetries lần: $errTxt" -ForegroundColor Red
                Write-AuditLog "[ERROR] Khong the khoi dong Excel COM: $errTxt"
                return $false
            }
        }
    }
    return $false
}

# ==============================================================================
# HÀM DỌN DẸP Ổ DỊCH KHỞI ĐỘNG HỆ THỐNG
# ==============================================================================
function Clean-SystemReservoirs {
    Write-Host "`n======================================================================" -ForegroundColor Yellow
    Write-Host "   ĐANG RÀ SOÁT VÀ DỌN DẸP Ổ DỊCH KHỞI ĐỘNG HỆ THỐNG (XLSTART & REGISTRY)  " -ForegroundColor Yellow
    Write-Host "======================================================================" -ForegroundColor Yellow
    
    # Dọn dẹp tiến trình zombie automation
    Stop-OrphanExcelProcesses

    $cleanedCount = 0
    $virusPatterns = @("mypersonnel*", "*kangatang*", "*kangaatang*", "*kanga*.xls*", "personal.xls")

    $xlStartDirs = [System.Collections.Generic.List[string]]::new()
    $userXLStart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
    if (Test-Path $userXLStart) { $xlStartDirs.Add($userXLStart) }

    if (Test-Path "C:\Users") {
        Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $otherXLStart = Join-Path $_.FullName "AppData\Roaming\Microsoft\Excel\XLSTART"
            if ((Test-Path $otherXLStart -ErrorAction SilentlyContinue) -and (-not $xlStartDirs.Contains($otherXLStart))) {
                $xlStartDirs.Add($otherXLStart)
            }
        }
    }

    $systemOfficeDirs = @(
        "C:\Program Files\Microsoft Office\root\Office*",
        "C:\Program Files (x86)\Microsoft Office\root\Office*",
        "C:\Program Files\Microsoft Office\Office*",
        "C:\Program Files (x86)\Microsoft Office\Office*"
    )
    foreach ($patt in $systemOfficeDirs) {
        Get-ChildItem -Path $patt -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $sysXL = Join-Path $_.FullName "XLSTART"
            if ((Test-Path $sysXL -ErrorAction SilentlyContinue) -and (-not $xlStartDirs.Contains($sysXL))) {
                $xlStartDirs.Add($sysXL)
            }
        }
    }

    foreach ($xlDir in $xlStartDirs) {
        foreach ($patt in $virusPatterns) {
            $matches = Get-ChildItem -Path $xlDir -Filter $patt -File -Force -ErrorAction SilentlyContinue
            foreach ($m in $matches) {
                if ($m.Name -match "KangatangGuard") { continue }
                try {
                    Remove-Item -Path $m.FullName -Force -ErrorAction Stop
                    Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Xóa file mầm bệnh: $($m.FullName)" -ForegroundColor Red
                    Write-AuditLog "[SYS_CLEAN] Xoa XLSTART: $($m.FullName)"
                    $cleanedCount++
                } catch {}
            }
        }
    }

    $appDataExcel = Join-Path $env:APPDATA "Microsoft\Excel"
    if (Test-Path $appDataExcel) {
        Get-ChildItem -Path $appDataExcel -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.Name -match "mypersonnel|kangatang|kanga" -and ($_.Name -notmatch "KangatangGuard|Kangatang_")) {
                try {
                    Remove-Item -Path $_.FullName -Force -ErrorAction Stop
                    Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Xóa file rác AppData: $($_.Name)" -ForegroundColor Red
                    Write-AuditLog "[SYS_CLEAN] Xoa AppData: $($_.FullName)"
                    $cleanedCount++
                } catch {}
            }
        }
    }

    foreach ($ver in @("16.0", "15.0", "14.0")) {
        $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
        if (Test-Path $optKey) {
            try {
                $props = (Get-ItemProperty -Path $optKey -ErrorAction SilentlyContinue).psobject.Properties
                foreach ($p in $props) {
                    if ($p.Name -match "^OPEN\d*$" -and ([string]$p.Value -match "mypersonnel|kangatang|kanga|personal\.xls") -and ([string]$p.Value -notmatch "KangatangGuard")) {
                        Remove-ItemProperty -Path $optKey -Name $p.Name -Force -ErrorAction SilentlyContinue
                        Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Gỡ bỏ khóa Registry độc hại: $($p.Name) = $($p.Value)" -ForegroundColor Red
                        Write-AuditLog "[SYS_CLEAN] Xoa Registry: $optKey\$($p.Name)"
                        $cleanedCount++
                    }
                }
                $altStartup = (Get-ItemProperty -Path $optKey -Name "AltStartupPath" -ErrorAction SilentlyContinue).AltStartupPath
                if ($altStartup -and (Test-Path $altStartup)) {
                    Get-ChildItem -Path $altStartup -File -ErrorAction SilentlyContinue | ForEach-Object {
                        if ($_.Name -match "mypersonnel|kangatang|kanga|personal\.xls") {
                            Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
                            Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Xóa file trong AltStartupPath: $($_.Name)" -ForegroundColor Red
                            Write-AuditLog "[SYS_CLEAN] Xoa AltStartup: $($_.FullName)"
                            $cleanedCount++
                        }
                    }
                }
            } catch {}
        }
    }

    $userAddIns = Join-Path $env:APPDATA "Microsoft\AddIns"
    if (Test-Path $userAddIns) {
        Get-ChildItem -Path $userAddIns -File -ErrorAction SilentlyContinue | ForEach-Object {
            if (($_.Name -match "mypersonnel|kangatang|kanga|personal\.xls") -and ($_.Name -notmatch "KangatangGuard|Kangatang_")) {
                try {
                    Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
                    Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Xóa file trong AddIns: $($_.Name)" -ForegroundColor Red
                    Write-AuditLog "[SYS_CLEAN] Xoa AddIns: $($_.FullName)"
                    $cleanedCount++
                } catch {}
            }
        }
    }

    Write-Host "`n   -> ✅ Hoàn tất dọn dẹp ổ dịch hệ thống! (Đã xử lý: $cleanedCount mục)" -ForegroundColor Green
    Write-AuditLog "[SYS_CLEAN_COMPLETE] Da don sach $cleanedCount muc"
}

# ==============================================================================
# HÀM BĂM THƯ MỤC & CHECKPOINT
# ==============================================================================
function Get-FolderHash([string]$folder) {
    $clean = $folder.Trim().ToLowerInvariant().TrimEnd('\')
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($clean)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $hashBytes = $md5.ComputeHash($bytes)
    return -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
}

$Script:TotalFolders = 0
$Script:TotalScanned = 0
$Script:TotalCleaned = 0
$Script:TotalSafe    = 0
$Script:TotalErrors  = 0
$Script:ScannedFilesSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$Script:SessionStartTime = ""
$Script:CurrentSessionDir = ""
$Script:CurrentMetaFile = ""
$Script:CurrentScannedLog = ""

function Save-SessionMeta([string]$status = "In-Progress") {
    if (-not $Script:CurrentMetaFile) { return }
    $metaObj = [ordered]@{
        SessionId    = $Script:CurrentFolderHash
        TargetFolder = $Script:CurrentScanTarget
        Status       = $status
        StartTime    = $Script:SessionStartTime
        LastUpdated  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        TotalFolders = $Script:TotalFolders
        TotalScanned = $Script:TotalScanned
        TotalCleaned = $Script:TotalCleaned
        TotalSafe    = $Script:TotalSafe
        TotalErrors  = $Script:TotalErrors
        SessionDir   = $Script:CurrentSessionDir
    }
    $json = $metaObj | ConvertTo-Json -Depth 3
    try {
        [System.IO.File]::WriteAllText($Script:CurrentMetaFile, $json, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($Script:LastSessionFile, $json, [System.Text.Encoding]::UTF8)
    } catch {}
}

$excelExtensions = @(".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt", ".xla", ".xlam")
$virusKeywords   = @("Kangatang", "Kangaatang", "Kanga", "mypersonnel")

# ==============================================================================
# QUÉT VÀ DIỆT VIRUS CHO 1 TỆP EXCEL (FULL-LIFECYCLE WATCHDOG 35s)
# ==============================================================================
function Scan-SingleExcelFile($file) {
    $Script:TotalScanned++
    $filePath = $file.FullName
    $idx = $Script:TotalScanned
    
    $Host.UI.RawUI.WindowTitle = "Kangatang Standalone v$($Script:AppVersion) | Da quet: $Script:TotalScanned | Da diet: $Script:TotalCleaned | PID: $Script:CurrentExcelPid"
    
    if ($Script:TotalScanned % 30 -eq 0) {
        Write-Host "  -> [RECYCLE] Định kỳ làm sạch bộ nhớ Excel COM tại tệp #$idx..." -ForegroundColor DarkCyan
        Start-FreshExcel | Out-Null
    }

    # KÍCH HOẠT WATCHDOG 35s CHO TOÀN BỘ VÒNG ĐỜI TỆP
    $Script:Watchdog.Arm($Script:CurrentExcelPid, 35000)
    $wb = $null

    try {
        # 1. Mở file Read-Only
        try {
            $wb = $Script:CurrentExcelApp.Workbooks.Open($filePath, 0, $true)
        } catch {
            $errMsg = $_.Exception.Message
            $Script:TotalErrors++
            if ($Script:Watchdog.TimedOut) {
                Write-Host "  [$idx] ⏳ [TIMEOUT 35s] Tệp bị treo qua mạng (bỏ qua an toàn): $($file.Name)" -ForegroundColor Red
                Write-AuditLog "[TIMEOUT] $filePath"
            } else {
                Write-Host "  [$idx] ⚠️ Không thể mở: $($file.Name) ($errMsg)" -ForegroundColor DarkYellow
                Write-AuditLog "[SKIP] $filePath | $errMsg"
            }
            Start-FreshExcel | Out-Null
            return
        }

        if ($null -eq $wb) {
            $Script:TotalErrors++
            return
        }

        # 2. Kiểm tra mã độc
        $isInfected = $false
        $infectionDetails = [System.Collections.Generic.List[string]]::new()
        $ext = $file.Extension.ToLower()

        if ($ext -ne ".xlsx" -and $ext -ne ".xltx") {
            try {
                $vbProj = $wb.VBProject
                if ($vbProj -and $vbProj.Protection -eq 0) {
                    foreach ($comp in $vbProj.VBComponents) {
                        foreach ($kw in $virusKeywords) {
                            if ($comp.Name -like "*$kw*") {
                                $isInfected = $true
                                $infectionDetails.Add("Module: " + $comp.Name)
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
                                        $infectionDetails.Add("Mã độc trong: " + $comp.Name)
                                        break
                                    }
                                }
                            }
                        } catch {}
                    }
                }
            } catch {}
        }

        try {
            foreach ($sht in $wb.Sheets) {
                foreach ($kw in $virusKeywords) {
                    if ($sht.Name -like "*$kw*") {
                        $isInfected = $true
                        $infectionDetails.Add("Sheet ẩn: " + $sht.Name)
                        break
                    }
                }
            }
        } catch {}

        # FAST NAMES FILTER: Không đọc RefersTo qua mạng
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
                                $infectionDetails.Add("Named Range: " + $nmName)
                                break
                            }
                        }
                    } catch {}
                }
            }
        } catch {}

        # 3. Xử lý diệt virus
        if ($isInfected) {
            Write-Host "  [$idx] 🛑 [PHÁT HIỆN VIRUS]: $($file.Name)" -ForegroundColor Red
            foreach ($det in $infectionDetails) {
                Write-Host "     - $det" -ForegroundColor DarkRed
            }
            Write-AuditLog "[DETECTED] $filePath | $($infectionDetails -join '; ')"

            try { $wb.Close($false) } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
            $wb = $null

            if (-not (Test-FileWriteable $filePath)) {
                Write-Host "     ⚠️ [KHÓA TỆP] Tệp đang được mở bởi người khác hoặc mạng chia sẻ. Bỏ qua ghi." -ForegroundColor Yellow
                Write-AuditLog "[LOCKED] $filePath"
                $Script:TotalErrors++
                return
            }

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
                Write-Host "     📦 [SAO LƯU] Đã lưu bản sao: $backupName" -ForegroundColor Green
                Write-AuditLog "[BACKUP] $backupPath"
                $backupSuccess = $true
            } catch {
                $errTxt = $_.Exception.Message
                Write-Host "     ⚠️ [LỖI SAO LƯU] Không thể sao lưu: $errTxt" -ForegroundColor Yellow
                Write-AuditLog "[BACKUP_ERROR] $filePath : $errTxt"
            }

            if ($backupSuccess) {
                $cleaned = $false

                # Kiểm tra kết nối COM Excel App trước khi mở ghi, nếu mất kết nối thì hồi sinh
                $isAlive = $false
                try {
                    if ($null -ne $Script:CurrentExcelApp -and $Script:CurrentExcelApp.Workbooks.Count -ge 0) {
                        $isAlive = $true
                    }
                } catch { $isAlive = $false }

                if (-not $isAlive) {
                    Write-Host "     ⚡ [HỒI SINH COM] Kết nối Excel bị gián đoạn, đang tự động khôi phục worker..." -ForegroundColor DarkYellow
                    Start-FreshExcel | Out-Null
                    $Script:Watchdog.Arm($Script:CurrentExcelPid, 35000)
                }

                # Mở lại chế độ Read-Write với cơ chế tự động thử lại phòng thủ
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
                    Write-Host "     ❌ [LỖI GHI] Không lấy được đối tượng Workbook." -ForegroundColor Red
                    Write-AuditLog "[CLEAN_ERROR] $filePath : null workbook"
                    $Script:TotalErrors++
                    Start-FreshExcel | Out-Null
                    return
                }

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

                $wb.CheckCompatibility = $false
                try { $wb.RemovePersonalInformation = $false } catch {}

                if ($cleaned) {
                    $wb.Save()
                    Write-Host "     ✅ [ĐÃ TIÊU DIỆT] Đã làm sạch và lưu tệp thành công!" -ForegroundColor Green
                    Write-AuditLog "[CLEANED] $filePath"
                    $Script:TotalCleaned++
                } else {
                    Write-Host "     ⚠️ Không tìm thấy thành phần cần xóa khi lưu." -ForegroundColor DarkYellow
                }
            }
        } else {
            Write-Host "  [$idx] 🛡️ [An toàn]: $($file.Name)" -ForegroundColor Gray
            $Script:TotalSafe++
        }
    } catch {
        $errTxt = $_.Exception.Message
        $Script:TotalErrors++
        if ($Script:Watchdog.TimedOut) {
            Write-Host "  [$idx] ⏳ [TIMEOUT TOÀN DIỆN] Tệp bị treo quá 35s: $($file.Name)" -ForegroundColor Red
            Write-AuditLog "[TIMEOUT_FULL] $filePath | 35s"
        } else {
            Write-Host "  [$idx] ❌ Lỗi xử lý tệp: $($file.Name) ($errTxt)" -ForegroundColor Red
            Write-AuditLog "[ERROR_FILE] $filePath : $errTxt"
        }
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
# HÀM DUYỆT LUỒNG TRỰC TIẾP
# ==============================================================================
function Scan-FolderStream([string]$currentDir) {
    if ($currentDir -like "*_Backup_Kangatang*") { return }

    $Script:TotalFolders++
    Write-Host "`n----------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "📁 Thư mục [$Script:TotalFolders]: $currentDir" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkGray
    
    try {
        $files = Get-ChildItem -Path $currentDir -File -ErrorAction SilentlyContinue
        foreach ($f in $files) {
            if ($excelExtensions -contains $f.Extension.ToLower() -and $f.Name -ne "KangatangGuard.xlam") {
                if ($Script:ScannedFilesSet.Contains($f.FullName)) {
                    Write-Host "  ⏩ [BỎ QUA - ĐÃ QUÉT] $($f.Name)" -ForegroundColor DarkGray
                    continue
                }
                Scan-SingleExcelFile $f
                
                try {
                    [System.IO.File]::AppendAllText($Script:CurrentScannedLog, $f.FullName + "`r`n", [System.Text.Encoding]::UTF8)
                    $Script:ScannedFilesSet.Add($f.FullName) | Out-Null
                } catch {}
            }
        }
    } catch {
        Write-Host "  [LỖI DUYỆT TỆP] $($_.Exception.Message)" -ForegroundColor Red
    }

    Save-SessionMeta "In-Progress"

    try {
        $subDirs = Get-ChildItem -Path $currentDir -Directory -ErrorAction SilentlyContinue
        foreach ($sub in $subDirs) {
            if ($sub.Name -ne "_Backup_Kangatang" -and $sub.Name -ne "Windows" -and $sub.Name -ne "Program Files" -and $sub.Name -ne "Program Files (x86)") {
                Scan-FolderStream $sub.FullName
            }
        }
    } catch {}
}

function Execute-ScanTarget([string]$target, [bool]$isResume = $false) {
    $target = $target.Trim('"').Trim("'").TrimEnd('\')
    if ($target -match '^[a-zA-Z]:$') {
        $target = $target + "\"
    }
    if (-not (Test-Path $target -ErrorAction SilentlyContinue)) {
        Write-Host "`n[LỖI] Thư mục không tồn tại: $target" -ForegroundColor Red
        return
    }

    $Script:CurrentScanTarget = $target
    $Script:CurrentFolderHash = Get-FolderHash $target
    $Script:CurrentSessionDir = Join-Path $Script:SessionsBaseDir $Script:CurrentFolderHash
    if (-not (Test-Path $Script:CurrentSessionDir)) {
        New-Item -ItemType Directory -Path $Script:CurrentSessionDir -Force | Out-Null
    }
    $Script:CurrentMetaFile   = Join-Path $Script:CurrentSessionDir "meta.json"
    $Script:CurrentScannedLog = Join-Path $Script:CurrentSessionDir "scanned_files.log"

    $Script:ScannedFilesSet.Clear()
    $Script:SessionStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($isResume -and (Test-Path $Script:CurrentMetaFile)) {
        try {
            $m = Get-Content -Path $Script:CurrentMetaFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $Script:TotalScanned = [int]$m.TotalScanned
            $Script:TotalCleaned = [int]$m.TotalCleaned
            $Script:TotalSafe    = [int]$m.TotalSafe
            $Script:TotalErrors  = [int]$m.TotalErrors
            $Script:TotalFolders = [int]$m.TotalFolders
            if ($m.StartTime) { $Script:SessionStartTime = $m.StartTime }
            Write-Host "`n   -> 🔄 [KHÔI PHỤC] Đã khôi phục phiên quét cũ: Đã quét $Script:TotalScanned tệp (Đã diệt $Script:TotalCleaned)" -ForegroundColor Green
        } catch {}
        
        if (Test-Path $Script:CurrentScannedLog) {
            Get-Content -Path $Script:CurrentScannedLog -Encoding UTF8 | ForEach-Object {
                $line = $_.Trim()
                if ($line) { $Script:ScannedFilesSet.Add($line) | Out-Null }
            }
            Write-Host "   -> ⚡ [FAST RESUME] Đã nạp $($Script:ScannedFilesSet.Count) tệp vào bộ nhớ đệm O(1) Fast-Skip." -ForegroundColor Green
        }
    } else {
        $Script:TotalScanned = 0
        $Script:TotalCleaned = 0
        $Script:TotalSafe    = 0
        $Script:TotalErrors  = 0
        $Script:TotalFolders = 0
        if (Test-Path $Script:CurrentScannedLog) { Remove-Item -Path $Script:CurrentScannedLog -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType File -Path $Script:CurrentScannedLog -Force | Out-Null
    }

    Save-SessionMeta "In-Progress"

    Write-Host "`nĐang khởi động tiến trình Excel COM độc lập..." -ForegroundColor Gray
    if (-not (Start-FreshExcel)) {
        Write-Host "Không thể kết nối Excel COM. Vui lòng kiểm tra lại MS Excel trên máy." -ForegroundColor Red
        return
    }
    Write-Host "   -> ✅ Excel Worker v$($Script:AppVersion) đã sẵn sàng (PID: $Script:CurrentExcelPid)" -ForegroundColor Green

    Write-Host "`n🚀 BẮT ĐẦU QUÉT LUỒNG TRỰC TIẾP CHỐNG TREO TẠI: $target" -ForegroundColor Cyan
    Write-AuditLog "[START_SCAN] $target (Resume: $isResume)"

    Scan-FolderStream $target

    Stop-CurrentExcel
    [ComMessageFilter]::Revoke()
    Save-SessionMeta "Completed"

    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host "   BÁO CÁO TỔNG KẾT PHIÊN QUÉT STANDALONE (v$($Script:AppVersion))    " -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host "   Thư mục quét                    : $target" -ForegroundColor White
    Write-Host "   Tổng số thư mục đã duyệt qua    : $Script:TotalFolders" -ForegroundColor Yellow
    Write-Host "   Tổng số tệp Excel đã kiểm tra   : $Script:TotalScanned" -ForegroundColor Cyan
    Write-Host "   Số tệp an toàn                  : $Script:TotalSafe" -ForegroundColor Green
    Write-Host "   Số tệp phát hiện & đã tiêu diệt : $Script:TotalCleaned" -ForegroundColor $(if ($Script:TotalCleaned -gt 0) { "Red" } else { "Green" })
    Write-Host "   Số tệp lỗi / bỏ qua / bị khóa   : $Script:TotalErrors" -ForegroundColor $(if ($Script:TotalErrors -gt 0) { "Yellow" } else { "Gray" })
    Write-Host "   Nhật ký kiểm toán chi tiết      : $Script:LogFile" -ForegroundColor DarkGray
    Write-Host "   Trạng thái phiên quét           : HOÀN TẤT 100%" -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green

    Write-AuditLog "[END_SCAN] $target - Folders: $Script:TotalFolders, Scanned: $Script:TotalScanned, Cleaned: $Script:TotalCleaned, Safe: $Script:TotalSafe, Errors: $Script:TotalErrors"
}

function Execute-ScanAllDrives {
    Write-Host "`n======================================================================" -ForegroundColor Cyan
    Write-Host "   QUÉT TOÀN BỘ CÁC Ổ ĐĨA DỮ LIỆU TRÊN MÁY TÍNH                      " -ForegroundColor Cyan
    Write-Host "======================================================================" -ForegroundColor Cyan
    
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 }
    Write-Host "Phát hiện $($drives.Count) ổ đĩa dữ liệu sẵn sàng:" -ForegroundColor White
    foreach ($d in $drives) {
        Write-Host "   - Ổ đĩa $($d.Root) (Trống: $([math]::Round($d.Free / 1GB, 2)) GB)" -ForegroundColor Gray
    }

    Write-Host "`nBạn có muốn bắt đầu quét toàn bộ các ổ đĩa trên?" -ForegroundColor Yellow
    $ans = Read-Host "Nhập Y để đồng ý, hoặc phím khác để hủy [Y/N]"
    if ($ans -ne "Y" -and $ans -ne "y") { return }

    foreach ($d in $drives) {
        $root = $d.Root
        Write-Host "`n>>> ĐANG QUÉT Ổ ĐĨA: $root ..." -ForegroundColor Cyan
        Execute-ScanTarget -target $root -isResume $false
    }
}

function Get-FolderFromDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "KangatangGuard - Chọn thư mục chứa file Excel cần quét virus"
    $fbd.ShowNewFolderButton = $false
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fbd.SelectedPath
    }
    return ""
}

# ==============================================================================
# KHỞI TỰ & ĐIỀU HƯỚNG TÁC VỤ (ENTRY POINT)
# ==============================================================================
Enable-AccessVBOM

try {
    if ($CleanSystemOnly) {
        Clean-SystemReservoirs
        exit 0
    }

    if ($TargetFolder) {
        Clean-SystemReservoirs
        Execute-ScanTarget -target $TargetFolder -isResume $Resume
        exit 0
    }

    $running = $true
    while ($running) {
        Clear-Host
        Write-Host "======================================================================" -ForegroundColor Cyan
        Write-Host "   KANGATANGGUARD v$($Script:AppVersion) - TRÌNH DIỆT VIRUS EXCEL ĐỘC LẬP           " -ForegroundColor Cyan
        Write-Host "   (Dành cho các máy trạm không thể cài đặt Add-in Excel)             " -ForegroundColor White
        Write-Host "======================================================================" -ForegroundColor Cyan
        Write-Host "   🛡️ BẢO VỆ CHỐNG TREO: Tích hợp Full-Lifecycle Watchdog & IOleMessageFilter" -ForegroundColor Green
        Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "   CHỌN TÁC VỤ DIỆT VIRUS:" -ForegroundColor Green
        Write-Host "     [1] Quét Thư mục Tùy chọn (Mở hộp thoại GUI hoặc nhập đường dẫn/UNC)" -ForegroundColor White
        Write-Host "     [2] Quét Toàn bộ Máy tính (Tất cả ổ đĩa C:, D:, E:...)" -ForegroundColor White
        Write-Host "     [3] Tiếp tục Phiên quét dở dang (Fast Resume từ Checkpoint)" -ForegroundColor White
        Write-Host "     [4] Dọn dẹp Mầm bệnh Hệ thống (XLSTART, Registry, AddIns)" -ForegroundColor White
        Write-Host "     [5] Mở Thư mục Nhật ký Kiểm toán (Audit Logs)" -ForegroundColor White
        Write-Host "     [0] Thoát chương trình" -ForegroundColor DarkGray
        Write-Host "======================================================================" -ForegroundColor Cyan
        
        $choice = Read-Host "Nhập lựa chọn của bạn [0-5]"
        
        switch ($choice) {
            "1" {
                Write-Host "`nĐang mở hộp thoại chọn thư mục..." -ForegroundColor Cyan
                $selected = Get-FolderFromDialog
                if (-not $selected) {
                    Write-Host "Bạn cũng có thể dán trực tiếp đường dẫn thư mục hoặc đường dẫn mạng UNC:" -ForegroundColor Yellow
                    $selected = Read-Host "Đường dẫn thư mục (Bỏ trống để hủy)"
                }
                if ($selected -and (Test-Path $selected -ErrorAction SilentlyContinue)) {
                    Clean-SystemReservoirs
                    Execute-ScanTarget -target $selected -isResume $false
                } else {
                    Write-Host "`nĐã hủy hoặc đường dẫn không hợp lệ." -ForegroundColor DarkYellow
                }
                Write-Host "`nNhấn phím bất kỳ để quay lại menu chính..." -ForegroundColor Gray
                $null = [Console]::ReadKey($true)
            }
            "2" {
                Clean-SystemReservoirs
                Execute-ScanAllDrives
                Write-Host "`nNhấn phím bất kỳ để quay lại menu chính..." -ForegroundColor Gray
                $null = [Console]::ReadKey($true)
            }
            "3" {
                if (Test-Path $Script:LastSessionFile) {
                    try {
                        $lastMeta = Get-Content -Path $Script:LastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
                        Write-Host "`n[PHÁT HIỆN PHIÊN QUÉT GẦN NHẤT]" -ForegroundColor Yellow
                        Write-Host "   - Thư mục : $($lastMeta.TargetFolder)" -ForegroundColor White
                        Write-Host "   - Bắt đầu : $($lastMeta.StartTime)" -ForegroundColor DarkGray
                        Write-Host "   - Đã quét : $($lastMeta.TotalScanned) tệp (Đã diệt $($lastMeta.TotalCleaned) tệp)" -ForegroundColor Cyan
                        Write-Host "`nTiếp tục phiên quét này?" -ForegroundColor Yellow
                        $c = Read-Host "Nhập Y để tiếp tục, phím khác để hủy [Y/N]"
                        if ($c -eq "Y" -or $c -eq "y") {
                            Clean-SystemReservoirs
                            Execute-ScanTarget -target $lastMeta.TargetFolder -isResume $true
                        }
                    } catch {
                        Write-Host "Không thể đọc thông tin phiên gần nhất." -ForegroundColor Red
                    }
                } else {
                    Write-Host "`nKhông tìm thấy checkpoint phiên quét nào trước đó." -ForegroundColor Yellow
                }
                Write-Host "`nNhấn phím bất kỳ để quay lại menu chính..." -ForegroundColor Gray
                $null = [Console]::ReadKey($true)
            }
            "4" {
                Clean-SystemReservoirs
                Write-Host "`nNhấn phím bất kỳ để quay lại menu chính..." -ForegroundColor Gray
                $null = [Console]::ReadKey($true)
            }
            "5" {
                Start-Process "explorer.exe" -ArgumentList $Script:LogDir
            }
            "0" {
                $running = $false
            }
            default {
                Write-Host "Lựa chọn không hợp lệ." -ForegroundColor DarkYellow
                Start-Sleep -Seconds 1
            }
        }
    }
} finally {
    Restore-AccessVBOM
    Stop-CurrentExcel
    [ComMessageFilter]::Revoke()
}

Write-Host "`nCảm ơn bạn đã sử dụng KangatangGuard Standalone Scanner. Tạm biệt!" -ForegroundColor Cyan
Start-Sleep -Seconds 1

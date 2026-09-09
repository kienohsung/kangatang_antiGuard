# ==============================================================================
# Kangatang_FolderScanner.ps1
# Tien trinh Quet Luong Truc Tiep & Tu dong Phuc hoi (Auto-Recovery Worker)
# Phien ban: v3.6.0 (Session Checkpointing & Fast Resume Architecture)
# Dac diem:
#   - Ho tro Tiep tuc phien quet dang do (Fast Resume) qua HashSet O(1) bo qua sieu toc.
#   - Checkpoint ben bi Append-only (chong hong du lieu tuyet doi khi sap nguon).
#   - Tich hop Native C# Watchdog (gioi han 25s/tep) chong treo 100% tren mang SMB.
#   - Kiem tra khoa ghi truoc khi diet (Pre-flight Write Lock Check) tranh dialog xung dot.
#   - Goi COM 3 tham so nguyen ban chuan chi khong gay loi marshaller/binder.
#   - Tu dong hoi sinh Excel COM ngay lap tuc ca trong luong quet va luong lam sach.
#   - Dinh ky lam moi tien trinh Excel moi 30 tep de chong ro ri bo nho.
#   - Quan ly chinh xac PID tien trinh Excel qua Win32 Hwnd API.
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

$Host.UI.RawUI.WindowTitle = "KangatangGuard v3.6.0 - Trinh quet luong chong treo & Tiep tuc phien quet"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANGGUARD v3.6.0 - FAST RESUME & ZERO-HANG STREAM SCANNER      " -ForegroundColor Cyan
Write-Host "   Kien truc Quan ly Phien Quet Dang Do & Chong Treo Mang SMB        " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

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
if ($Resume -and (-not $TargetFolder -or -not (Test-Path $TargetFolder))) {
    if (Test-Path $lastSessionFile) {
        try {
            $lastMeta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($lastMeta.TargetFolder -and (Test-Path $lastMeta.TargetFolder)) {
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

if (-not $TargetFolder -or -not (Test-Path $TargetFolder)) {
    # Kiem tra xem co phien dang do gan nhat de goi y khong
    if (Test-Path $lastSessionFile) {
        try {
            $lastMeta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($lastMeta.Status -eq "In-Progress" -and $lastMeta.TargetFolder -and (Test-Path $lastMeta.TargetFolder)) {
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
    
    if (-not $TargetFolder -or -not (Test-Path $TargetFolder)) {
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
Write-AuditLog "[STREAM_SCAN_START] Bat dau quet luong v3.6.0 tai: $TargetFolder (Resume: $Resume)"

# ==============================================================================
# BO MAY WATCHDOG NATIVE C# (ZERO-HANG HARD TIMEOUT ENGINE)
# ==============================================================================
Add-Type -TypeDefinition @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

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
"@

$Script:Watchdog = New-Object ExcelWatchdog

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
# QUAN LY TIEN TRINH EXCEL COM NGUYEN TU (ATOMIC PROCESS MANAGEMENT)
# ==============================================================================
$Script:CurrentExcelApp = $null
$Script:CurrentExcelPid = 0

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
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Start-FreshExcel {
    Write-AuditLog "[DEBUG] Trong Start-FreshExcel: Truoc Stop-CurrentExcel"
    Stop-CurrentExcel
    Write-AuditLog "[DEBUG] Trong Start-FreshExcel: Sau Stop-CurrentExcel"
    
    $maxRetries = 5
    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            $pidsBefore = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            Write-AuditLog "[DEBUG] Trong Start-FreshExcel (Lan $attempt): Truoc New-Object"
            $Script:CurrentExcelApp = New-Object -ComObject Excel.Application
            Write-AuditLog "[DEBUG] Trong Start-FreshExcel (Lan $attempt): Sau New-Object"
            
            # VACCINE HRESULT 0x800AC472: Cho 300ms de message pump va cac add-in khoi dong on dinh
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

            # VACCINE HRESULT 0x800AC472: Bao boc TUNG THUOC TINH RIENG BIET trong try-catch
            # Tuyet doi khong de bat ky thuoc tinh nao (nhu AlertBeforeOverwriting) lam loi tien trinh khoi dong!
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
            Write-AuditLog "[DEBUG] Start-FreshExcel Lan $attempt that bai: $errTxt"
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
# KHOI TAO PHIEN QUET & CHECKPOINT (SESSION CHECKPOINTING v3.6.0)
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

# Ghi checkpoint khoi tao
Save-SessionMeta "In-Progress"

Write-Host "`nDang khoi dong Excel COM doc lap (PID quan ly rieng)..." -ForegroundColor Gray
Write-AuditLog "[DEBUG] Bat dau khoi dong Excel COM..."
if (-not (Start-FreshExcel)) {
    Write-AuditLog "[DEBUG] Start-FreshExcel tra ve false!"
    Write-Host "`nNhan Enter de thoat..." -ForegroundColor Gray
    try { Read-Host | Out-Null } catch {}
    exit 1
}
Write-AuditLog "[DEBUG] Start-FreshExcel thanh cong, PID: $Script:CurrentExcelPid"
Write-Host "   -> [OK] Excel Worker v3.6.0 da san sang (PID: $Script:CurrentExcelPid)" -ForegroundColor Green

$excelExtensions = @(".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt", ".xla", ".xlam")
$virusKeywords   = @("Kangatang", "Kangaatang", "Kanga", "mypersonnel")

# Ham quet va xu ly tung tep Excel voi phong thu da tang & Watchdog 25s
function Scan-SingleExcelFile($file) {
    $Script:TotalScanned++
    $filePath = $file.FullName
    $idx = $Script:TotalScanned
    
    # Cap nhat tieu de cua so thoi gian thuc
    $Host.UI.RawUI.WindowTitle = "KangatangGuard v3.6.0 | Da quet: $Script:TotalScanned | Da diet: $Script:TotalCleaned | PID: $Script:CurrentExcelPid"
    
    # Dinh ky lam moi tien trinh Excel moi 30 tep de chong tran bo nho
    if ($Script:TotalScanned % 30 -eq 0) {
        Write-Host "  -> [RECYCLE] Dinh ky lam sach bo nho COM tai tep #$idx..." -ForegroundColor DarkCyan
        Start-FreshExcel | Out-Null
    }

    $wb = $null
    $openSuccess = $false

    # Kich hoat Watchdog 25 giay truoc khi mo tep
    $Script:Watchdog.Arm($Script:CurrentExcelPid, 25000)

    try {
        # Mo ReadOnly = $true voi 3 tham so nguyen ban on dinh tuyet doi
        $wb = $Script:CurrentExcelApp.Workbooks.Open($filePath, 0, $true)
        $openSuccess = $true
    } catch {
        $errMsg = $_.Exception.Message
        $Script:TotalErrors++

        if ($Script:Watchdog.TimedOut) {
            Write-Host "  [$idx] [TIMEOUT] Tep bi treo qua 25s (bo qua an toan): $($file.Name)" -ForegroundColor Red
            Write-AuditLog "[TIMEOUT] $filePath | Qua thoi gian 25s"
            Start-FreshExcel | Out-Null
        } else {
            Write-Host "  [$idx] Khong the mo: $($file.Name) | Loi: $errMsg" -ForegroundColor DarkYellow
            Write-AuditLog "[SKIP] Khong the mo: $filePath | Loi: $errMsg"
            # Auto-Recovery ngay lap tuc de bao ve kenh COM cho cac tep tiep theo
            Start-FreshExcel | Out-Null
        }
        return
    } finally {
        $Script:Watchdog.Disarm()
    }

    if (-not $openSuccess -or $null -eq $wb) {
        $Script:TotalErrors++
        return
    }

    $isInfected = $false
    $infectionDetails = [System.Collections.Generic.List[string]]::new()

    # Kiem tra VBProject Components (Chi kiem tra voi tep macro, khong truy van voi .xlsx de tranh loi COM)
    $ext = $file.Extension.ToLower()
    if ($ext -ne ".xlsx" -and $ext -ne ".xltx") {
        try {
            $vbProj = $wb.VBProject
            if ($vbProj) {
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
                            $lines = $cm.Lines(1, $cm.CountOfLines)
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

    # Kiem tra Sheet an
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

    # Kiem tra Hidden Named Ranges
    try {
        foreach ($nm in $wb.Names) {
            foreach ($kw in $virusKeywords) {
                if ($nm.Name -like "*$kw*" -or $nm.RefersTo -like "*$kw*") {
                    $isInfected = $true
                    $infectionDetails.Add("Named Range doc hai: " + $nm.Name)
                    break
                }
            }
        }
    } catch {}

    # Xu ly khi phat hien virus
    if ($isInfected) {
        Write-Host "  [PHAT HIEN VIRUS] [$idx]: $($file.Name)" -ForegroundColor Red
        foreach ($det in $infectionDetails) {
            Write-Host "     - $det" -ForegroundColor DarkRed
        }
        Write-AuditLog "[DETECTED] $filePath | $($infectionDetails -join '; ')"

        # Dong file Read-Only de chuan bi lam sach
        try { $wb.Close($false) } catch {}
        $wb = $null

        # Kiem tra khoa ghi truoc khi can thiep (Pre-flight Write Lock Check)
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
            Write-Host "     [SAO LUU] Da tao ban sao tai: $backupPath" -ForegroundColor Green
            Write-AuditLog "[BACKUP] $backupPath"
            $backupSuccess = $true
        } catch {
            $errTxt = $_.Exception.Message
            Write-Host "     [LOI SAO LUU] Khong the tao ban sao: $errTxt" -ForegroundColor Yellow
            Write-AuditLog "[BACKUP_ERROR] Khong the backup $filePath : $errTxt"
        }

        if ($backupSuccess) {
            $cleaned = $false
            # Kich hoat Watchdog 25 giay cho tien trinh lam sach
            $Script:Watchdog.Arm($Script:CurrentExcelPid, 25000)

            try {
                # Mo lai che do Read-Write bang 3 tham so on dinh
                $wb = $Script:CurrentExcelApp.Workbooks.Open($filePath, $false, $false)
                
                if ($null -eq $wb) {
                    Write-Host "     [LOI GHI] Khong the lay doi tuong Workbook khi mo Read-Write." -ForegroundColor Red
                    Write-AuditLog "[CLEAN_ERROR] $filePath : Workbook is null"
                    $Script:TotalErrors++
                    Start-FreshExcel | Out-Null
                    return
                }

                if ($ext -ne ".xlsx" -and $ext -ne ".xltx") {
                    try {
                        $vbProj = $wb.VBProject
                        if ($vbProj) {
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
                                    $lines = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
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

                for ($nIdx = $wb.Names.Count; $nIdx -ge 1; $nIdx--) {
                    $nm = $wb.Names.Item($nIdx)
                    foreach ($kw in $virusKeywords) {
                        if ($nm.Name -like "*$kw*" -or $nm.RefersTo -like "*$kw*") {
                            $nm.Delete()
                            $cleaned = $true
                            break
                        }
                    }
                }

                # Tat cac hop thoai xac nhan tuong thich / thong tin ca nhan khi luu qua mang
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
            } catch {
                $errTxt = $_.Exception.Message
                if ($Script:Watchdog.TimedOut) {
                    Write-Host "     [TIMEOUT LAM SACH] Qua trinh lam sach bi treo qua 25s: $($file.Name)" -ForegroundColor Red
                    Write-AuditLog "[CLEAN_TIMEOUT] $filePath"
                } else {
                    Write-Host "     [LOI] Loi trong qua trinh lam sach: $errTxt" -ForegroundColor Red
                    Write-AuditLog "[CLEAN_ERROR] $filePath : $errTxt"
                }
                $Script:TotalErrors++
                # Auto-Recovery ngay trong khoi Catch cua tien trinh lam sach
                Start-FreshExcel | Out-Null
            } finally {
                $Script:Watchdog.Disarm()
                if ($null -ne $wb) {
                    try { $wb.Close($false) } catch {}
                    $wb = $null
                }
            }
        }
    } else {
        Write-Host "  [OK] [$idx] An toan: $($file.Name)" -ForegroundColor Gray
        $Script:TotalSafe++
    }

    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
        $wb = $null
    }
}

# Ham quet luong de quy: Quet ngay tuc thi tung thu muc
function Scan-FolderStream([string]$currentDir) {
    if ($currentDir -like "*_Backup_Kangatang*") { return }

    $Script:TotalFolders++
    Write-Host "`n----------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Folder [$Script:TotalFolders]: $currentDir" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkGray
    
    Write-Progress -Activity "KangatangGuard v3.6.0 - Dang quet luong chong treo" -Status "Thu muc #$Script:TotalFolders: $currentDir" -CurrentOperation "Da quet: $Script:TotalScanned tep | Da diet: $Script:TotalCleaned"

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
                
                # Ghi ngay vao checkpoint log
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

    # Dinh ky luu checkpoint sau moi thu muc
    Save-SessionMeta "In-Progress"

    # 2. Lay danh sach thu muc con va lan luot quet tiep
    try {
        $subDirs = Get-ChildItem -Path $currentDir -Directory -ErrorAction SilentlyContinue
        foreach ($sub in $subDirs) {
            if ($sub.Name -ne "_Backup_Kangatang") {
                Scan-FolderStream $sub.FullName
            }
        }
    } catch {
        $errTxt = $_.Exception.Message
        Write-Host "  [LOI DUYET THU MUC CON] $errTxt" -ForegroundColor Red
    }
}

# 3. Kich hoat quet luong ngay lap tuc!
Write-Host "`nBAT DAU QUET LUONG TRUC TIEP CHONG TREO (v3.6.0)..." -ForegroundColor Cyan
Scan-FolderStream $TargetFolder

Write-Progress -Activity "KangatangGuard v3.6.0" -Completed

# 4. Giai phong va dong tien trinh Excel COM
Stop-CurrentExcel

# 5. Cap nhat checkpoint hoan tat 100%
Save-SessionMeta "Completed"

# 6. Bao cao tong ket
Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   BAO CAO TONG KET QUET LUONG CHONG TREO (v3.6.0)                   " -ForegroundColor Green
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

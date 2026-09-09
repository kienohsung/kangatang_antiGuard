# ==============================================================================
# Kangatang_FolderScanner.ps1
# Tien trinh Quet Luong Truc Tiep & Tu dong Phuc hoi (Auto-Recovery Worker)
# Phien ban: v3.5.1 (Enterprise Robust Batch Scanner)
# Dac diem:
#   - Quet luong ngay lap tuc lan luot tung thu muc, khong co do tre cho gom tep.
#   - Quan ly vong doi tien trinh Excel rieng biet (PID Tracking).
#   - Tu dong hoi sinh Excel COM ngay lap tuc khi gap file loi/hong (Auto-Recovery).
#   - Dinh ky lam moi tien trinh Excel moi 30 tep de chong ro ri bo nho.
#   - Che do mo phong thu Read-Only tranh khoa tep va khong bi chan boi hop thoai.
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = ""
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$Host.UI.RawUI.WindowTitle = "KangatangGuard v3.5.1 - Trinh quet luong truc tiep & Tu dong phuc hoi"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANGGUARD v3.5.1 - TRINH QUET LUONG TU DONG PHUC HOI (ROBUST) " -ForegroundColor Cyan
Write-Host "   Kien truc Quan ly Tien trinh Doc lap & Chuyen biet Chuyen sau      " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Xu ly va chuan hoa tham so duong dan thu muc
if ($TargetFolder) {
    # Xu ly truong hop CLI thoat dau ngoac kep lam dinh dau kep o cuoi: "D:" -> D:"
    $TargetFolder = $TargetFolder.Trim('"').Trim("'").TrimEnd('\')
    if ($TargetFolder -match '^[a-zA-Z]:$') {
        $TargetFolder = $TargetFolder + "\"
    }
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

if ($TargetFolder -match '^[a-zA-Z]:$') {
    $TargetFolder = $TargetFolder + "\"
} else {
    $TargetFolder = $TargetFolder.TrimEnd('\')
}

Write-Host "`n[THU MUC BAT DAU] : $TargetFolder" -ForegroundColor Yellow

# Thiet lap thu muc nhat ky (Audit Log)
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logFile = Join-Path $logDir ("scan_log_" + (Get-Date -Format "yyyyMMdd") + ".txt")

function Write-AuditLog([string]$msg) {
    $timeStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timeStr | $msg"
    try {
        Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

Write-AuditLog "[STREAM_SCAN_START] Bat dau quet luong tai: $TargetFolder"

# ==============================================================================
# QUAN LY TIEN TRINH EXCEL COM NGUYEN TU (ATOMIC PROCESS MANAGEMENT)
# ==============================================================================
$Script:CurrentExcelApp = $null
$Script:CurrentExcelPid = 0

function Stop-CurrentExcel {
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
    Stop-CurrentExcel
    try {
        $pidsBefore = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $Script:CurrentExcelApp = New-Object -ComObject Excel.Application
        $pidsAfter = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        
        $newPid = $pidsAfter | Where-Object { $pidsBefore -notcontains $_ } | Select-Object -First 1
        if ($newPid) { $Script:CurrentExcelPid = $newPid }

        $Script:CurrentExcelApp.Visible = $false
        $Script:CurrentExcelApp.DisplayAlerts = $false
        $Script:CurrentExcelApp.ScreenUpdating = $false
        $Script:CurrentExcelApp.EnableEvents = $false
        $Script:CurrentExcelApp.AskToUpdateLinks = $false
        $Script:CurrentExcelApp.AlertBeforeOverwriting = $false
        try { $Script:CurrentExcelApp.FeatureInstall = 0 } catch {}
        try { $Script:CurrentExcelApp.AutomationSecurity = 3 } catch {}
        return $true
    } catch {
        $errTxt = $_.Exception.Message
        Write-Host "[LOI] Khong the khoi dong Excel COM: $errTxt" -ForegroundColor Red
        Write-AuditLog "[STREAM_SCAN_ERROR] Khong the khoi dong Excel COM: $errTxt"
        return $false
    }
}

Write-Host "`nDang khoi dong Excel COM doc lap (PID quan ly rieng)..." -ForegroundColor Gray
if (-not (Start-FreshExcel)) {
    Write-Host "`nNhan Enter de thoat..." -ForegroundColor Gray
    Read-Host | Out-Null
    exit 1
}
Write-Host "   -> [OK] Excel Worker da san sang (PID: $Script:CurrentExcelPid)" -ForegroundColor Green

# Bien dem thoi gian thuc
$Script:TotalScanned = 0
$Script:TotalCleaned = 0
$Script:TotalSafe    = 0
$Script:TotalErrors  = 0
$Script:TotalFolders = 0

$excelExtensions = @(".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt", ".xla", ".xlam")
$virusKeywords   = @("Kangatang", "Kangaatang", "Kanga", "mypersonnel")

# Ham quet va xu ly tung tep Excel voi phong thu da tang
function Scan-SingleExcelFile($file) {
    $Script:TotalScanned++
    $filePath = $file.FullName
    $idx = $Script:TotalScanned
    
    # Cap nhat tieu de cua so thoi gian thuc
    $Host.UI.RawUI.WindowTitle = "KangatangGuard v3.5.1 | Da quet: $Script:TotalScanned | Da diet: $Script:TotalCleaned | PID: $Script:CurrentExcelPid"
    
    # Dinh ky lam moi tien trinh Excel moi 30 tep de chong tran bo nho
    if ($Script:TotalScanned % 30 -eq 0) {
        Write-Host "  -> [RECYCLE] Dinh ky lam sach bo nho COM tai tep #$idx..." -ForegroundColor DarkCyan
        Start-FreshExcel | Out-Null
    }

    $wb = $null
    $openSuccess = $false

    try {
        # Mo ReadOnly = $true voi 3 tham so chuan: FilePath, UpdateLinks=0, ReadOnly=$true
        $wb = $Script:CurrentExcelApp.Workbooks.Open($filePath, 0, $true)
        $openSuccess = $true
    } catch {
        $errMsg = $_.Exception.Message
        Write-Host "  [$idx] Khong the mo: $($file.Name) | Loi: $errMsg" -ForegroundColor DarkYellow
        Write-AuditLog "[SKIP] Khong the mo: $filePath | Loi: $errMsg"
        $Script:TotalErrors++

        # Kiem tra neu Excel bi ngat ket noi thi lap tuc hoi sinh
        $isAlive = $false
        try {
            if ($null -ne $Script:CurrentExcelApp -and $Script:CurrentExcelApp.Version) {
                $isAlive = $true
            }
        } catch {}

        if (-not $isAlive) {
            Write-Host "  -> [AUTO-RECOVERY] Phat hien Excel COM bi ngat, dang tu dong khoi phuc..." -ForegroundColor Magenta
            Write-AuditLog "[AUTO_RECOVERY] Tu dong khoi phuc Excel COM tai: $filePath"
            Start-FreshExcel | Out-Null
        }
        return
    }

    if (-not $openSuccess -or $null -eq $wb) {
        $Script:TotalErrors++
        return
    }

    $isInfected = $false
    $infectionDetails = [System.Collections.Generic.List[string]]::new()

    # Kiem tra VBProject Components
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
            try {
                # Mo lai che do Read-Write de lam sach
                $wb = $Script:CurrentExcelApp.Workbooks.Open($filePath, $false, $false, [System.Type]::Missing, [System.Type]::Missing, [System.Type]::Missing, $true)
                
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
                Write-Host "     [LOI] Loi trong qua trinh lam sach: $errTxt" -ForegroundColor Red
                Write-AuditLog "[CLEAN_ERROR] $filePath : $errTxt"
                $Script:TotalErrors++
            }
        }
    } else {
        Write-Host "  [OK] [$idx] An toan: $($file.Name)" -ForegroundColor Gray
        $Script:TotalSafe++
    }

    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
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
    
    Write-Progress -Activity "KangatangGuard v3.5.1 - Dang quet luong" -Status "Thu muc #$Script:TotalFolders: $currentDir" -CurrentOperation "Da quet: $Script:TotalScanned tep | Da diet: $Script:TotalCleaned"

    # 1. Quet ngay lap tuc tat ca tep Excel co trong thu muc nay
    try {
        $files = Get-ChildItem -Path $currentDir -File -ErrorAction SilentlyContinue
        foreach ($f in $files) {
            if ($excelExtensions -contains $f.Extension.ToLower() -and $f.Name -ne "KangatangGuard.xlam") {
                Scan-SingleExcelFile $f
            }
        }
    } catch {
        $errTxt = $_.Exception.Message
        Write-Host "  [LOI DUYET TEP] $errTxt" -ForegroundColor Red
    }

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
Write-Host "`nBAT DAU QUET LUONG TRUC TIEP..." -ForegroundColor Cyan
Scan-FolderStream $TargetFolder

Write-Progress -Activity "KangatangGuard v3.5.1" -Completed

# 4. Giai phong va dong tien trinh Excel COM
Stop-CurrentExcel

# 5. Bao cao tong ket
Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   BAO CAO TONG KET QUET LUONG TRUC TIEP (v3.5.1)                    " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green
Write-Host "   Thu muc bat dau                 : $TargetFolder" -ForegroundColor White
Write-Host "   Tong so thu muc da duyet qua    : $Script:TotalFolders" -ForegroundColor Yellow
Write-Host "   Tong so tep Excel da kiem tra   : $Script:TotalScanned" -ForegroundColor Cyan
Write-Host "   So tep an toan                  : $Script:TotalSafe" -ForegroundColor Green
Write-Host "   So tep phat hien & da tieu diet : $Script:TotalCleaned" -ForegroundColor $(if ($Script:TotalCleaned -gt 0) { "Red" } else { "Green" })
Write-Host "   So tep loi khong mo duoc        : $Script:TotalErrors" -ForegroundColor $(if ($Script:TotalErrors -gt 0) { "Yellow" } else { "Gray" })
Write-Host "   Nhat ky chi tiet                : $logFile" -ForegroundColor DarkGray
Write-Host "======================================================================" -ForegroundColor Green

Write-AuditLog "[STREAM_SCAN_END] $TargetFolder - ThuMuc: $Script:TotalFolders, Tep: $Script:TotalScanned, Diet: $Script:TotalCleaned, AnToan: $Script:TotalSafe, Loi: $Script:TotalErrors"

Write-Host "`nToan bo tien trinh quet luong da hoan tat ma khong lam giam hieu nang Excel." -ForegroundColor Cyan
Write-Host "Nhan Enter de hoan tat va dong cua so..." -ForegroundColor Gray
Read-Host | Out-Null

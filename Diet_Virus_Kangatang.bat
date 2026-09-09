@echo off
setlocal
chcp 65001 >nul
title HỆ THỐNG DIỆT VIRUS EXCEL KANGATANG - EXCEL CLEANER v3.6.0

:: 1. Kiểm tra quyền Administrator và tự động kích hoạt nếu chưa có
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Đang yêu cầu quyền Administrator để xử lý tất cả tài khoản...
    powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList '/c \"\"%~f0\"\"' -Verb RunAs"
    exit /b
)

:: 2. Đọc và thực thi toàn bộ logic PowerShell từ chính tệp này
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$raw = [System.IO.File]::ReadAllText('%~f0', [System.Text.Encoding]::UTF8); $marker = '# === POWERSHELL_' + 'CODE_START ==='; $idx = $raw.LastIndexOf($marker); if ($idx -ge 0) { $code = $raw.Substring($idx + $marker.Length); Invoke-Expression $code }"

echo.
echo ======================================================================
echo Nhấn phím bất kỳ để thoát chương trình...
pause >nul
exit /b

# === POWERSHELL_CODE_START ===
# Thiết lập chuẩn mã hóa UTF-8 để hiển thị đầy đủ tiếng Việt có dấu
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   HỆ THỐNG DIỆT VIRUS MACRO KANGATANG - EXCEL CLEANER (v3.6.0)       " -ForegroundColor Cyan
Write-Host "   Phiên bản Fast Resume & Quét Luồng Chống Treo v3.6.0              " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# --- BIẾN TOÀN CỤC ---
$Script:LogEntries = [System.Collections.Generic.List[string]]::new()
$Script:AccessVBOMBackup = @{}

# --- HÀM GIẢI PHÓNG BỘ NHỚ COM ---
function Release-ComResource {
    param ([System.Object]$ComObject)
    if ($null -ne $ComObject) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        } catch {}
    }
}

function Invoke-GarbageClean {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# --- HÀM GHI LOG KIỂM TOÁN (v3.1.0) ---
function Write-ScanLog {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "$timestamp | $Message"
    $Script:LogEntries.Add($logLine)
    # Đồng thời ghi console debug
    Write-Host "   [LOG] $Message" -ForegroundColor DarkGray
}

function Export-ScanReport {
    $logDir = Join-Path $env:APPDATA "KangatangGuard"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $timeStr = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFileName = "scan_report_" + $timeStr + ".log"
    $logFilePath = Join-Path $logDir $logFileName
    
    $Script:LogEntries | Out-File -FilePath $logFilePath -Encoding utf8 -Force
    Write-Host "`n   📄 [LOG] Báo cáo kiểm toán đã xuất: $logFilePath" -ForegroundColor Cyan
}

# --- HÀM BẬT/HOÀN NGUYÊN AccessVBOM REGISTRY (v3.1.0) ---
function Enable-AccessVBOM {
    Write-Host "`n[0/4] 🔑 Đang bật quyền truy cập VBA Project Object Model (AccessVBOM)..." -ForegroundColor Yellow
    $officeVersions = @("14.0", "15.0", "16.0")
    
    foreach ($ver in $officeVersions) {
        $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
        if (Test-Path $regPath) {
            # Lưu giá trị cũ để hoàn nguyên sau
            try {
                $currentVal = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
                if ($null -eq $currentVal) { $currentVal = -1 }
                $Script:AccessVBOMBackup[$regPath] = $currentVal
            } catch {
                $Script:AccessVBOMBackup[$regPath] = -1
            }
            
            try {
                Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction Stop
                Write-Host "   -> ✅ Đã bật AccessVBOM cho Office $ver" -ForegroundColor Green
                Write-ScanLog "AccessVBOM ON: Office $ver ($regPath)"
            } catch {
                Write-Host "   -> ⚠️ Không thể ghi Registry Office $ver : $($_.Exception.Message)" -ForegroundColor DarkYellow
                Write-ScanLog "AccessVBOM ERROR: Office $ver - $($_.Exception.Message)"
            }
        }
    }
}

function Restore-AccessVBOM {
    Write-Host "`n🔒 Đang hoàn nguyên Registry AccessVBOM về trạng thái ban đầu..." -ForegroundColor Yellow
    foreach ($regPath in $Script:AccessVBOMBackup.Keys) {
        $originalVal = $Script:AccessVBOMBackup[$regPath]
        try {
            if ($originalVal -eq -1) {
                # Giá trị không tồn tại trước đó -> xóa property
                Remove-ItemProperty -Path $regPath -Name "AccessVBOM" -Force -ErrorAction SilentlyContinue
            } else {
                Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value $originalVal -Type DWord -Force -ErrorAction Stop
            }
            Write-Host "   -> ✅ Đã hoàn nguyên: $regPath" -ForegroundColor Green
            Write-ScanLog "AccessVBOM RESTORED: $regPath -> $originalVal"
        } catch {
            Write-Host "   -> ⚠️ Lỗi hoàn nguyên: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
}

# --- HÀM BACKUP FILE TRƯỚC KHI SỬA (v3.1.0) ---
function Backup-ExcelFile {
    param ([string]$FilePath)
    
    $parentDir = Split-Path -Parent $FilePath
    $backupDir = Join-Path $parentDir "_Backup_Kangatang"
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }
    
    $fileName = Split-Path -Leaf $FilePath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $ext = [System.IO.Path]::GetExtension($fileName)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupName = $baseName + "_backup_" + $timestamp + $ext
    $backupPath = Join-Path $backupDir $backupName
    
    try {
        Copy-Item -Path $FilePath -Destination $backupPath -Force -ErrorAction Stop
        Write-Host "   📦 [BACKUP] Đã sao lưu: $backupName" -ForegroundColor Cyan
        Write-ScanLog "BACKUP: $FilePath -> $backupPath"
        return $true
    } catch {
        Write-Host "   ⚠️ [BACKUP LỖI] Không thể sao lưu: $($_.Exception.Message)" -ForegroundColor DarkYellow
        Write-ScanLog "BACKUP_ERROR: $FilePath - $($_.Exception.Message)"
        return $false
    }
}

# --- HÀM TÌM KIẾM TẤT CẢ THƯ MỤC CẤU HÌNH EXCEL ---
function Get-AllExcelPaths {
    $foundPaths = [System.Collections.Generic.List[string]]::new()
    
    # 1. Thư mục của tài khoản hiện tại
    $currentUserAppData = Join-Path $env:APPDATA "Microsoft\Excel"
    if (Test-Path $currentUserAppData) {
        $foundPaths.Add($currentUserAppData)
    }
    
    # 2. Quét qua tất cả tài khoản trong C:\Users (OHSUNG VINA, Admin, mrKienIT,...)
    if (Test-Path "C:\Users") {
        $userProfiles = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue
        foreach ($profile in $userProfiles) {
            $userExcelPath = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Excel"
            if (Test-Path $userExcelPath) {
                if (-not $foundPaths.Contains($userExcelPath)) {
                    $foundPaths.Add($userExcelPath)
                }
            }
        }
    }
    
    return $foundPaths
}

# --- HÀM TÌM KIẾM TẤT CẢ THƯ MỤC XLSTART ---
function Get-AllXLStartPaths {
    $xlStartDirs = [System.Collections.Generic.List[string]]::new()
    
    # 1. Lấy từ thư mục Excel của mọi người dùng
    $excelDirs = Get-AllExcelPaths
    foreach ($dir in $excelDirs) {
        $userXLStart = Join-Path $dir "XLSTART"
        if (Test-Path $userXLStart) {
            if (-not $xlStartDirs.Contains($userXLStart)) {
                $xlStartDirs.Add($userXLStart)
            }
        }
    }
    
    # 2. Quét các thư mục XLSTART cài đặt hệ thống của Office (32-bit va 64-bit)
    $systemOfficePaths = @(
        "C:\Program Files\Microsoft Office\root\Office*",
        "C:\Program Files (x86)\Microsoft Office\root\Office*",
        "C:\Program Files\Microsoft Office\Office*",
        "C:\Program Files (x86)\Microsoft Office\Office*"
    )
    
    foreach ($pattern in $systemOfficePaths) {
        $officeDirs = Get-ChildItem -Path $pattern -Directory -ErrorAction SilentlyContinue
        foreach ($officeDir in $officeDirs) {
            $sysXLStart = Join-Path $officeDir.FullName "XLSTART"
            if (Test-Path $sysXLStart) {
                if (-not $xlStartDirs.Contains($sysXLStart)) {
                    $xlStartDirs.Add($sysXLStart)
                }
            }
        }
    }
    
    return $xlStartDirs
}

# ==============================================================================
# BƯỚC 0: BẬT AccessVBOM (v3.1.0 - MỚI)
# ==============================================================================
Enable-AccessVBOM

# ==============================================================================
# BƯỚC 1: ĐÓNG TOÀN BỘ TIẾN TRÌNH EXCEL
# ==============================================================================
Write-Host "`n[1/4] 🛑 Đang đóng toàn bộ tiến trình Microsoft Excel để giải phóng khóa tệp..." -ForegroundColor Yellow
$excelProcesses = Get-Process -Name excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    $excelProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Host "   -> ✅ Đã dừng $($excelProcesses.Count) tiến trình Excel đang chạy." -ForegroundColor Green
    Write-ScanLog "Đóng $($excelProcesses.Count) tiến trình Excel."
} else {
    Write-Host "   -> ✅ Không có tiến trình Excel nào đang chạy ngầm." -ForegroundColor Green
}

# ==============================================================================
# BƯỚC 2: DỌN DẸP THƯ MỤC APPDATA EXCEL CỦA TOÀN BỘ NGƯỜI DÙNG
# ==============================================================================
Write-Host "`n[2/4] 🛡️ Đang rà soát và dọn dẹp thư mục cấu hình AppData Excel..." -ForegroundColor Yellow
$allExcelDirs = Get-AllExcelPaths

if ($allExcelDirs.Count -eq 0) {
    Write-Host "   -> ⚠️ Không tìm thấy thư mục cấu hình Microsoft Excel nào." -ForegroundColor DarkYellow
} else {
    Write-Host "   -> Phát hiện $($allExcelDirs.Count) thư mục cấu hình Excel trong hệ thống." -ForegroundColor Cyan
    foreach ($excelDir in $allExcelDirs) {
        Write-Host "`n   📁 Thư mục đang xử lý: $excelDir" -ForegroundColor Gray
        $items = Get-ChildItem -Path $excelDir -Force -ErrorAction SilentlyContinue
        
        foreach ($item in $items) {
            # BẢO VỆ: Giữ lại tệp cấu hình hợp lệ Excel*.xlb, thư mục XLSTART, Templates
            if ($item.Name -like "Excel*.xlb" -or $item.Name -eq "XLSTART" -or $item.Name -eq "Templates") {
                Write-Host "      [Giữ lại] -> $($item.Name)" -ForegroundColor DarkGray
                continue
            }
            
            # BẢO VỆ: Giữ lại KangatangGuard.xlam (Add-in bảo vệ)
            if ($item.Name -eq "KangatangGuard.xlam") {
                Write-Host "      [Giữ lại] -> $($item.Name) (Add-in bảo vệ)" -ForegroundColor DarkGray
                continue
            }
            
            # Xóa các tệp/thư mục rác do virus sinh ra (mypersonnel*, file tạm *.tmp, v.v.)
            try {
                Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
                Write-Host "      [Đã xóa rác]  -> $($item.Name)" -ForegroundColor Red
                Write-ScanLog "XÓA RÁC AppData: $($item.FullName)"
            } catch {
                Write-Host "      [Cảnh báo] Không thể xóa $($item.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
    }
}

# ==============================================================================
# BƯỚC 3: TIÊU DIỆT TỆP MẦM BỆNH TRONG TẤT CẢ THƯ MỤC XLSTART
# ==============================================================================
Write-Host "`n[3/4] 🛡️ Đang truy quét ổ dịch trong TẤT CẢ thư mục XLSTART..." -ForegroundColor Yellow
$allXLStartDirs = Get-AllXLStartPaths

if ($allXLStartDirs.Count -eq 0) {
    Write-Host "   -> ⚠️ Không tìm thấy thư mục XLSTART nào trên máy." -ForegroundColor DarkYellow
} else {
    Write-Host "   -> Phát hiện $($allXLStartDirs.Count) thư mục XLSTART cần làm sạch." -ForegroundColor Cyan
    
    $virusPatterns = @("mypersonnel*", "*kangatang*", "*kangaatang*", "*kanga*.xls*", "personal.xls")
    
    foreach ($xlDir in $allXLStartDirs) {
        Write-Host "`n   📂 Kiểm tra XLSTART: $xlDir" -ForegroundColor Gray
        
        $infectedFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        foreach ($pattern in $virusPatterns) {
            $matched = Get-ChildItem -Path $xlDir -Filter $pattern -File -Force -ErrorAction SilentlyContinue
            if ($matched) {
                foreach ($m in $matched) {
                    # BẢO VỆ: Không xóa KangatangGuard.xlam
                    if ($m.Name -eq "KangatangGuard.xlam") { continue }
                    if (-not $infectedFiles.Contains($m)) {
                        $infectedFiles.Add($m)
                    }
                }
            }
        }
        
        if ($infectedFiles.Count -eq 0) {
            Write-Host "      -> ✅ Không phát hiện file mầm bệnh mypersonnel/kangatang." -ForegroundColor Green
        } else {
            foreach ($vFile in $infectedFiles) {
                try {
                    Remove-Item -Path $vFile.FullName -Force -ErrorAction Stop
                    Write-Host "      🛑 [ĐÃ TIÊU DIỆT] -> Xóa thành công file mầm bệnh: $($vFile.Name)" -ForegroundColor Red
                    Write-ScanLog "TIÊU DIỆT XLSTART: $($vFile.FullName)"
                } catch {
                    Write-Host "      ⚠️ [Lỗi xóa file] -> $($vFile.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
            }
        }
    }
}

# ==============================================================================
# BƯỚC 3.5: RÀ SOÁT REGISTRY OPTIONS, ALTSTARTUP VÀ ADDINS (v3.2.0 - OFFICE 365)
# ==============================================================================
Write-Host "`n[3.5/4] 🛡️ Đang rà soát Registry Options, Add-ins & Khởi tạo XLSTART chuẩn..." -ForegroundColor Yellow

# 1. Đảm bảo thư mục XLSTART của tài khoản hiện tại tồn tại (phòng chống lỗi thiếu thư mục trên Office 365)
$currentUserXLStart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (-not (Test-Path $currentUserXLStart)) {
    try {
        New-Item -ItemType Directory -Path $currentUserXLStart -Force | Out-Null
        Write-Host "   -> 🛡️ [KHỞI TẠO] Đã tạo thư mục XLSTART chuẩn cho Office 365: $currentUserXLStart" -ForegroundColor Green
        Write-ScanLog "TẠO XLSTART: $currentUserXLStart"
    } catch {}
}

# 2. Rà soát Registry Options (OPEN*, AltStartupPath)
foreach ($ver in @("16.0", "15.0", "14.0")) {
    $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path $optKey) {
        $props = (Get-ItemProperty -Path $optKey -ErrorAction SilentlyContinue).psobject.Properties
        foreach ($p in $props) {
            if ($p.Name -match "^OPEN\d*$") {
                $val = [string]$p.Value
                if ($val -match "mypersonnel|kangatang|kanga|personal\.xls") {
                    try {
                        Remove-ItemProperty -Path $optKey -Name $p.Name -Force -ErrorAction Stop
                        Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Gỡ bỏ khóa Registry tự nạp mã độc: $($p.Name) = $val" -ForegroundColor Red
                        Write-ScanLog "XÓA REGISTRY MALWARE: $optKey\$($p.Name) = $val"
                    } catch {
                        Write-Host "   ⚠️ [Lỗi gỡ Registry] $($_.Exception.Message)" -ForegroundColor DarkYellow
                    }
                }
            }
        }
        
        # Kiểm tra AltStartupPath
        $altStartup = (Get-ItemProperty -Path $optKey -Name "AltStartupPath" -ErrorAction SilentlyContinue).AltStartupPath
        if ($altStartup -and (Test-Path $altStartup)) {
            Write-Host "   📂 Phát hiện AltStartupPath: $altStartup (Đang kiểm tra...)" -ForegroundColor Gray
            $altFiles = Get-ChildItem -Path $altStartup -Filter "*.*" -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -match "mypersonnel|kangatang|kanga|personal\.xls" }
            foreach ($af in $altFiles) {
                try {
                    Remove-Item -Path $af.FullName -Force -ErrorAction Stop
                    Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Xóa file mầm bệnh trong AltStartupPath: $($af.Name)" -ForegroundColor Red
                    Write-ScanLog "XÓA ALTSTARTUP FILE: $($af.FullName)"
                } catch {}
            }
        }
    }
}

# 3. Rà soát thư mục Microsoft\AddIns
$userAddIns = Join-Path $env:APPDATA "Microsoft\AddIns"
if (Test-Path $userAddIns) {
    $badAddins = Get-ChildItem -Path $userAddIns -File -ErrorAction SilentlyContinue |
                 Where-Object { ($_.Name -match "mypersonnel|kangatang|kanga|personal\.xls") -and ($_.Name -ne "KangatangGuard.xlam") }
    foreach ($ba in $badAddins) {
        try {
            Remove-Item -Path $ba.FullName -Force -ErrorAction Stop
            Write-Host "   🛑 [ĐÃ TIÊU DIỆT] Xóa file mầm bệnh trong AddIns: $($ba.Name)" -ForegroundColor Red
            Write-ScanLog "XÓA ADDINS MALWARE: $($ba.FullName)"
        } catch {}
    }
}

# ==============================================================================
# BƯỚC 4: HÀM QUÉT VÀ LÀM SẠCH FILE EXCEL TRONG THƯ MỤC TÙY CHỌN (v3.2.0)
# ==============================================================================
function Clean-ExcelVirusInFolder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TargetFolder
    )
    
    # Chuẩn hóa đường dẫn: loại bỏ dấu ngoặc kép hoặc ngoặc đơn
    $TargetFolder = $TargetFolder.Replace([char]34, [string]::Empty).Replace([char]39, [string]::Empty).Trim()
    
    if (-not (Test-Path $TargetFolder)) {
        Write-Host "[Lỗi] Đường dẫn thư mục quét không tồn tại: $TargetFolder" -ForegroundColor Red
        return
    }
    
    Write-Host "`n[4/4] 🚀 Khởi động trình quét VBA Engine trong thư mục: $TargetFolder" -ForegroundColor Yellow
    Write-ScanLog "BẮT ĐẦU QUÉT THƯ MỤC: $TargetFolder"
    
    # 1. Tìm tất cả các file Excel có thể chứa macro / virus
    Write-Host "Đang tìm kiếm các tệp Excel..." -ForegroundColor Gray
    $excelExtensions = @(".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt", ".xla", ".xlam")
    $excelFiles = Get-ChildItem -Path $TargetFolder -Recurse -File -ErrorAction SilentlyContinue | 
                  Where-Object { ($excelExtensions -contains $_.Extension.ToLower()) -and ($_.Name -ne "KangatangGuard.xlam") }
    
    if (-not $excelFiles -or $excelFiles.Count -eq 0) {
        Write-Host "-> Không tìm thấy tệp Excel (.xls, .xlsm, .xlsb,...) nào trong thư mục được chọn." -ForegroundColor Green
        return
    }
    
    Write-Host "-> Đã tìm thấy $($excelFiles.Count) tệp Excel cần quét cấu trúc." -ForegroundColor Cyan
    
    # 2. Khởi tạo Excel COM Object ẩn
    $excelApp = $null
    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $excelApp.ScreenUpdating = $false
        $excelApp.EnableEvents = $false
        try { $excelApp.AutomationSecurity = 3 } catch {}
    } catch {
        Write-Host "[Lỗi] Không thể khởi động Excel COM Application. Vui lòng đảm bảo MS Excel đã cài đặt." -ForegroundColor Red
        return
    }
    
    $cleanedCount = 0
    $errorCount = 0
    $safeCount = 0
    
    try {
        foreach ($file in $excelFiles) {
            $filePath = $file.FullName
            Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "🔍 [Kiểm tra] $filePath" -ForegroundColor Cyan
            
            $workbook = $null
            $virusCleanedInFile = $false
            
            try {
                $workbook = $excelApp.Workbooks.Open($filePath, 0, $false, 5, "", "", $true)
                
                $project = $null
                try {
                    $project = $workbook.VBProject
                } catch {
                    Write-Host "   ⚠️ [Bỏ qua] Không truy cập được VBProject." -ForegroundColor DarkYellow
                    Write-ScanLog "SKIP (no VBProject): $filePath"
                    $errorCount++
                    if ($workbook) { $workbook.Close($false) }
                    continue
                }
                
                if ($project -ne $null) {
                    $components = $project.VBComponents
                    
                    # 1. Duyệt và xóa Module có tên chứa Kangatang / Kanga / mypersonnel
                    for ($i = $components.Count; $i -ge 1; $i--) {
                        $comp = $components.Item($i)
                        $compName = $comp.Name
                        
                        if ($compName -match "Kangatang|Kangaatang|Kanga|mypersonnel") {
                            Write-Host "   🛑 [PHÁT HIỆN] Tìm thấy Module mã độc: $compName" -ForegroundColor Red
                            Write-ScanLog "PHÁT HIỆN Module: $compName trong $filePath"
                            try {
                                $components.Remove($comp)
                                $virusCleanedInFile = $true
                            } catch {
                                Write-Host "      ⚠️ [Lỗi xóa Module] $($_.Exception.Message)" -ForegroundColor DarkYellow
                            }
                        } else {
                            # Kiểm tra mã bên trong ThisWorkbook hoặc Sheet code module
                            try {
                                $codeMod = $comp.CodeModule
                                if ($codeMod -and $codeMod.CountOfLines -gt 0) {
                                    $content = $codeMod.Lines(1, $codeMod.CountOfLines)
                                    if ($content -match "Kangatang|Kangaatang|mypersonnel") {
                                        Write-Host "   🛑 [PHÁT HIỆN] Tìm thấy mã độc lây nhiễm trong: $compName" -ForegroundColor Red
                                        Write-ScanLog "PHÁT HIỆN Code lây nhiễm: $compName trong $filePath"
                                        $codeMod.DeleteLines(1, $codeMod.CountOfLines)
                                        $virusCleanedInFile = $true
                                    }
                                }
                            } catch {}
                        }
                        Release-ComResource $comp
                    }
                    Release-ComResource $components
                }
                
                # 2. Duyệt và xóa các Sheet rác / ẩn chứa từ khóa Kangatang
                for ($i = $workbook.Sheets.Count; $i -ge 1; $i--) {
                    $sheet = $workbook.Sheets.Item($i)
                    $sheetName = $sheet.Name
                    
                    if ($sheetName -match "Kangatang|Kangaatang|Kanga|mypersonnel") {
                        Write-Host "   🛑 [PHÁT HIỆN] Tìm thấy Sheet ẩn mã độc: $sheetName" -ForegroundColor Red
                        Write-ScanLog "PHÁT HIỆN Sheet: $sheetName trong $filePath"
                        try {
                            $sheet.Visible = -1
                            $sheet.Delete()
                            $virusCleanedInFile = $true
                        } catch {
                            Write-Host "      ⚠️ [Lỗi xóa Sheet] $($_.Exception.Message)" -ForegroundColor DarkYellow
                        }
                    }
                    Release-ComResource $sheet
                }
                
                # 3. (v3.1.0) Duyệt và xóa Hidden Named Ranges do virus tạo
                for ($i = $workbook.Names.Count; $i -ge 1; $i--) {
                    $nm = $workbook.Names.Item($i)
                    $nmName = $nm.Name
                    $shouldDelete = $false
                    
                    # Kiểm tra tên Name chứa từ khóa virus
                    if ($nmName -match "Kangatang|Kangaatang|Kanga|mypersonnel") {
                        $shouldDelete = $true
                    }
                    
                    # Kiểm tra cả tham chiếu RefersTo của Name ẩn
                    if (-not $shouldDelete) {
                        try {
                            if (-not $nm.Visible) {
                                $refStr = $nm.RefersTo
                                if ($refStr -match "Kangatang|Kangaatang|Kanga|mypersonnel") {
                                    $shouldDelete = $true
                                }
                            }
                        } catch {}
                    }
                    
                    if ($shouldDelete) {
                        Write-Host "   🛑 [PHÁT HIỆN] Tìm thấy Named Range mã độc: $nmName" -ForegroundColor Red
                        Write-ScanLog "PHÁT HIỆN Named Range: $nmName trong $filePath"
                        try {
                            $nm.Delete()
                            $virusCleanedInFile = $true
                        } catch {
                            Write-Host "      ⚠️ [Lỗi xóa Named Range] $($_.Exception.Message)" -ForegroundColor DarkYellow
                        }
                    }
                }
                
                # Backup và Lưu file nếu phát hiện và đã làm sạch
                if ($virusCleanedInFile) {
                    # v3.1.0: Backup trước khi lưu
                    $backupOK = Backup-ExcelFile -FilePath $filePath
                    $workbook.Save()
                    $cleanedCount++
                    Write-Host "   ✅ [THÀNH CÔNG] Đã tiêu diệt virus và lưu lại tệp an toàn!" -ForegroundColor Green
                    Write-ScanLog "ĐÃ DIỆT: $filePath (Backup: $backupOK)"
                } else {
                    $safeCount++
                    Write-Host "   🛡️ [AN TOÀN] Tệp tin sạch, không có mã độc Kangatang." -ForegroundColor Gray
                    Write-ScanLog "AN TOÀN: $filePath"
                }
                
                $workbook.Close($false)
                Release-ComResource $workbook
                
            } catch {
                Write-Host "   ⚠️ [Lỗi xử lý tệp] $($_.Exception.Message)" -ForegroundColor DarkYellow
                Write-ScanLog "LỖI: $filePath - $($_.Exception.Message)"
                $errorCount++
                if ($workbook -ne $null) {
                    try { $workbook.Close($false) } catch {}
                    Release-ComResource $workbook
                }
            }
        }
    } finally {
        if ($excelApp -ne $null) {
            try { $excelApp.Quit() } catch {}
            Release-ComResource $excelApp
        }
        Invoke-GarbageClean
    }
    
    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host "   KẾT QUẢ QUÉT TẬP TIN HOÀN TẤT                                      " -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host "   - Tổng số file đã quét       : $($excelFiles.Count)" -ForegroundColor Cyan
    Write-Host "   - Số file đã diệt sạch virus : $cleanedCount" -ForegroundColor Green
    Write-Host "   - Số file an toàn            : $safeCount" -ForegroundColor Gray
    Write-Host "   - Số file lỗi / bỏ qua       : $errorCount" -ForegroundColor Yellow
    Write-ScanLog "KẾT QUẢ: Quét=$($excelFiles.Count), Diệt=$cleanedCount, Sạch=$safeCount, Lỗi=$errorCount"
}

# ==============================================================================
# THỰC THI VÀ TƯƠNG TÁC NGƯỜI DÙNG
# ==============================================================================
Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "✅ ĐÃ DỌN SẠCH CÁC TỆP MẦM BỆNH TRONG HỆ THỐNG (BƯỚC 0, 1, 2, 3, 3.5)" -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green

Write-Host "`nBạn có muốn tiếp tục quét và diệt virus Kangatang cho các tệp Excel trong thư mục làm việc không?" -ForegroundColor Cyan

# v3.1.0: Sử dụng FolderBrowserDialog thay vì nhập đường dẫn thủ công
$choice = Read-Host "Nhập Y để chọn thư mục quét (sẽ mở cửa sổ chọn), hoặc phím bất kỳ để kết thúc (Y/N)"

if ($choice -eq "Y" -or $choice -eq "y") {
    # v3.1.0: Mở cửa sổ GUI chọn thư mục
    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "KangatangGuard - Chọn thư mục chứa file Excel cần quét virus"
    $folderDialog.ShowNewFolderButton = $false
    
    $dialogResult = $folderDialog.ShowDialog()
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderDialog.SelectedPath
        Write-Host "`n   📂 Thư mục được chọn: $selectedFolder" -ForegroundColor Cyan
        Clean-ExcelVirusInFolder -TargetFolder $selectedFolder
    } else {
        Write-Host "`n   ⚠️ Đã hủy chọn thư mục." -ForegroundColor DarkYellow
    }
} else {
    Write-Host "`n✅ Đã hoàn thành! Toàn bộ mầm bệnh khởi động ngầm trong XLSTART và AppData đã được dọn sạch." -ForegroundColor Green
}

# v3.1.0: Hoàn nguyên Registry AccessVBOM
Restore-AccessVBOM

# v3.1.0: Xuất báo cáo kiểm toán
Export-ScanReport

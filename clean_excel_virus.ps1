# ==============================================================================
# HỆ THỐNG DIỆT VIRUS MACRO KANGATANG (Kanga/Laroux) VÀ FILE MẦM BỆNH MYPERSONNEL
# Phiên bản: v2.1.0 (Paste-Proof và Tối ưu chạy 1-Click)
# 1. Đóng toàn bộ tiến trình Excel đang chạy để giải phóng khóa tệp tin.
# 2. Tự động tìm và dọn dẹp TOÀN BỘ thư mục AppData\Roaming\Microsoft\Excel của mọi User (Admin, OHSUNG VINA, mrKienIT,...).
# 3. Tiêu diệt triệt để tệp mầm bệnh mypersonnel*.xls và các biến thể trong toàn bộ thư mục XLSTART.
# 4. Quét sâu, gỡ bỏ các VBA Module, Sheet ẩn độc hại và làm sạch mã lây nhiễm trong ThisWorkbook.
# ==============================================================================

# Đảm bảo mã hóa UTF-8 để hiển thị tiếng Việt chính xác trong Console
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   HE THONG DIET VIRUS MACRO KANGATANG - EXCEL CLEANER (v2.1.0)       " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# ==============================================================================
# HÀM HỖ TRỢ: GIẢI PHÓNG COM OBJECT & BỘ NHỚ
# ==============================================================================
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

# ==============================================================================
# HÀM HỖ TRỢ: TỰ ĐỘNG TÌM KIẾM TOÀN BỘ THƯ MỤC EXCEL VÀ XLSTART
# ==============================================================================
function Get-AllExcelPaths {
    $foundPaths = [System.Collections.Generic.List[string]]::new()
    
    # 1. Thư mục AppData của User hiện tại
    $currentUserAppData = Join-Path $env:APPDATA "Microsoft\Excel"
    if (Test-Path $currentUserAppData) {
        $foundPaths.Add($currentUserAppData)
    }
    
    # 2. Quét qua tất cả Profile trong C:\Users (Bao gồm OHSUNG VINA, Admin, mrKienIT,...)
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

function Get-AllXLStartPaths {
    $xlStartDirs = [System.Collections.Generic.List[string]]::new()
    
    # 1. Lấy từ toàn bộ thư mục AppData Excel của tất cả Users
    $excelDirs = Get-AllExcelPaths
    foreach ($dir in $excelDirs) {
        $userXLStart = Join-Path $dir "XLSTART"
        if (Test-Path $userXLStart) {
            if (-not $xlStartDirs.Contains($userXLStart)) {
                $xlStartDirs.Add($userXLStart)
            }
        }
    }
    
    # 2. Quét các thư mục XLSTART cấp hệ thống của Microsoft Office (Program Files và Program Files x86)
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
# BƯỚC 1: ĐÓNG TIẾN TRÌNH EXCEL
# ==============================================================================
Write-Host "`n[1/4] [!] Dang dong toan bo tien trinh Microsoft Excel..." -ForegroundColor Yellow
$excelProcesses = Get-Process -Name excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    $excelProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Host "   -> [DONE] Da dung $($excelProcesses.Count) tien trinh Excel dang chay." -ForegroundColor Green
} else {
    Write-Host "   -> [DONE] Khong co tien trinh Excel nao dang chay ngam." -ForegroundColor Green
}

# ==============================================================================
# BƯỚC 2: DỌN DẸP THƯ MỤC APPDATA EXCEL CỦA TOÀN BỘ NGƯỜI DÙNG
# ==============================================================================
Write-Host "`n[2/4] [!] Dang ra soat va don dep thu muc cau hinh AppData Excel..." -ForegroundColor Yellow
$allExcelDirs = Get-AllExcelPaths

if ($allExcelDirs.Count -eq 0) {
    Write-Host "   -> [NOTE] Khong tim thay thu muc cau hinh Microsoft Excel nao." -ForegroundColor DarkYellow
} else {
    Write-Host "   -> Phat hien $($allExcelDirs.Count) thu muc cau hinh Excel trong he thong." -ForegroundColor Cyan
    foreach ($excelDir in $allExcelDirs) {
        Write-Host "`n   [Thu muc] Dang xu ly: $excelDir" -ForegroundColor Gray
        $items = Get-ChildItem -Path $excelDir -Force -ErrorAction SilentlyContinue
        
        foreach ($item in $items) {
            # BẢO VỆ: Giữ lại tệp cấu hình hợp lệ Excel*.xlb, thư mục XLSTART, và Templates
            if ($item.Name -like "Excel*.xlb" -or $item.Name -eq "XLSTART" -or $item.Name -eq "Templates") {
                Write-Host "      [Giu lai] -> $($item.Name)" -ForegroundColor DarkGray
                continue
            }
            
            # Xóa các tệp / thư mục rác do virus sinh ra (mypersonnel*, file tạm *.tmp, v.v.)
            try {
                Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
                Write-Host "      [Da xoa rac]  -> $($item.Name)" -ForegroundColor Red
            } catch {
                Write-Host "      [Canh bao] Khong the xoa $($item.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
    }
}

# ==============================================================================
# BƯỚC 3: TIÊU DIỆT TỆP MẦM BỆNH TRONG TẤT CẢ THƯ MỤC XLSTART
# ==============================================================================
Write-Host "`n[3/4] [!] Dang truy quet o dich trong TAT CA thu muc XLSTART..." -ForegroundColor Yellow
$allXLStartDirs = Get-AllXLStartPaths

if ($allXLStartDirs.Count -eq 0) {
    Write-Host "   -> [NOTE] Khong tim thay thu muc XLSTART nao tren may." -ForegroundColor DarkYellow
} else {
    Write-Host "   -> Phat hien $($allXLStartDirs.Count) thu muc XLSTART can lam sach." -ForegroundColor Cyan
    
    # Danh sách mẫu tên tệp virus Kangatang/Laroux thường dùng làm mầm bệnh
    $virusPatterns = @("mypersonnel*", "*kangatang*", "*kangaatang*", "*kanga*.xls*", "personal.xls")
    
    foreach ($xlDir in $allXLStartDirs) {
        Write-Host "`n   [XLSTART] Kiem tra: $xlDir" -ForegroundColor Gray
        
        # 1. Tìm và xóa file mầm bệnh theo tên
        $infectedFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        foreach ($pattern in $virusPatterns) {
            $matched = Get-ChildItem -Path $xlDir -Filter $pattern -File -Force -ErrorAction SilentlyContinue
            if ($matched) {
                foreach ($m in $matched) {
                    if (-not $infectedFiles.Contains($m)) {
                        $infectedFiles.Add($m)
                    }
                }
            }
        }
        
        if ($infectedFiles.Count -eq 0) {
            Write-Host "      -> [DONE] Khong phat hien file mam benh mypersonnel/kangatang." -ForegroundColor Green
        } else {
            foreach ($vFile in $infectedFiles) {
                try {
                    Remove-Item -Path $vFile.FullName -Force -ErrorAction Stop
                    Write-Host "      [DA TIEU DIET] -> Xoa thanh cong file mam benh: $($vFile.Name)" -ForegroundColor Red
                } catch {
                    Write-Host "      [Loi xoa file] -> $($vFile.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
            }
        }
    }
}

# ==============================================================================
# BƯỚC 4: HÀM QUÉT & LÀM SẠCH FILE EXCEL TRONG THƯ MỤC TÙY CHỌN
# ==============================================================================
function Clean-ExcelVirusInFolder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TargetFolder
    )
    
    # Chuẩn hóa đường dẫn: loại bỏ dấu ngoặc kép hoặc ngoặc đơn nếu có
    $TargetFolder = $TargetFolder.Replace([char]34, [string]::Empty).Replace([char]39, [string]::Empty).Trim()
    
    if (-not (Test-Path $TargetFolder)) {
        Write-Host "[Loi] Duong dan thu muc quet khong ton tai: $TargetFolder" -ForegroundColor Red
        return
    }
    
    Write-Host "`n[4/4] [!] Khoi dong trinh quet VBA Engine trong thu muc: $TargetFolder" -ForegroundColor Yellow
    
    # 1. Tìm tất cả các file Excel có thể chứa macro / virus
    Write-Host "Dang tim kiem cac tep Excel..." -ForegroundColor Gray
    $excelFiles = Get-ChildItem -Path $TargetFolder -Recurse -File -ErrorAction SilentlyContinue | 
                  Where-Object { $_.Extension -match '^\.(xlsm|xls|xlsb|xltm|xlt|xla|xlam)$' }
    
    if (-not $excelFiles -or $excelFiles.Count -eq 0) {
        Write-Host "-> Khong tim thay tep Excel (.xls, .xlsm, .xlsb,...) nao trong thu muc duoc chon." -ForegroundColor Green
        return
    }
    
    Write-Host "-> Da tim thay $($excelFiles.Count) tep Excel can quet cau truc." -ForegroundColor Cyan
    
    # 2. Khởi tạo Excel COM Object ẩn
    $excelApp = $null
    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $excelApp.ScreenUpdating = $false
        $excelApp.EnableEvents = $false
        # Vô hiệu hóa thực thi macro tự động khi mở file để tránh virus tự kích hoạt
        try { $excelApp.AutomationSecurity = 3 } catch {}
    } catch {
        Write-Host "[Loi] Khong the khoi dong Excel COM Application. Vui long dam bao MS Excel da cai dat." -ForegroundColor Red
        return
    }
    
    $cleanedCount = 0
    $errorCount = 0
    $safeCount = 0
    
    try {
        foreach ($file in $excelFiles) {
            $filePath = $file.FullName
            Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "[Kiem tra] $filePath" -ForegroundColor Cyan
            
            $workbook = $null
            $virusCleanedInFile = $false
            
            try {
                # Mở Workbook ở chế độ ghi để làm sạch
                $workbook = $excelApp.Workbooks.Open($filePath, 0, $false, 5, "", "", $true)
                
                # Kiểm tra VBProject
                $project = $null
                try {
                    $project = $workbook.VBProject
                } catch {
                    Write-Host "   [Bo qua] Khong truy cap duoc VBProject. (Can bat 'Trust access to the VBA project object model' trong Excel Trust Center)." -ForegroundColor DarkYellow
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
                            Write-Host "   [PHAT HIEN] Tim thay Module ma doc: $compName" -ForegroundColor Red
                            try {
                                $components.Remove($comp)
                                $virusCleanedInFile = $true
                            } catch {
                                Write-Host "      [Loi xoa Module] $($_.Exception.Message)" -ForegroundColor DarkYellow
                            }
                        } else {
                            # Kiểm tra mã bên trong ThisWorkbook hoặc Sheet code module
                            try {
                                $codeMod = $comp.CodeModule
                                if ($codeMod -and $codeMod.CountOfLines -gt 0) {
                                    $content = $codeMod.Lines(1, $codeMod.CountOfLines)
                                    if ($content -match "Kangatang|Kangaatang|mypersonnel") {
                                        Write-Host "   [PHAT HIEN] Tim thay ma doc lay nhiem trong: $compName" -ForegroundColor Red
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
                        Write-Host "   [PHAT HIEN] Tim thay Sheet an ma doc: $sheetName" -ForegroundColor Red
                        try {
                            # Unhide sheet trước khi xóa nếu cần
                            $sheet.Visible = -1
                            $sheet.Delete()
                            $virusCleanedInFile = $true
                        } catch {
                            Write-Host "      [Loi xoa Sheet] $($_.Exception.Message)" -ForegroundColor DarkYellow
                        }
                    }
                    Release-ComResource $sheet
                }
                
                # Lưu file nếu phát hiện và đã làm sạch
                if ($virusCleanedInFile) {
                    $workbook.Save()
                    $cleanedCount++
                    Write-Host "   [THANH CONG] Da tieu diet virus va luu lai tep an toan!" -ForegroundColor Green
                } else {
                    $safeCount++
                    Write-Host "   [AN TOAN] Tep tin sach, khong co ma doc Kangatang." -ForegroundColor Gray
                }
                
                $workbook.Close($false)
                Release-ComResource $workbook
                
            } catch {
                Write-Host "   [Loi xu ly tep] $($_.Exception.Message)" -ForegroundColor DarkYellow
                $errorCount++
                if ($workbook -ne $null) {
                    try { $workbook.Close($false) } catch {}
                    Release-ComResource $workbook
                }
            }
        }
    } finally {
        # Đảm bảo đóng và giải phóng hoàn toàn Excel Application
        if ($excelApp -ne $null) {
            try { $excelApp.Quit() } catch {}
            Release-ComResource $excelApp
        }
        Invoke-GarbageClean
    }
    
    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host "   KET QUA QUET TAP TIN HOAN TAT                                      " -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host "   - Tong so file da quet      : $($excelFiles.Count)" -ForegroundColor Cyan
    Write-Host "   - So file da diet sach virus : $cleanedCount" -ForegroundColor Green
    Write-Host "   - So file an toan           : $safeCount" -ForegroundColor Gray
    Write-Host "   - So file loi / bo qua      : $errorCount" -ForegroundColor Yellow
}

# ==============================================================================
# THỰC THI & TƯƠNG TÁC NGƯỜI DÙNG
# ==============================================================================
Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "[DONE] DA DON SACH CAC TEP MAM BENH TRONG HE THONG (BUOC 1, 2, 3)" -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green

Write-Host "`nBan co muon tiep tuc quet va diet virus Kangatang cho cac tep Excel trong thu muc lam viec khong?" -ForegroundColor Cyan
$choice = Read-Host "Nhap 'Y' de tiep tuc quet thu muc, hoac phim bat ky de ket thuc (Y/N)"

if ($choice -eq "Y" -or $choice -eq "y") {
    $folderToScan = Read-Host "`nHay dan duong dan thu muc can quet (Vi du: D:\DuLieu hoac C:\Users\OHSUNG VINA\Documents)"
    Clean-ExcelVirusInFolder -TargetFolder $folderToScan
} else {
    Write-Host "`nDa hoan thanh! Toan bo mam benh khoi dong ngam trong XLSTART va AppData da duoc don sach." -ForegroundColor Green
}
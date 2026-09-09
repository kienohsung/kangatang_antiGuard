# ==============================================================================
# Install_KangatangGuard.ps1
# PowerShell Installer: Tao Excel Add-in (.xlam) va cai dat vao XLSTART
# Phien ban: v3.5.3
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - INSTALLER v3.5.3                                  " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# ==============================================================================
# BUOC 1: BAT REGISTRY AccessVBOM CHO OFFICE (BAT BUOC DE INJECT VBA)
# ==============================================================================
Write-Host "`n[1/3] Dang bat quyen truy cap VBA Project Object Model (AccessVBOM)..." -ForegroundColor Yellow

$officeVersions = @("14.0", "15.0", "16.0")  # Office 2010, 2013, 2016/2019/365
$regBackup = @{}

foreach ($ver in $officeVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $regPath) {
        # Luu gia tri cu de co the rollback
        try {
            $currentVal = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
            if ($null -eq $currentVal) { $currentVal = 0 }
            $regBackup[$regPath] = $currentVal
        } catch {
            $regBackup[$regPath] = 0
        }
        
        # Bat AccessVBOM = 1
        try {
            Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction Stop
            Write-Host "   -> [OK] Da bat AccessVBOM cho Office $ver" -ForegroundColor Green
        } catch {
            Write-Host "   -> [Loi] Khong the ghi Registry Office $ver : $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
}

if ($regBackup.Count -eq 0) {
    Write-Host "   -> [Canh bao] Khong tim thay cai dat Office nao trong Registry." -ForegroundColor DarkYellow
    Write-Host "   -> Vui long kiem tra Excel da duoc cai dat tren may nay." -ForegroundColor DarkYellow
}

# ==============================================================================
# BUOC 2: TAO FILE .XLAM BANG EXCEL COM VA INJECT MA VBA
# ==============================================================================
Write-Host "`n[2/3] Dang tao file KangatangGuard.xlam..." -ForegroundColor Yellow

$vbaCodeFile = Join-Path $ScriptDir "KangatangGuard_Code.vba"
if (-not (Test-Path $vbaCodeFile)) {
    Write-Host "   [LOI] Khong tim thay file KangatangGuard_Code.vba!" -ForegroundColor Red
    Write-Host "   File can nam cung thu muc voi script nay." -ForegroundColor Red
    exit 1
}

# Doc noi dung file VBA
$vbaContent = [System.IO.File]::ReadAllText($vbaCodeFile, [System.Text.Encoding]::UTF8)

# Parse sections
function Get-VbaSection {
    param ([string]$Content, [string]$SectionName)
    $startMarker = "'### SECTION: $SectionName ###"
    $endMarker = "'### END_SECTION: $SectionName ###"
    
    $startIdx = $Content.IndexOf($startMarker)
    $endIdx = $Content.IndexOf($endMarker)
    
    if ($startIdx -ge 0 -and $endIdx -ge 0) {
        $codeStart = $startIdx + $startMarker.Length
        $code = $Content.Substring($codeStart, $endIdx - $codeStart).Trim()
        # Xoa dong comment dau tien (mo ta section)
        $lines = $code -split "`r?`n"
        $filteredLines = $lines | Where-Object { $_ -notmatch "^'---.*---$" }
        return ($filteredLines -join "`r`n").Trim()
    }
    return ""
}

$thisWorkbookCode = Get-VbaSection -Content $vbaContent -SectionName "ThisWorkbook"
$clsAppEventsCode = Get-VbaSection -Content $vbaContent -SectionName "clsAppEvents"
$modScannerCode   = Get-VbaSection -Content $vbaContent -SectionName "modKangatangScanner"
$modLoggerCode    = Get-VbaSection -Content $vbaContent -SectionName "modLogger"

# Xac dinh duong dan XLSTART
$xlStartPath = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (-not (Test-Path $xlStartPath)) {
    New-Item -ItemType Directory -Path $xlStartPath -Force | Out-Null
}

$xlamPath = Join-Path $xlStartPath "KangatangGuard.xlam"

# Dong tat ca tien trinh Excel dang chay
Write-Host "   Dang dong cac tien trinh Excel..." -ForegroundColor Gray
$excelProcs = Get-Process -Name excel -ErrorAction SilentlyContinue
if ($excelProcs) {
    $excelProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# Khoi tao Excel COM
$excelApp = $null
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.ScreenUpdating = $false
    $excelApp.EnableEvents = $false
} catch {
    Write-Host "   [LOI] Khong the khoi dong Excel COM: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

try {
    # Tao workbook moi
    $wb = $excelApp.Workbooks.Add()
    
    # Ham helper chen ma VBA dam bao dung 1 Option Explicit o dau
    function Inject-CleanVbaModule($component, $sourceCode) {
        if (-not $sourceCode) { return }
        $cm = $component.CodeModule
        if ($cm.CountOfLines -gt 0) {
            $cm.DeleteLines(1, $cm.CountOfLines)
        }
        # Loc bo moi dong Option Explicit va dong chu thich tieu de cu
        $lines = $sourceCode -split "`r?`n"
        $bodyLines = $lines | Where-Object {
            $trimmed = $_.Trim()
            $trimmed -ne "Option Explicit" -and $trimmed -notmatch "^'---.*---$"
        }
        $cleanBody = ($bodyLines -join "`r`n").Trim()
        $finalCode = "Option Explicit`r`n`r`n" + $cleanBody
        $cm.AddFromString($finalCode)
    }

    # --- Inject ThisWorkbook code ---
    Write-Host "   -> Inject code vao ThisWorkbook..." -ForegroundColor Gray
    $twb = $wb.VBProject.VBComponents.Item("ThisWorkbook")
    Inject-CleanVbaModule $twb $thisWorkbookCode
    
    # --- Tao Class Module: clsAppEvents ---
    Write-Host "   -> Tao Class Module: clsAppEvents..." -ForegroundColor Gray
    $clsComp = $wb.VBProject.VBComponents.Add(2)  # 2 = vbext_ct_ClassModule
    $clsComp.Name = "clsAppEvents"
    Inject-CleanVbaModule $clsComp $clsAppEventsCode
    
    # --- Tao Standard Module: modKangatangScanner ---
    Write-Host "   -> Tao Module: modKangatangScanner..." -ForegroundColor Gray
    $modScanner = $wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    $modScanner.Name = "modKangatangScanner"
    Inject-CleanVbaModule $modScanner $modScannerCode
    
    # --- Tao Standard Module: modLogger ---
    Write-Host "   -> Tao Module: modLogger..." -ForegroundColor Gray
    $modLog = $wb.VBProject.VBComponents.Add(1)
    $modLog.Name = "modLogger"
    Inject-CleanVbaModule $modLog $modLoggerCode
    
    # Xoa Sheet1 mac dinh (add-in khong can sheet)
    # Giu lai 1 sheet vi Excel bat buoc phai co it nhat 1
    
    # Luu file dang .xlam (55 = xlAddIn format)
    Write-Host "   -> Luu file .xlam..." -ForegroundColor Gray
    
    # Xoa file cu neu ton tai (go bo thuoc tinh Read-Only truoc neu co)
    if (Test-Path $xlamPath) {
        try { Set-ItemProperty -Path $xlamPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
        Remove-Item -Path $xlamPath -Force -ErrorAction SilentlyContinue
    }
    
    # Cau hinh thuoc tinh Add-in chuan
    $wb.IsAddin = $true
    $wb.SaveAs($xlamPath, 55)  # 55 = xlOpenXMLAddIn (.xlam)
    $wb.Close($false)
    
    # Dat thuoc tinh Read-Only cho file trong XLSTART
    try {
        Set-ItemProperty -Path $xlamPath -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue
    } catch {}
    
    Write-Host "   -> [OK] Da luu vao XLSTART: $xlamPath" -ForegroundColor Green
    
    # ==============================================================================
    # DUAL REGISTRATION: SAO CHEP VA DANG KY VAO ADDINS CHO OFFICE 365
    # ==============================================================================
    Write-Host "`n[DUAL-REG] Dang thiet lap Add-in Registry cho Office 365..." -ForegroundColor Yellow
    $addInsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
    if (-not (Test-Path $addInsDir)) {
        New-Item -ItemType Directory -Path $addInsDir -Force | Out-Null
    }
    
    $xlamAddInsPath = Join-Path $addInsDir "KangatangGuard.xlam"
    if (Test-Path $xlamAddInsPath) {
        try { Set-ItemProperty -Path $xlamAddInsPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
        Remove-Item -Path $xlamAddInsPath -Force -ErrorAction SilentlyContinue
    }
    
    Copy-Item -Path $xlamPath -Destination $xlamAddInsPath -Force
    try { Set-ItemProperty -Path $xlamAddInsPath -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}
    Write-Host "   -> [OK] Da sao chep sang AddIns: $xlamAddInsPath" -ForegroundColor Green
    
    # Dang ky khoa OPEN* trong Registry Options cho Office 16.0, 15.0, 14.0
    foreach ($ver in @("16.0", "15.0", "14.0")) {
        $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
        if (Test-Path $optKey) {
            $props = (Get-ItemProperty -Path $optKey -ErrorAction SilentlyContinue).psobject.Properties
            $alreadyRegistered = $false
            $maxOpenIndex = -1
            $targetSlot = $null
            
            # Kiem tra xem da dang ky chua
            foreach ($p in $props) {
                if ($p.Name -eq "OPEN" -or $p.Name -match "^OPEN(\d+)$") {
                    $val = [string]$p.Value
                    if ($val -like "*KangatangGuard.xlam*") {
                        $alreadyRegistered = $true
                        break
                    }
                    if ($p.Name -eq "OPEN") {
                        if ($maxOpenIndex -lt 0) { $maxOpenIndex = 0 }
                    } elseif ($p.Name -match "^OPEN(\d+)$") {
                        $idxNum = [int]$matches[1]
                        if ($idxNum -gt $maxOpenIndex) { $maxOpenIndex = $idxNum }
                    }
                }
            }
            
            if (-not $alreadyRegistered) {
                if ($maxOpenIndex -lt 0) {
                    # Chua co khoa OPEN nao
                    $targetSlot = "OPEN"
                } else {
                    $nextNum = $maxOpenIndex + 1
                    $targetSlot = "OPEN$nextNum"
                }
                
                $regValue = "/R `"$xlamAddInsPath`""
                Set-ItemProperty -Path $optKey -Name $targetSlot -Value $regValue -Type String -Force -ErrorAction SilentlyContinue
                Write-Host "   -> [OK] Da dang ky Registry $optKey\$targetSlot" -ForegroundColor Green
            } else {
                Write-Host "   -> [OK] Add-in da duoc dang ky san trong Registry Office $ver" -ForegroundColor Green
            }
        }
    }
    
} catch {
    Write-Host "   [LOI] Loi khi tao file .xlam: $($_.Exception.Message)" -ForegroundColor Red
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch {}
    }
} finally {
    if ($excelApp -ne $null) {
        try { $excelApp.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ==============================================================================
# BUOC 3: TAO THU MUC LOG VA CAI DAT WORKER SCANNER DOC LAP (v3.5.3)
# ==============================================================================
Write-Host "`n[3/3] Dang thiet lap thu muc Log va cai dat Background Worker..." -ForegroundColor Yellow
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
Write-Host "   -> [OK] Thu muc Log: $logDir" -ForegroundColor Green

# Sao chep Kangatang_FolderScanner.ps1 vao APPDATA\KangatangGuard
$srcWorker = Join-Path $ScriptDir "Kangatang_FolderScanner.ps1"
if (Test-Path $srcWorker) {
    $dstWorker = Join-Path $logDir "Kangatang_FolderScanner.ps1"
    Copy-Item -Path $srcWorker -Destination $dstWorker -Force
    Write-Host "   -> [OK] Da cai dat Background Worker: $dstWorker" -ForegroundColor Green
} else {
    Write-Host "   -> [CANH BAO] Khong tim thay $srcWorker de sao chep." -ForegroundColor DarkYellow
}

# ==============================================================================
# HOAN TAT
# ==============================================================================
Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "   CAI DAT KANGATANG GUARD HOAN TAT!                                    " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   File Add-in    : $xlamPath" -ForegroundColor Cyan
Write-Host "   Thu muc Log    : $logDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "   Huong dan su dung:" -ForegroundColor Yellow
Write-Host "   1. Mo Excel binh thuong (khong can lam gi them)" -ForegroundColor White
Write-Host "   2. Add-in se TU DONG quet moi file ban mo" -ForegroundColor White
Write-Host "   3. Menu 'KangatangGuard' xuat hien tren thanh cong cu" -ForegroundColor White
Write-Host "   4. Kiem tra log tai: $logDir" -ForegroundColor White
Write-Host ""

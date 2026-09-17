# ==============================================================================
# Install_KangatangGuard.ps1
# PowerShell Installer: Tao Excel Add-in (.xlam) va cai dat vao XLSTART
# Phien ban: v3.8.7 (Modern HD Large Icon Ribbon & Refined Toolbar Icons)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$InstallerVersion = "3.8.7"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - INSTALLER v$InstallerVersion                                  " -ForegroundColor Cyan
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

    # Ham helper tich hop Ribbon CustomUI XML (Icon HD Large 32x32)
    function Add-RibbonCustomUI([string]$targetXlam) {
        Write-Host "   -> Tich hop Ribbon CustomUI XML (Icon HD 32x32)..." -ForegroundColor Gray
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        
        $zip = [System.IO.Compression.ZipFile]::Open($targetXlam, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            # 1. Cap nhat _rels/.rels
            $relsEntry = $zip.GetEntry("_rels/.rels")
            $relsDoc = New-Object System.Xml.XmlDocument
            $relsStream = $relsEntry.Open()
            $relsDoc.Load($relsStream)
            $relsStream.Close()

            $ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            $hasCustomUi = $false
            foreach ($rel in $relsDoc.Relationships.Relationship) {
                if ($rel.Type -like "*customui*") {
                    $hasCustomUi = $true
                    break
                }
            }

            if (-not $hasCustomUi) {
                $newRel = $relsDoc.CreateElement("Relationship", $ns)
                $newRel.SetAttribute("Id", "rIdCustomUI14")
                $newRel.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
                $newRel.SetAttribute("Target", "customUI/customUI14.xml")
                $relsDoc.Relationships.AppendChild($newRel) | Out-Null

                $relsEntry.Delete()
                $newRelsEntry = $zip.CreateEntry("_rels/.rels", [System.IO.Compression.CompressionLevel]::Optimal)
                $newRelsStream = $newRelsEntry.Open()
                $relsDoc.Save($newRelsStream)
                $newRelsStream.Close()
            }

            # 2. Tao customUI/customUI14.xml voi 2 tabs: Tab Kangatang Guard chuyen biet & Tab AddIns
            $customUiXml = @'
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabKangatangGuard" label="Kangatang Guard">
        <group id="grpScan" label="Diệt Virus &amp; Bảo vệ">
          <button id="btnScanFile" label="Quét tệp này" size="large" imageMso="FileCheckOut" onAction="Ribbon_ScanActiveWorkbook" screentip="Quét tệp này" supertip="Quét và tiêu diệt virus macro trên tệp Excel đang mở." />
          <button id="btnScanFolder" label="Quét thư mục..." size="large" imageMso="FolderBrowse" onAction="Ribbon_ScanFolderDialog" screentip="Quét thư mục..." supertip="Quét ngầm toàn bộ thư mục mà không làm gián đoạn Excel." />
          <button id="btnResume" label="Tiếp tục quét" size="large" imageMso="PlayMacro" onAction="Ribbon_ResumeScanDialog" screentip="Tiếp tục phiên quét" supertip="Tiếp tục phiên quét dở dang trước đó với tốc độ cao." />
        </group>
        <group id="grpTools" label="Hệ thống &amp; Tiện ích">
          <button id="btnLog" label="Nhật ký (Log)" size="large" imageMso="OpenReportDetails" onAction="Ribbon_OpenLogFolder" screentip="Nhật ký kiểm toán" supertip="Mở thư mục chứa nhật ký quét và diệt virus." />
          <button id="btnUpdate" label="Cập nhật LAN" size="large" imageMso="ServerRefresh" onAction="Ribbon_CheckForLanUpdates" screentip="Cập nhật mạng LAN" supertip="Kiểm tra và cập nhật phiên bản mới từ máy chủ LAN." />
          <button id="btnAbout" label="Thông tin" size="large" imageMso="Info" onAction="Ribbon_ShowAbout" screentip="Thông tin" supertip="Thông tin phiên bản và xuất xứ KangatangGuard." />
        </group>
      </tab>
      <tab idMso="TabAddIns">
        <group id="grpKangatangAddins" label="Kangatang Guard v3.8.7">
          <button id="btnScanFileA" label="Quét tệp này" size="large" imageMso="FileCheckOut" onAction="Ribbon_ScanActiveWorkbook" screentip="Quét tệp này" supertip="Quét và tiêu diệt virus macro trên tệp Excel đang mở." />
          <button id="btnScanFolderA" label="Quét thư mục..." size="large" imageMso="FolderBrowse" onAction="Ribbon_ScanFolderDialog" screentip="Quét thư mục..." supertip="Quét ngầm toàn bộ thư mục mà không làm gián đoạn Excel." />
          <button id="btnResumeA" label="Tiếp tục quét" size="large" imageMso="PlayMacro" onAction="Ribbon_ResumeScanDialog" screentip="Tiếp tục phiên quét" supertip="Tiếp tục phiên quét dở dang trước đó với tốc độ cao." />
          <separator id="sepA1" />
          <button id="btnLogA" label="Nhật ký" size="large" imageMso="OpenReportDetails" onAction="Ribbon_OpenLogFolder" screentip="Nhật ký kiểm toán" supertip="Mở thư mục chứa nhật ký quét và diệt virus." />
          <button id="btnUpdateA" label="Cập nhật LAN" size="large" imageMso="ServerRefresh" onAction="Ribbon_CheckForLanUpdates" screentip="Cập nhật mạng LAN" supertip="Kiểm tra và cập nhật phiên bản mới từ máy chủ LAN." />
          <button id="btnAboutA" label="Thông tin" size="large" imageMso="Info" onAction="Ribbon_ShowAbout" screentip="Thông tin" supertip="Thông tin phiên bản và xuất xứ KangatangGuard." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@
            $uiEntry = $zip.GetEntry("customUI/customUI14.xml")
            if ($uiEntry) { $uiEntry.Delete() }
            $newUiEntry = $zip.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
            $uiStream = $newUiEntry.Open()
            $uiWriter = New-Object System.IO.StreamWriter($uiStream, [System.Text.Encoding]::UTF8)
            $uiWriter.Write($customUiXml)
            $uiWriter.Flush()
            $uiWriter.Close()
            Write-Host "   -> [OK] Da tich hop Ribbon CustomUI XML (Large 32x32 Icons) thanh cong!" -ForegroundColor Green
        } catch {
            Write-Host "   -> [CANH BAO] Khong the tich hop Ribbon XML: $($_.Exception.Message)" -ForegroundColor Yellow
        } finally {
            $zip.Dispose()
        }
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
    
    # Dam bao file trong XLSTART khong bi danh dau Read-Only
    try {
        Set-ItemProperty -Path $xlamPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
    } catch {}
    
    Write-Host "   -> [OK] Da luu vao XLSTART: $xlamPath" -ForegroundColor Green
    
    # Tich hop Ribbon CustomUI XML (Icon HD 32x32)
    Add-RibbonCustomUI -targetXlam $xlamPath
    
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
    try { Set-ItemProperty -Path $xlamAddInsPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
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
# BUOC 3: TAO THU MUC LOG, SESSIONS VA CAI DAT WORKER SCANNER DOC LAP (v3.6.0)
# ==============================================================================
Write-Host "`n[3/3] Dang thiet lap thu muc Log, Sessions va cai dat Background Worker..." -ForegroundColor Yellow
$logDir = Join-Path $env:APPDATA "KangatangGuard"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$sessionsDir = Join-Path $logDir "Sessions"
if (-not (Test-Path $sessionsDir)) {
    New-Item -ItemType Directory -Path $sessionsDir -Force | Out-Null
}
Write-Host "   -> [OK] Thu muc Log: $logDir" -ForegroundColor Green
Write-Host "   -> [OK] Thu muc Sessions: $sessionsDir" -ForegroundColor Green

# Sao chep Kangatang_FolderScanner.ps1 vao APPDATA\KangatangGuard (Luu y: KHONG chep vao XLSTART vi Excel se tu mo script nhu workbook)
$srcWorker = Join-Path $ScriptDir "Kangatang_FolderScanner.ps1"
if (Test-Path $srcWorker) {
    # 1. Thu muc hoat dong mac dinh APPDATA\KangatangGuard
    $dstWorker1 = Join-Path $logDir "Kangatang_FolderScanner.ps1"
    Copy-Item -Path $srcWorker -Destination $dstWorker1 -Force
    Write-Host "   -> [OK] Da cai dat Background Worker vao KangatangGuard: $dstWorker1" -ForegroundColor Green
    
    # 2. Xoa khoi XLSTART neu vo tinh co (tranh Excel tu mo va quet nham)
    $badXlStartWorker = Join-Path $xlStartPath "Kangatang_FolderScanner.ps1"
    if (Test-Path $badXlStartWorker) {
        Remove-Item -Path $badXlStartWorker -Force -ErrorAction SilentlyContinue
    }
    
    # 3. Dang ky ScannerScript vao Registry de Add-in uu tien doc tu vi tri an toan (Chong Antivirus/EDR xoa file trong APPDATA)
    $guardRegKey = "HKCU:\Software\KangatangGuard"
    if (-not (Test-Path $guardRegKey)) {
        New-Item -Path $guardRegKey -Force | Out-Null
    }
    Set-ItemProperty -Path $guardRegKey -Name "ScannerScript" -Value $srcWorker -Force
    Set-ItemProperty -Path $guardRegKey -Name "InstalledVersion" -Value $InstallerVersion -Force
    Set-ItemProperty -Path $guardRegKey -Name "UpdateSource" -Value "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang" -Force
    Write-Host "   -> [OK] Da dang ky ScannerScript vao Registry: $srcWorker" -ForegroundColor Green
    Write-Host "   -> [OK] Da dang ky InstalledVersion: v$InstallerVersion" -ForegroundColor Green
    Write-Host "   -> [OK] Da dang ky UpdateSource: \\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang" -ForegroundColor Green
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

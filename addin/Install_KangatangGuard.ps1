# ==============================================================================
# Install_KangatangGuard.ps1
# PowerShell Installer: Tao Excel Add-in (.xlam) va cai dat vao %APPDATA%\Microsoft\AddIns
# Phien ban: v3.20.1 (100% Native VBA Folder Scanner & Zero-PS1 Antivirus-Safe Architecture)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$InstallerVersion = "3.20.1"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD - INSTALLER v$InstallerVersion                                  " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# ==============================================================================
# BUOC 1: BAT REGISTRY AccessVBOM CHO OFFICE (BAT BUOC DE INJECT & QUET VBA)
# ==============================================================================
Write-Host "`n[1/3] Dang bat quyen truy cap VBA Project Object Model (AccessVBOM)..." -ForegroundColor Yellow

$officeVersions = @("12.0", "14.0", "15.0", "16.0")  # Office 2007, 2010, 2013, 2016/2019/2021/365
$regBackup = @{}

foreach ($ver in $officeVersions) {
    $excelKey = "HKCU:\Software\Microsoft\Office\$ver\Excel"
    $regPath  = "$excelKey\Security"
    if (-not (Test-Path $regPath) -and (Test-Path $excelKey)) {
        New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
    }
    if (Test-Path $regPath) {
        try {
            $currentVal = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
            if ($null -eq $currentVal) { $currentVal = 0 }
            $regBackup[$regPath] = $currentVal
        } catch {
            $regBackup[$regPath] = 0
        }
        
        try {
            Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction Stop
            Set-ItemProperty -Path $regPath -Name "AllowNetworkLocations" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            Write-Host "   -> [OK] Da bat AccessVBOM cho Office $ver (HKCU)" -ForegroundColor Green
        } catch {
            Write-Host "   -> [Loi] Khong the ghi Registry Office $ver : $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
    
    # Go bo Policy HKCU neu dang khoa AccessVBOM = 0
    $polPath = "HKCU:\Software\Policies\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $polPath) {
        try {
            Remove-ItemProperty -Path $polPath -Name "AccessVBOM" -Force -ErrorAction SilentlyContinue
        } catch {}
    }
    
    # Thu bat tren HKLM va WOW6432Node neu co quyen
    foreach ($hklmSec in @(
        "HKLM:\Software\Microsoft\Office\$ver\Excel\Security",
        "HKLM:\Software\WOW6432Node\Microsoft\Office\$ver\Excel\Security"
    )) {
        if (Test-Path $hklmSec) {
            try {
                Set-ItemProperty -Path $hklmSec -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    }
}

# Xoa LastVbomWarn cu neu co
if (Test-Path "HKCU:\Software\KangatangGuard") {
    Remove-ItemProperty -Path "HKCU:\Software\KangatangGuard" -Name "LastVbomWarn" -Force -ErrorAction SilentlyContinue
}

if ($regBackup.Count -eq 0) {
    # Dam bao toi thieu 16.0 va 15.0 luon duoc tao san
    foreach ($defVer in @("16.0", "15.0")) {
        $defPath = "HKCU:\Software\Microsoft\Office\$defVer\Excel\Security"
        New-Item -Path $defPath -Force -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $defPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    }
    Write-Host "   -> [OK] Da khoi tao san AccessVBOM = 1 cho Office 15.0/16.0." -ForegroundColor Green
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
$modShieldCode    = Get-VbaSection -Content $vbaContent -SectionName "modKangatangShield"
$modThreatCode    = Get-VbaSection -Content $vbaContent -SectionName "modThreatCollector"
$modLoggerCode    = Get-VbaSection -Content $vbaContent -SectionName "modLogger"

# v3.11.0: Chan build neu thieu cac core section bat buoc
foreach ($sec in @(@("ThisWorkbook",$thisWorkbookCode), @("clsAppEvents",$clsAppEventsCode), @("modKangatangScanner",$modScannerCode), @("modKangatangShield",$modShieldCode), @("modLogger",$modLoggerCode))) {
    if ([string]::IsNullOrWhiteSpace($sec[1])) {
        Write-Host "   [LOI] Section VBA rong hoac thieu: $($sec[0])" -ForegroundColor Red
        exit 1
    }
}

# Xac dinh duong dan AddIns (chuan Excel Add-Ins Manager) va XLSTART
$addInsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path $addInsDir)) {
    New-Item -ItemType Directory -Path $addInsDir -Force | Out-Null
}
$xlStartPath = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (-not (Test-Path $xlStartPath)) {
    New-Item -ItemType Directory -Path $xlStartPath -Force | Out-Null
}

$xlamPath = Join-Path $addInsDir "KangatangGuard.xlam"

# Chi dong Excel neu KHONG co tien trinh quet thu muc dang chay
$scanActive = $false
$sessCheckFile = Join-Path $env:APPDATA "KangatangGuard\last_session.json"
if (Test-Path $sessCheckFile) {
    try {
        $sItem = Get-Item $sessCheckFile -ErrorAction SilentlyContinue
        $sMeta = Get-Content $sessCheckFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($sMeta.Status -eq "IN_PROGRESS" -and $sItem -and (((Get-Date) - $sItem.LastWriteTime).TotalSeconds -lt 60)) {
            $scanActive = $true
        }
    } catch {}
}
if (-not $scanActive) {
    Write-Host "   Dang dong cac tien trinh Excel..." -ForegroundColor Gray
    $excelProcs = Get-Process -Name excel -ErrorAction SilentlyContinue
    if ($excelProcs) {
        $excelProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
} else {
    Write-Host "   -> Phat hien phien quet thu muc dang chay, giu nguyen cac Worker Excel hien tai..." -ForegroundColor Cyan
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

            # 2. Tao customUI/customUI14.xml voi 1 tab duy nhat: Tab "Kangatang Guard" chuyen biet tren Ribbon
            $customUiFile = Join-Path $ScriptDir "customUI14.xml"
            $uiBytes = $null
            if (Test-Path $customUiFile) {
                # Doc truc tiep byte mang UTF-8 tu file template - KHONG bi anh huong boi PowerShell string encoding
                $uiBytes = [System.IO.File]::ReadAllBytes($customUiFile)
            } else {
                # Fallback: Sinh tu chuoi UTF-8 chuan
                $customUiXml = @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabKangatangGuard" label="Kangatang Guard">
        <group id="grpScan" label="Diệt Virus &amp; Bảo vệ">
          <button id="btnScanFile" label="Quét tệp này" size="large" imageMso="TrustCenter" onAction="Ribbon_ScanActiveWorkbook" screentip="Quét tệp này" supertip="Quét và tiêu diệt virus macro trên tệp Excel đang mở." />
          <button id="btnScanFolder" label="Quét thư mục..." size="large" imageMso="Folder" onAction="Ribbon_ScanFolderDialog" screentip="Quét thư mục..." supertip="Quét ngầm toàn bộ thư mục mà không làm gián đoạn Excel." />
          <button id="btnResume" label="Tiếp tục quét" size="large" imageMso="PlayMacro" onAction="Ribbon_ResumeScanDialog" screentip="Tiếp tục phiên quét" supertip="Tiếp tục phiên quét dở dang trước đó với tốc độ cao." />
        </group>
        <group id="grpTools" label="Hệ thống &amp; Tiện ích">
          <button id="btnFixCopyPaste" label="Khôi phục Copy/Paste" size="large" imageMso="Undo" onAction="Ribbon_RestoreClipboardAndUI" screentip="Khôi phục Copy/Paste &amp; Chuột phải" supertip="Khôi phục lại chức năng Copy/Paste, phím tắt Ctrl+C, Ctrl+V và menu chuột phải nếu bị virus làm hỏng hoặc khóa." />
          <button id="btnLog" label="Nhật ký (Log)" size="large" imageMso="ImportTextFile" onAction="Ribbon_OpenLogFolder" screentip="Nhật ký kiểm toán" supertip="Mở thư mục chứa nhật ký quét và diệt virus." />
          <button id="btnUpdate" label="Cập nhật LAN" size="large" imageMso="Synchronize" onAction="Ribbon_CheckForLanUpdates" screentip="Cập nhật mạng LAN" supertip="Kiểm tra và cập nhật phiên bản mới từ máy chủ LAN." />
          <button id="btnAbout" label="Thông tin" size="large" imageMso="Info" onAction="Ribbon_ShowAbout" screentip="Thông tin" supertip="Thông tin phiên bản và xuất xứ KangatangGuard." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"@
                $uiBytes = [System.Text.Encoding]::UTF8.GetBytes($customUiXml)
            }
            $uiEntry = $zip.GetEntry("customUI/customUI14.xml")
            if ($uiEntry) { $uiEntry.Delete() }
            $newUiEntry = $zip.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
            $uiStream = $newUiEntry.Open()
            $uiStream.Write($uiBytes, 0, $uiBytes.Length)
            $uiStream.Flush()
            $uiStream.Close()
            Write-Host "   -> [OK] Da tich hop Ribbon CustomUI XML (UTF-8 Native Byte Stream) thanh cong!" -ForegroundColor Green
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
    
    # --- Tao Standard Module: modKangatangShield (v3.10.0 Runtime Shield) ---
    Write-Host "   -> Tao Module: modKangatangShield..." -ForegroundColor Gray
    $modShield = $wb.VBProject.VBComponents.Add(1)
    $modShield.Name = "modKangatangShield"
    Inject-CleanVbaModule $modShield $modShieldCode
    
    # --- Tao Standard Module: modThreatCollector (v3.11.0 Threat Telemetry) ---
    if (-not [string]::IsNullOrWhiteSpace($modThreatCode)) {
        Write-Host "   -> Tao Module: modThreatCollector..." -ForegroundColor Gray
        $modThreat = $wb.VBProject.VBComponents.Add(1)
        $modThreat.Name = "modThreatCollector"
        Inject-CleanVbaModule $modThreat $modThreatCode
    }
    
    # --- Tao Standard Module: modLogger ---
    Write-Host "   -> Tao Module: modLogger..." -ForegroundColor Gray
    $modLog = $wb.VBProject.VBComponents.Add(1)
    $modLog.Name = "modLogger"
    Inject-CleanVbaModule $modLog $modLoggerCode
    
    # Xoa Sheet1 mac dinh (add-in khong can sheet)
    # Giu lai 1 sheet vi Excel bat buoc phai co it nhat 1
    
    # Luu file dang .xlam qua tep trung gian de khong bao gio xung dot voi Excel dang chay
    Write-Host "   -> Luu file .xlam..." -ForegroundColor Gray
    $tempBuildXlam = Join-Path $ScriptDir "KangatangGuard_build_tmp.xlam"
    if (Test-Path $tempBuildXlam) {
        Remove-Item -Path $tempBuildXlam -Force -ErrorAction SilentlyContinue
    }
    
    # Cau hinh thuoc tinh Add-in chuan
    try {
        $wb.BuiltinDocumentProperties("Title").Value = "KangatangGuard"
        $wb.BuiltinDocumentProperties("Comments").Value = "Đây là addin phòng chống virus kangatang - Phát triển bởi OhsungIT team !"
    } catch {}
    $wb.IsAddin = $true
    $wb.SaveAs($tempBuildXlam, 55)  # 55 = xlOpenXMLAddIn (.xlam)
    $wb.Close($false)
    
    # Tich hop Ribbon CustomUI XML (Icon HD 32x32)
    Add-RibbonCustomUI -targetXlam $tempBuildXlam
    
    # Dong bo ban sao xlam vao thu muc nguon repository (addin\KangatangGuard.xlam)
    $repoXlam = Join-Path $ScriptDir "KangatangGuard.xlam"
    try {
        if (Test-Path $repoXlam) { Set-ItemProperty -Path $repoXlam -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue }
        Copy-Item -Path $tempBuildXlam -Destination $repoXlam -Force
        Write-Host "   -> [OK] Da cap nhat ban sao repository: $repoXlam" -ForegroundColor Green
    } catch {}
    
    # Dong bo vao %APPDATA%\Microsoft\AddIns\KangatangGuard.xlam
    try {
        if (Test-Path $xlamPath) { Set-ItemProperty -Path $xlamPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue }
        Copy-Item -Path $tempBuildXlam -Destination $xlamPath -Force -ErrorAction Stop
        Set-ItemProperty -Path $xlamPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
        Write-Host "   -> [OK] Da luu vao AddIns: $xlamPath" -ForegroundColor Green
    } catch {
        Write-Host "   -> [OK] Da luu ban cap nhat tai repository (AddIns dang duoc su dung boi phien quet hien tai)." -ForegroundColor Cyan
    }
    Remove-Item -Path $tempBuildXlam -Force -ErrorAction SilentlyContinue

    # ==============================================================================
    # DANG KY EXCEL ADD-INS MANAGER & DON DEP XUNG DOT XLSTART (v3.16.0)
    # ==============================================================================
    Write-Host "`n[ADDIN-REGISTER] Dang ky KangatangGuard.xlam vao Excel Add-Ins Manager & don xung dot XLSTART..." -ForegroundColor Yellow
    $xlamXlStartPath = Join-Path $xlStartPath "KangatangGuard.xlam"
    if (Test-Path $xlamXlStartPath) {
        try { Set-ItemProperty -Path $xlamXlStartPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
        Remove-Item -Path $xlamXlStartPath -Force -ErrorAction SilentlyContinue
        Write-Host "   -> [OK] Da xoa ban sao trung lap trong XLSTART: $xlamXlStartPath" -ForegroundColor Green
    }
    
    # Dang ky khoa OPEN* trong Registry Options (giu nguyen cac Add-in khac nhu MyExcel.xlam)
    foreach ($ver in @("16.0", "15.0", "14.0")) {
        $excelKey = "HKCU:\Software\Microsoft\Office\$ver\Excel"
        if (-not (Test-Path $excelKey) -and $ver -ne "16.0") { continue }
        $optKey = "$excelKey\Options"
        if (-not (Test-Path $optKey)) {
            New-Item -Path $optKey -Force -ErrorAction SilentlyContinue | Out-Null
        }
        if (Test-Path $optKey) {
            $otherOpenValues = New-Object System.Collections.Generic.List[string]
            $oldOpenNames    = New-Object System.Collections.Generic.List[string]
            $props = (Get-ItemProperty -Path $optKey -ErrorAction SilentlyContinue).psobject.Properties
            foreach ($p in $props) {
                if ($p.Name -eq "OPEN" -or $p.Name -match "^OPEN(\d+)$") {
                    $oldOpenNames.Add($p.Name)
                    $val = [string]$p.Value
                    if ($val -and ($val -notlike "*KangatangGuard*")) {
                        $otherOpenValues.Add($val)
                    }
                }
            }
            # Xoa cac khoa OPEN* cu va ghi lai tuan tu (OPEN, OPEN1, OPEN2...) de khong bi đứt quãng
            foreach ($oldName in $oldOpenNames) {
                Remove-ItemProperty -Path $optKey -Name $oldName -Force -ErrorAction SilentlyContinue
            }
            $otherOpenValues.Add('"KangatangGuard.xlam"')
            for ($idx = 0; $idx -lt $otherOpenValues.Count; $idx++) {
                $slotName = if ($idx -eq 0) { "OPEN" } else { "OPEN$idx" }
                Set-ItemProperty -Path $optKey -Name $slotName -Value $otherOpenValues[$idx] -Type String -Force -ErrorAction SilentlyContinue
            }
            Write-Host "   -> [OK] Da dang ky KangatangGuard.xlam vao $optKey (bao toan $($otherOpenValues.Count - 1) Add-in khac)" -ForegroundColor Green
        }
        
        # Chi xoa bo Add-in Manager entry tro den UNC mang cu (khong xoa muc cuc bo)
        $mgrKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Add-in Manager"
        if (Test-Path $mgrKey) {
            $mProps = (Get-ItemProperty -Path $mgrKey -ErrorAction SilentlyContinue).psobject.Properties
            foreach ($mp in $mProps) {
                if ($mp.Name -like "\\*" -and ($mp.Name -like "*KangatangGuard*" -or $mp.Name -like "*addin_kangatang*")) {
                    Remove-ItemProperty -Path $mgrKey -Name $mp.Name -Force -ErrorAction SilentlyContinue
                    Write-Host "   -> [OK] Da go bo UNC Add-in Manager: $mgrKey\$($mp.Name)" -ForegroundColor Green
                }
            }
        }
        
        # Don dep AddInLoadTimes va DisabledItems
        $altKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\AddInLoadTimes"
        if (Test-Path $altKey) {
            Remove-ItemProperty -Path $altKey -Name "KangatangGuard.xlam" -Force -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $altKey -Name "KangatangGuard" -Force -ErrorAction SilentlyContinue
        }
        $disKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
        if (Test-Path $disKey) {
            $disProps = (Get-ItemProperty -Path $disKey -ErrorAction SilentlyContinue).psobject.Properties
            foreach ($dp in $disProps) {
                if ($dp.Name -notlike "PS*") {
                    $bytes = [byte[]]$dp.Value
                    if ($bytes) {
                        $str = [System.Text.Encoding]::Unicode.GetString($bytes)
                        if ($str -like "*KangatangGuard*") {
                            Remove-ItemProperty -Path $disKey -Name $dp.Name -Force -ErrorAction SilentlyContinue
                            Write-Host "   -> [OK] Da go bo khoi DisabledItems: $disKey\$($dp.Name)" -ForegroundColor Green
                        }
                    }
                }
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
# BUOC 3: TAO THU MUC LOG, SESSIONS VA DON SACH LEGACY PS1 TRONG APPDATA (v3.17.0)
# ==============================================================================
Write-Host "`n[3/3] Dang thiet lap thu muc Log, Sessions va don sach legacy PS1 (Antivirus-Safe)..." -ForegroundColor Yellow
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

# v3.17.3: Sao chep Kangatang_FolderScanner.ps1 (0-IOC, 0-P/Invoke) vao thu muc tin cay Microsoft\AddIns
# va don sach legacy .ps1 trong APPDATA\KangatangGuard de tranh Seqrite/Quick Heal nhan nham!
$srcFolderScanner = Join-Path $ScriptDir "Kangatang_FolderScanner.ps1"
$dstAddInsScanner = Join-Path $env:APPDATA "Microsoft\AddIns\Kangatang_FolderScanner.ps1"
if (Test-Path $srcFolderScanner) {
    try { Set-ItemProperty -Path $dstAddInsScanner -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
    Copy-Item -Path $srcFolderScanner -Destination $dstAddInsScanner -Force -ErrorAction SilentlyContinue
    Write-Host "   -> [OK] Da cai dat Trinh quet thu muc ngoai tien trinh: $dstAddInsScanner" -ForegroundColor Green
}

foreach ($oldPs1 in @(
    (Join-Path $logDir "Kangatang_FolderScanner.ps1"),
    (Join-Path $logDir "staged_update\Kangatang_FolderScanner.ps1"),
    (Join-Path $xlStartPath "Kangatang_FolderScanner.ps1")
)) {
    if (Test-Path $oldPs1) {
        try { Set-ItemProperty -Path $oldPs1 -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
        Remove-Item -Path $oldPs1 -Force -ErrorAction SilentlyContinue
        Write-Host "   -> [OK] Da go bo legacy script khoi AppData: $oldPs1" -ForegroundColor Green
    }
}

$guardRegKey = "HKCU:\Software\KangatangGuard"
if (-not (Test-Path $guardRegKey)) {
    New-Item -Path $guardRegKey -Force | Out-Null
}
Remove-ItemProperty -Path $guardRegKey -Name "ScannerScript" -Force -ErrorAction SilentlyContinue
Set-ItemProperty -Path $guardRegKey -Name "InstalledVersion" -Value $InstallerVersion -Force
Set-ItemProperty -Path $guardRegKey -Name "UpdateSource" -Value "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang" -Force
Write-Host "   -> [OK] Da dang ky InstalledVersion: v$InstallerVersion (100% Native VBA Scanner)" -ForegroundColor Green
Write-Host "   -> [OK] Da dang ky UpdateSource: \\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang" -ForegroundColor Green

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

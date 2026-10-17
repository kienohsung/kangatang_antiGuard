# ==============================================================================
# Kangatang_FolderScanner.ps1
# Trinh khoi chay Quet Thu Muc qua Engine Native VBA cua KangatangGuard.xlam
# Phien ban: v3.17.1 (Unified Version Sync & Antivirus-Safe Pure COM Wrapper)
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

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

# Dong bo phien ban dong tu version.json hoac Registry InstalledVersion (khong bao gio lech phien ban voi Add-in)
$AppVersion = "3.17.1"
$verJsonPath = Join-Path $ScriptDir "version.json"
if (Test-Path $verJsonPath) {
    try {
        $vMeta = Get-Content -Path $verJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($vMeta.version) { $AppVersion = $vMeta.version }
    } catch {}
} elseif (Test-Path "HKCU:\Software\KangatangGuard") {
    try {
        $regVer = (Get-ItemProperty -Path "HKCU:\Software\KangatangGuard" -Name "InstalledVersion" -ErrorAction SilentlyContinue).InstalledVersion
        if ($regVer) { $AppVersion = $regVer }
    } catch {}
}

$Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - Trinh Quet Thu Muc (Native VBA Engine)"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD v$AppVersion - NATIVE VBA FOLDER SCANNER           " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Dam bao AccessVBOM = 1 trong Registry HKCU
foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
    $secPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $secPath) {
        Set-ItemProperty -Path $secPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    }
}

# 2. Xac dinh thu muc can quet neu chua truyen tham so
$logDir = Join-Path $env:APPDATA "KangatangGuard"
$lastSessionFile = Join-Path $logDir "last_session.json"

if ($Resume -and (-not $TargetFolder) -and (Test-Path $lastSessionFile)) {
    try {
        $meta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($meta.TargetFolder) { $TargetFolder = $meta.TargetFolder }
    } catch {}
}

if (-not $TargetFolder) {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "KangatangGuard v$AppVersion - Chon thu muc can quet:"
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TargetFolder = $dlg.SelectedPath
    } else {
        Write-Host "`n[HUY] Nguoi dung da huy chon thu muc." -ForegroundColor Yellow
        exit 0
    }
}

# 3. Xac dinh duong dan KangatangGuard.xlam
$xlamCandidates = @(
    (Join-Path $env:APPDATA "Microsoft\AddIns\KangatangGuard.xlam"),
    (Join-Path $ScriptDir "KangatangGuard.xlam"),
    "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam"
)
$xlamPath = $null
foreach ($c in $xlamCandidates) {
    if ($c -and (Test-Path $c)) {
        $xlamPath = $c
        break
    }
}

if (-not $xlamPath) {
    Write-Host "`n[LOI] Khong tim thay file KangatangGuard.xlam!" -ForegroundColor Red
    exit 1
}

Write-Host "`n[1/2] Dang khoi tao Engine KangatangGuard.xlam (v$AppVersion)..." -ForegroundColor Yellow
Write-Host "   -> Add-in : $xlamPath" -ForegroundColor Gray
Write-Host "   -> Thu muc: $TargetFolder" -ForegroundColor Gray

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wbAddin = $excel.Workbooks.Open($xlamPath, 0, $true)
    Write-Host "`n[2/2] Dang thuc thi RunNativeFolderScan (v$AppVersion)..." -ForegroundColor Yellow
    $excel.Run("'$($wbAddin.Name)'!RunNativeFolderScan", [string]$TargetFolder, [bool]$Resume, $true)
    if (Test-Path $lastSessionFile) {
        $resMeta = Get-Content -Path $lastSessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        Write-Host "`n[HOAN TAT - v$AppVersion] Tong so tep: $($resMeta.TotalFiles) | Da quet: $($resMeta.TotalScanned) | Da lam sach: $($resMeta.TotalCleaned)" -ForegroundColor Green
    }
    $wbAddin.Close($false)
} catch {
    Write-Host "`n[LOI] Loi khi quet thu muc: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

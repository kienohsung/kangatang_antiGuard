# ==============================================================================
# Kangatang_Standalone_Scanner.ps1
# Trinh Quet & Diet Virus Macro Excel Doc Lap (Standalone Wrapper v3.20.2)
# Phien ban: v3.20.2 (Real-Time Streaming Console & 100% Antivirus-Safe)
# ==============================================================================

param (
    [Parameter(Mandatory=$false)]
    [string]$TargetFolder = "",
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    [Parameter(Mandatory=$false)]
    [switch]$CleanSystemOnly
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

$AppVersion = "3.20.2"
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

$Host.UI.RawUI.WindowTitle = "KangatangGuard v$AppVersion - Trinh Diet Virus Excel Doc Lap (Standalone)"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   KANGATANG GUARD v$AppVersion - STANDALONE SCANNER (ANTIVIRUS-SAFE) " -ForegroundColor Cyan
Write-Host "   Day la addin phong chong virus kangatang - Phat trien boi OhsungIT team !" -ForegroundColor White
Write-Host "======================================================================" -ForegroundColor Cyan

foreach ($ver in @("16.0", "15.0", "14.0", "12.0")) {
    $secPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $secPath) {
        Set-ItemProperty -Path $secPath -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    }
}

$xlamCandidates = @(
    (Join-Path $ScriptDir "KangatangGuard.xlam"),
    (Join-Path $env:APPDATA "Microsoft\AddIns\KangatangGuard.xlam"),
    "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\KangatangGuard.xlam"
)
$xlamPath = $null
foreach ($c in $xlamCandidates) {
    if ($c -and (Test-Path -LiteralPath $c)) {
        $xlamPath = $c
        break
    }
}

if (-not $xlamPath) {
    Write-Host "`n[LOI] Khong tim thay file KangatangGuard.xlam!" -ForegroundColor Red
    exit 1
}

Write-Host "`n[BUOC 1] Dang quet & lam sach thu muc khoi dong Excel (XLSTART & Startup Vaccine)..." -ForegroundColor Yellow
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wbAddin = $excel.Workbooks.Open($xlamPath, 0, $true)
    $excel.Run("'$($wbAddin.Name)'!ShieldStartupSweep") | Out-Null
    $wbAddin.Close($false)
    Write-Host "   -> [OK] Da bao ve thu muc khoi dong Excel thanh cong!" -ForegroundColor Green
} catch {
    Write-Host "   -> [CANH BAO] Loi khi quet he thong khoi dong: $($_.Exception.Message)" -ForegroundColor Yellow
} finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

if ($CleanSystemOnly) {
    exit 0
}

$folderScannerPath = Join-Path $ScriptDir "Kangatang_FolderScanner.ps1"
if (-not (Test-Path -LiteralPath $folderScannerPath)) {
    $folderScannerPath = Join-Path $env:APPDATA "Microsoft\AddIns\Kangatang_FolderScanner.ps1"
}

if (Test-Path -LiteralPath $folderScannerPath) {
    if ($Resume) {
        & $folderScannerPath -TargetFolder $TargetFolder -Resume
    } else {
        & $folderScannerPath -TargetFolder $TargetFolder
    }
}

'==============================================================================
' KangatangGuard - Excel Add-in Diet Virus Macro Kangatang
' Phien ban: v3.12.0 (Deep Copy/Paste Healing, Auto-Restore on Safe Files & Unrestricted Selection)
' Mo ta: Tu dong quet va tieu diet virus macro Kangatang/Laroux/mypersonnel
'         ngay khi mo file Excel. Ho tro Trung tam Phan phoi May chu Tep LAN:
'         \\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang
'         va co che Nhan ban Kep (Dual-Mirror) luu tru cuc bo tren may 223.176!
'
' File nay chua toan bo ma VBA duoc phan tach theo section markers.
' PowerShell Installer se doc va inject tung phan vao dung module/class.
'==============================================================================

'### SECTION: ThisWorkbook ###
'--- Code nay se duoc inject vao ThisWorkbook cua file .xlam ---

Option Explicit

Private Sub Workbook_Open()
    ' v3.8.6: Core-Only Startup - Chi nap code cot loi, loai bo moi tac vu WMI/Network chan UI
    Call InitializeGuard
    ' Tri hoan kiem tra cap nhat LAN 45 giay de Excel nhan roi hoan toan va rate-limit 24h
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:45"), "'" & ThisWorkbook.Name & "'!CheckForLanUpdatesSilent"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call TerminateGuard
End Sub

'### END_SECTION: ThisWorkbook ###


'### SECTION: clsAppEvents ###
'--- Class Module: Bat su kien Application-level ---

Option Explicit

Public WithEvents xlApp As Excel.Application

Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    If bIsFolderScanning Then Exit Sub
    Call ScanWorkbook(Wb)
    ' v3.10.0: Virus co the chen sheet ngay sau khi mo -> kiem tra lai tre 1 giay
    Call ScheduleShieldCheck
    On Error GoTo 0
End Sub

Private Sub xlApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' v3.10.0: LUON kiem tra nhanh theo ten (re, < 5ms) truoc khi luu - bit lo hong ScanCache 5 phut.
    ' Neu nhiem: lam sach trong bo nho, KHONG goi wb.Save (tranh de quy) - lan luu hien tai se ghi ban sach.
    On Error Resume Next
    If bIsFolderScanning Then Exit Sub
    Dim reason As String
    If QuickCheckWorkbook(Wb, reason) Then
        WriteLog "[BEFORESAVE_HIT] " & Wb.FullName & " | " & reason
        Call ScanWorkbook(Wb, True)
    ElseIf IsScanCacheExpired(Wb) Then
        Call ScanWorkbook(Wb, True)
    End If
    On Error GoTo 0
End Sub

Private Sub xlApp_NewWorkbook(ByVal Wb As Workbook)
    On Error Resume Next
    If bIsFolderScanning Then Exit Sub
    Call ScanWorkbook(Wb)
    On Error GoTo 0
End Sub

Private Sub xlApp_SheetActivate(ByVal Sh As Object)
    ' v3.10.0: Virus Kangatang lay qua Application.OnSheetActivate -> kiem tra tre (debounce) sau moi lan doi sheet
    On Error Resume Next
    If bIsFolderScanning Then Exit Sub
    Call ScheduleShieldCheck
    On Error GoTo 0
End Sub

Private Sub xlApp_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    If bIsFolderScanning Then Exit Sub
    Call ScheduleShieldCheck
    Call RestoreExcelClipboardAndUI(Wb)
    On Error GoTo 0
End Sub

'### END_SECTION: clsAppEvents ###


'### SECTION: modKangatangScanner ###
'--- Standard Module: Logic quet chinh, backup, lam sach va dieu phoi worker ---

Option Explicit

' --- Windows API de hien thi Unicode tieng Viet co dau sac net ---
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' --- Bien toan cuc ---
Private oAppEvents As clsAppEvents
Public bIsFolderScanning As Boolean

' Scan Cache chong lag: key=UCase(FullName), value=Date
Private dicScanCache As Object
' v3.10.0: Chong bat trung nhieu cua so quet ngam cho cung 1 thu muc trong 1 phien
Private dicLaunchedFolders As Object

Public Const CURRENT_VERSION As String = "3.12.2"
Public Const DEFAULT_HUB_PRIMARY As String = "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang"
Public Const DEFAULT_HUB_BACKUP  As String = "\\192.168.223.176\KangatangGuard_Hub"
Public Const CENTRAL_BACKUP_HUB  As String = "\\192.168.223.7\file_shared\vietnam\z. ETC\2. Virus backupfile - DO NOT OPEN IT"

Public Const ADDIN_NAME As String = "KangatangGuard.xlam"
Private Const TOOLBAR_NAME As String = "KangatangGuard"
Private Const BACKUP_FOLDER_NAME As String = "_Backup_Kangatang"
Public Const LOG_SUBFOLDER As String = "KangatangGuard"
Private Const SCAN_CACHE_MINUTES As Long = 5
Private Const RELAUNCH_FOLDER_MINUTES As Long = 30
Private Const UPDATE_CHECK_HOURS As Long = 6

' --- v3.10.0: Chu ky virus da chuyen sang modKangatangShield (IsVirusName / CodeHasIOC) ---
' Loai bo tu khoa chuoi con "Kanga" (gay xoa nham sheet/module hop le nhu "Kangaroo").

' ===========================================================================
' HAM GIAI MA CHUOI UNICODE (v3.4.0)
' Chuyen doi cac ky tu \uXXXX thanh ChrW() giup ma VBA 100% doc lap voi CodePage ANSI
' ===========================================================================
Public Function Uni(ByVal txt As String) As String
    Dim pos As Long, hexCode As String
    Dim res As String
    res = txt
    pos = InStr(res, "\u")
    Do While pos > 0
        If pos + 5 <= Len(res) Then
            hexCode = Mid(res, pos + 2, 4)
            res = Left(res, pos - 1) & ChrW(CLng("&H" & hexCode)) & Mid(res, pos + 6)
        Else
            Exit Do
        End If
        pos = InStr(pos + 1, res, "\u")
    Loop
    Uni = res
End Function

' ===========================================================================
' HAM HIEN THI THONG BAO TIENG VIET CO DAU (v3.4.0)
' Goi truc tiep API MessageBoxW cua Windows de khong bao gio bi loi font dau ?
' ===========================================================================
Public Function MsgBoxW(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    Dim h As LongPtr
    On Error Resume Next
    h = Application.Hwnd
    On Error GoTo 0
    
    If Len(title) = 0 Then
        title = Uni("KangatangGuard v" & CURRENT_VERSION)
    End If
    
    MsgBoxW = MessageBoxW(h, StrPtr(prompt), StrPtr(title), buttons)
End Function

' ===========================================================================
' SCAN CACHE: Chong lag khi BeforeSave/AutoSave goi lien tuc
' ===========================================================================
Public Function IsScanCacheExpired(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    IsScanCacheExpired = True
    
    If dicScanCache Is Nothing Then Exit Function
    
    Dim sKey As String
    sKey = UCase(wb.FullName)
    If Len(sKey) = 0 Then sKey = UCase(wb.Name)
    
    If dicScanCache.Exists(sKey) Then
        Dim lastScan As Date
        lastScan = CDate(dicScanCache(sKey))
        If DateDiff("s", lastScan, Now) < (SCAN_CACHE_MINUTES * 60) Then
            IsScanCacheExpired = False
        End If
    End If
    On Error GoTo 0
End Function

Private Sub UpdateScanCache(ByVal wb As Workbook)
    On Error Resume Next
    If dicScanCache Is Nothing Then
        Set dicScanCache = CreateObject("Scripting.Dictionary")
    End If
    
    Dim sKey As String
    sKey = UCase(wb.FullName)
    If Len(sKey) = 0 Then sKey = UCase(wb.Name)
    
    dicScanCache(sKey) = Now
    On Error GoTo 0
End Sub

' ===========================================================================
'' Ham khoi tao va huy event handler (v3.8.6: Core-Only Zero-Delay Startup)
' ===========================================================================
Public Sub InitializeGuard()
    On Error Resume Next
    ' 0. Khoi phuc clipboard, phim tat va menu chuot phai neu bi virus chan (v3.11.0)
    Call RestoreExcelClipboardAndUI
    
    ' 1. Khoi tao Scan Cache trong bo nho RAM (< 1ms)
    Set dicScanCache = CreateObject("Scripting.Dictionary")
    Set dicLaunchedFolders = CreateObject("Scripting.Dictionary")
    bIsFolderScanning = False
    
    ' 2. Dang ky lang nghe su kien WorkbookOpen/BeforeSave cua ung dung (< 1ms)
    Set oAppEvents = New clsAppEvents
    Set oAppEvents.xlApp = Application
    
    ' 2b. v3.10.0: Quet quet toan bo workbook DA MO san (ke ca an: mypersonnel1.xls, PERSONAL.XLSB, add-in)
    ' Tri hoan 3 giay de khong lam cham khoi dong Excel (khong I/O mang)
    Application.OnTime Now + TimeSerial(0, 0, 3), "'" & ThisWorkbook.Name & "'!ShieldStartupSweep"
    
    ' 3. Don sach moi thanh cong cu CommandBars cu trong cache (< 1ms)
    Call RemoveMenu
    On Error GoTo 0
End Sub

Public Sub TerminateGuard()
    On Error Resume Next
    Call RemoveMenu
    Set oAppEvents.xlApp = Nothing
    Set oAppEvents = Nothing
    Set dicScanCache = Nothing
    bIsFolderScanning = False
    On Error GoTo 0
End Sub

' v3.10.0: Tai kich hoat hook su kien neu project add-in vua bi reset
' (xay ra khi xoa module-sheet Excel 5 tu VBA - xem Two-phase purge trong modKangatangShield)
Public Sub EnsureGuardAlive()
    On Error Resume Next
    If dicScanCache Is Nothing Then Set dicScanCache = CreateObject("Scripting.Dictionary")
    If dicLaunchedFolders Is Nothing Then Set dicLaunchedFolders = CreateObject("Scripting.Dictionary")
    If oAppEvents Is Nothing Then
        Set oAppEvents = New clsAppEvents
        Set oAppEvents.xlApp = Application
        WriteLog "[GUARD_REARMED] Application event hooks re-created after VBA project reset."
    End If
End Sub

' ===========================================================================
' v3.8.8: CHUAN HOA GIAO DIEN TAB "Kangatang Guard" TREN MODERN RIBBON
' Toan bo giao dien da duoc chuyen sang 1 Tab duy nhat tren Ribbon chinh.
' Ham CreateMenu duoc giu lai de don sach moi thanh cong cu cu tren tab Add-ins.
' ===========================================================================
Public Sub CreateMenu()
    On Error Resume Next
    Call RemoveMenu
    On Error GoTo 0
End Sub

Public Sub RemoveMenu()
    On Error Resume Next
    Application.CommandBars(TOOLBAR_NAME).Delete
    
    ' Don sach ca control cu tren Worksheet Menu Bar neu co
    Dim ctrl As Object
    For Each ctrl In Application.CommandBars("Worksheet Menu Bar").Controls
        If ctrl.Tag = "KangatangGuardMenu" Or ctrl.Tag Like "KG_*" Then
            ctrl.Delete
        End If
    Next ctrl
    On Error GoTo 0
End Sub

' ===========================================================================
' RIBBON UI CALLBACKS (v3.8.8)
' Dành riêng cho Tab "Kangatang Guard" trên Modern Office Ribbon
' ===========================================================================
Public Sub Ribbon_ScanActiveWorkbook(ByVal control As Object)
    Call ScanActiveWorkbook
End Sub

Public Sub Ribbon_ScanFolderDialog(ByVal control As Object)
    Call ScanFolderDialog
End Sub

Public Sub Ribbon_ResumeScanDialog(ByVal control As Object)
    Call ResumeScanDialog
End Sub

Public Sub Ribbon_OpenLogFolder(ByVal control As Object)
    Call OpenLogFolder
End Sub

Public Sub Ribbon_CheckForLanUpdates(ByVal control As Object)
    Call CheckForLanUpdatesManual
End Sub

Public Sub Ribbon_ShowAbout(ByVal control As Object)
    Call ShowAbout
End Sub

Public Sub Ribbon_RestoreClipboardAndUI(control As IRibbonControl)
    Call RestoreExcelClipboardAndUI
    If Application.Visible Then
        MsgBoxW Uni("\u0110\u00e3 kh\u00f4i ph\u1ee5c th\u00e0nh c\u00f4ng ch\u1ee9c n\u0103ng Copy/Paste, ph\u00edm t\u1eaft v\u00e0 menu chu\u1ed9t ph\u1ea3i Excel!"), _
            vbInformation, Uni("KangatangGuard - Kh\u00f4i ph\u1ee5c Copy/Paste")
    End If
End Sub

' ===========================================================================
' Ham kiem tra whitelist (v3.10.0: CHI bo qua chinh add-in KangatangGuard)
' Truoc day bo qua MOI .xlam va PERSONAL.XLSB -> diem mu dung noi virus tru an.
' ===========================================================================
Private Function IsWhitelisted(ByVal wb As Workbook) As Boolean
    IsWhitelisted = False
    On Error Resume Next
    
    If wb Is ThisWorkbook Then
        IsWhitelisted = True
        Exit Function
    End If
    
    If UCase(wb.Name) = UCase(ADDIN_NAME) Then
        IsWhitelisted = True
        Exit Function
    End If
    
    ' VACCINE v3.5.4: Bo qua cac tep khong phai file lam viec Excel thuan tuy
    Dim dotPos As Long
    dotPos = InStrRev(wb.Name, ".")
    If dotPos > 0 Then
        Dim ext As String
        ext = LCase(Mid(wb.Name, dotPos))
        If ext = ".ps1" Or ext = ".bat" Or ext = ".txt" Or ext = ".log" Or ext = ".csv" Then
            IsWhitelisted = True
            Exit Function
        End If
    End If
    
    On Error GoTo 0
End Function

' ===========================================================================
' HAM KIEM TRA MA DOC (INSPECTION ENGINE v3.10.0)
' 1. Sheet (ke ca VeryHidden / Module-sheet Excel 5 do virus chep vao): IsVirusName
' 2. Named Range: ten hoac RefersTo tro toi sheet virus
' 3. VBA: ten component + TOAN BO noi dung code (CodeHasIOC)
' Khong truy cap duoc VBProject -> KHONG coi la sach im lang: log [NO_VBOM] + canh bao 1 lan/ngay
' ===========================================================================
Public Function CheckWorkbookInfection(ByVal wb As Workbook, ByRef detailMsg As String) As Boolean
    On Error GoTo CheckError
    CheckWorkbookInfection = False
    detailMsg = ""
    
    If IsWhitelisted(wb) Then Exit Function
    
    Dim i As Long
    
    ' 1. Sheet
    Dim sht As Object
    For i = wb.Sheets.Count To 1 Step -1
        Set sht = wb.Sheets(i)
        If IsVirusName(sht.Name) Then
            detailMsg = detailMsg & Uni("  - Sheet \u1ea9n \u0111\u1ed9c h\u1ea1i: ") & sht.Name & vbCrLf
            CheckWorkbookInfection = True
        End If
    Next i
    
    ' 2. Named Ranges
    Dim nm As Object, refStr As String
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        refStr = ""
        On Error Resume Next
        refStr = nm.RefersTo
        Err.Clear
        On Error GoTo CheckError
        If IsVirusName(nm.Name) Or IsVirusReference(refStr) Then
            detailMsg = detailMsg & Uni("  - Named Range \u0111\u1ed9c h\u1ea1i: ") & nm.Name & vbCrLf
            CheckWorkbookInfection = True
        End If
    Next i
    
    ' 2b. (v3.10.0) Module-sheet Excel 5 con ton tai -> KHONG duoc cham wb.VBProject:
    '     da tai hien Excel crash (RPC_E_SERVERFAULT) / treo khi file mo o che do tat macro.
    If LegacyModuleSheetCount(wb) <> 0 Then
        WriteLog "[VBP_SKIPPED] " & wb.FullName & " - legacy Excel 5 module sheet present, VBProject not inspected (crash guard)."
        Exit Function
    End If
    
    ' 3. VBA Project
    Dim vbProj As Object
    On Error Resume Next
    Set vbProj = wb.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        Err.Clear
        On Error GoTo CheckError
        Call ReportVbomBlocked(wb)
        Exit Function
    End If
    Dim prot As Long
    prot = vbProj.Protection
    If Err.Number <> 0 Then
        prot = 0
        Err.Clear
    End If
    On Error GoTo CheckError
    
    If prot = 1 Then
        WriteLog "[VBA_LOCKED] " & wb.FullName & " - VBA project has a password, only sheets/names were checked."
        Exit Function
    End If
    
    Dim comp As Object, codeContent As String, nLines As Long
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents.Item(i)
        If IsVirusName(comp.Name) Then
            detailMsg = detailMsg & Uni("  - Module \u0111\u1ed9c h\u1ea1i: ") & comp.Name & vbCrLf
            CheckWorkbookInfection = True
        Else
            codeContent = ""
            nLines = 0
            On Error Resume Next
            nLines = comp.CodeModule.CountOfLines
            If nLines > 0 Then codeContent = comp.CodeModule.Lines(1, nLines)
            Err.Clear
            On Error GoTo CheckError
            If CodeHasIOC(codeContent) Then
                detailMsg = detailMsg & Uni("  - M\u00e3 \u0111\u1ed9c trong: ") & comp.Name & vbCrLf
                CheckWorkbookInfection = True
            End If
        End If
    Next i
    
    Exit Function
    
CheckError:
    WriteLog "[ERROR] CheckWorkbookInfection failed for " & wb.Name & ": " & Err.Description
    On Error GoTo 0
End Function

' ===========================================================================
' HAM QUET CHINH (v3.10.0)
' bFromBeforeSave = True: dang o trong su kien BeforeSave -> lam sach trong bo nho,
' KHONG goi wb.Save / KHONG dong workbook (tranh de quy va crash), lan luu hien tai se ghi ban sach.
' ===========================================================================
Public Sub ScanWorkbook(ByVal wb As Workbook, Optional ByVal bFromBeforeSave As Boolean = False)
    On Error GoTo ScanError
    
    If IsWhitelisted(wb) Then Exit Sub
    ' v3.10.0: dang cho Two-phase purge cho chinh workbook nay -> khong xu ly lai (tranh backup/OnTime trung)
    If IsPurgePending(wb) Then Exit Sub
    
    ' Uu tien tuyet doi (v3.11.0): Neu ten workbook la nguon virus (mypersonnel*, kangatang*, mypersonel*)
    ' Dong va cach ly ngay lap tuc - bat ke co chua code macro hoat dong hay khong,
    ' khong phu thuoc vao quyen AccessVBOM hay CheckWorkbookInfection!
    If IsVirusName(wb.Name) Then
        WriteLog "[VIRUS_CARRIER_DETECTED] Immediate quarantine for virus carrier: " & wb.FullName
        Call NeutralizeVirusHooks
        If bShieldContext Then
            Call QuarantineLoadedStartupWorkbook(wb)
        Else
            Call ScheduleShieldCheck
        End If
        Exit Sub
    End If
    
    Dim virusFound As Boolean
    Dim detailMsg As String
    virusFound = CheckWorkbookInfection(wb, detailMsg)
    
    UpdateScanCache wb
    
    If Not virusFound Then
        WriteLog "[SAFE] " & wb.FullName
        Call RestoreExcelClipboardAndUI(wb)
        Exit Sub
    End If
    
    WriteLog "[DETECTED] " & wb.FullName & vbCrLf & detailMsg
    
    ' Uu tien 1: Go hook su kien cua virus dang chay trong RAM (chan lay tiep sang file khac)
    Call NeutralizeVirusHooks
    
    If wb.ReadOnly Then
        ' Lam sach trong bo nho (khong ghi de tep) de chan lay lan trong phien lam viec
        Call CleanInfectedWorkbook(wb, False)
        WriteLog "[READONLY_DETECTED] " & wb.FullName & " is read-only, neutralized in memory only."
        MsgBoxW Uni("C\u1ea2NH B\u00c1O: PH\u00c1T HI\u1ec6N VIRUS KANGATANG TRONG T\u1ec6P CH\u1ec8 \u0110\u1eccC (READ-ONLY)!" & vbCrLf & vbCrLf & _
                    "T\u1ec7p: ") & wb.Name & vbCrLf & vbCrLf & _
                    Uni("Chi ti\u1ebft:") & vbCrLf & detailMsg & vbCrLf & _
                    Uni("\u0110\u00e3 v\u00f4 hi\u1ec7u h\u00f3a m\u00e3 \u0111\u1ed9c trong b\u1ed9 nh\u1edb (kh\u00f4ng ghi \u0111\u00e8 t\u1ec7p)." & vbCrLf & _
                    "L\u01b0u \u00fd: T\u1ec7p \u0111ang \u1edf ch\u1ebf \u0111\u1ed9 Ch\u1ec8 \u0111\u1ecdc (Read-Only) n\u00ean kh\u00f4ng th\u1ec3 t\u1ef1 \u0111\u1ed9ng ghi \u0111\u00e8." & vbCrLf & _
                    "Vui l\u00f2ng m\u1edf kh\u00f3a t\u1ec7p ho\u1eb7c ch\u1ecdn 'Save As' sang b\u1ea3n sao m\u1edbi."), _
                vbCritical, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - C\u1ea3nh b\u00e1o")
        Exit Sub
    End If
    
    Dim backupOK As Boolean
    backupOK = BackupBeforeClean(wb)
    If Not backupOK And Len(wb.Path) = 0 Then
        ' Workbook moi chua tung luu: khong co ban goc tren dia de sao luu.
        ' Chi xoa thanh phan khop chu ky virus -> an toan du lieu nguoi dung.
        WriteLog "[BACKUP_SKIPPED_UNSAVED] " & wb.Name
        backupOK = True
    End If
    
    If backupOK Then
        Dim cleanResult As Boolean
        bPurgeDeferred = False
        cleanResult = CleanInfectedWorkbook(wb, Not bFromBeforeSave)
        
        ' v3.10.0: Module-sheet Excel 5 -> Two-phase purge dang chay (OnTime), Phase 2 se thong bao ket qua
        If bPurgeDeferred Then
            bPurgeDeferred = False
            Exit Sub
        End If
        
        If cleanResult Then
            WriteLog "[CLEANED] " & wb.FullName
            
            Dim bLaunch As Boolean
            bLaunch = (Not bFromBeforeSave) And (Len(wb.Path) > 0) And (Not IsInStartupFolder(wb.FullName))
            
            Dim okMsg As String
            okMsg = Uni("C\u1ea2NH B\u00c1O: PH\u00c1T HI\u1ec6N VIRUS KANGATANG!" & vbCrLf & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & vbCrLf & _
                    Uni("Chi ti\u1ebft:") & vbCrLf & detailMsg & vbCrLf & _
                    Uni("Tr\u1ea1ng th\u00e1i: \u0110\u00c3 TI\u00caU DI\u1ec6T TH\u00c0NH C\u00d4NG!" & vbCrLf & _
                        "B\u1ea3n sao l\u01b0u g\u1ed1c \u0111\u00e3 \u0111\u01b0\u1ee3c c\u00e1ch ly an to\u00e0n v\u1ec1 Kho M\u00e1y ch\u1ee7 LAN:" & vbCrLf & _
                        "\\192.168.223.7\file_shared\vietnam\z. ETC\2. Virus backupfile - DO NOT OPEN IT")
            If bLaunch Then
                okMsg = okMsg & vbCrLf & vbCrLf & _
                        Uni("Ti\u1ebfn tr\u00ecnh qu\u00e9t ng\u1ea7m \u0111\u1ed9c l\u1eadp s\u1ebd T\u1ef0 \u0110\u1ed8NG QU\u00c9T TO\u00c0N B\u1ed8 TH\u01af M\u1ee4C ch\u1ee9a t\u1ec7p n\u00e0y m\u00e0 kh\u00f4ng l\u00e0m gi\u00e1n \u0111o\u1ea1n Excel c\u1ee7a b\u1ea1n!")
            End If
            MsgBoxW okMsg, vbExclamation, Uni("KangatangGuard - Ti\u00eau di\u1ec7t th\u00e0nh c\u00f4ng")
            
            If bLaunch Then Call LaunchBackgroundScanner(wb.Path)
        Else
            WriteLog "[ERROR] Clean failed: " & wb.FullName
            MsgBoxW Uni("C\u1ea2NH B\u00c1O: Ph\u00e1t hi\u1ec7n virus nh\u01b0ng KH\u00d4NG TH\u1ec2 l\u00e0m s\u1ea1ch ho\u00e0n to\u00e0n t\u1ef1 \u0111\u1ed9ng!" & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & _
                        Uni("Chi ti\u1ebft l\u1ed7i \u0111\u00e3 ghi v\u00e0o nh\u1eadt k\u00fd. Vui l\u00f2ng ch\u1ea1y Chay_Diet_Virus_Ngoai.bat ho\u1eb7c li\u00ean h\u1ec7 IT."), _
                    vbCritical, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - L\u1ed7i")
        End If
    Else
        WriteLog "[ERROR] Backup failed, file not modified: " & wb.FullName
        MsgBoxW Uni("C\u1ea2NH B\u00c1O: Ph\u00e1t hi\u1ec7n virus nh\u01b0ng KH\u00d4NG TH\u1ec2 t\u1ea1o b\u1ea3n sao l\u01b0u (Backup)!" & vbCrLf & _
                    "T\u1ec7p: ") & wb.Name & vbCrLf & _
                    Uni("\u0110\u1ec3 b\u1ea3o to\u00e0n d\u1eef li\u1ec7u, h\u1ec7 th\u1ed1ng kh\u00f4ng t\u1ef1 \u0111\u1ed9ng s\u1eeda t\u1ec7p khi ch\u01b0a sao l\u01b0u \u0111\u01b0\u1ee3c." & vbCrLf & _
                    "Vui l\u00f2ng sao l\u01b0u th\u1ee7 c\u00f4ng v\u00e0 ch\u1ea1y Chay_Diet_Virus_Ngoai.bat."), _
                vbCritical, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - C\u1ea3nh b\u00e1o an to\u00e0n")
    End If
    
    Exit Sub
    
ScanError:
    WriteLog "[ERROR] ScanWorkbook failed for " & wb.Name & ": " & Err.Description
    On Error GoTo 0
End Sub

' ===========================================================================
' HAM KHOI CHAY TIEN TRINH QUET NGAM DOC LAP (v3.6.0 - ASYNCHRONOUS WORKER)
' Chay ngoai tien trinh (Out-of-Process), thoat ngay trong 10ms
' Ho tro tham so bResume de tiep tuc phien quet truoc do
' ===========================================================================
Public Sub LaunchBackgroundScanner(ByVal folderPath As String, Optional ByVal bResume As Boolean = False, Optional ByVal bForce As Boolean = False)
    On Error GoTo LaunchErr
    
    ' VACCINE v3.5.4+: Cat bo toan bo dau \ o cuoi duong dan de tranh loi Command Line Escaping (\" trong Windows CLI)
    Dim cleanFolder As String
    cleanFolder = folderPath
    Do While Right(cleanFolder, 1) = "\" And Len(cleanFolder) > 0
        cleanFolder = Left(cleanFolder, Len(cleanFolder) - 1)
    Loop
    
    ' v3.10.0: Chong bat trung - moi thu muc chi khoi chay 1 lan trong 30 phut (tru Resume / nguoi dung chu dong)
    If Not bResume And Not bForce And Len(cleanFolder) > 0 Then
        If dicLaunchedFolders Is Nothing Then Set dicLaunchedFolders = CreateObject("Scripting.Dictionary")
        Dim folderKey As String
        folderKey = LCase(cleanFolder)
        If dicLaunchedFolders.Exists(folderKey) Then
            If DateDiff("n", dicLaunchedFolders(folderKey), Now) < RELAUNCH_FOLDER_MINUTES Then
                WriteLog "[BG_LAUNCH_SKIPPED] Already scanning recently: " & cleanFolder
                Exit Sub
            End If
        End If
        dicLaunchedFolders(folderKey) = Now
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim scannerScript As String
    scannerScript = ""
    
    ' Uu tien 1: Doc duong dan da dang ky trong Registry (Chong Antivirus/EDR xoa file trong APPDATA)
    Dim regPath As String
    On Error Resume Next
    regPath = wsh.RegRead("HKCU\Software\KangatangGuard\ScannerScript")
    On Error GoTo LaunchErr
    If Len(regPath) > 0 Then
        If fso.FileExists(regPath) Then
            scannerScript = regPath
        End If
    End If
    
    ' Uu tien 2: Kiem tra trong thu muc Add-in goc dang chay
    If Len(scannerScript) = 0 Then
        Dim addinDir As String
        addinDir = ThisWorkbook.Path & "\"
        If fso.FileExists(addinDir & "Kangatang_FolderScanner.ps1") Then
            scannerScript = addinDir & "Kangatang_FolderScanner.ps1"
        End If
    End If
    
    ' Uu tien 3: Kiem tra thu muc du an goc tren o D:
    If Len(scannerScript) = 0 Then
        Dim defaultRepoPath As String
        defaultRepoPath = "d:\7. AI tools\kangatang\addin\Kangatang_FolderScanner.ps1"
        If fso.FileExists(defaultRepoPath) Then
            scannerScript = defaultRepoPath
        End If
    End If
    
    ' Uu tien 4: Kiem tra Desktop mirror
    If Len(scannerScript) = 0 Then
        Dim desktopRepoPath As String
        desktopRepoPath = Environ("USERPROFILE") & "\Desktop\python\coding\AI tools\kangatang\addin\Kangatang_FolderScanner.ps1"
        If fso.FileExists(desktopRepoPath) Then
            scannerScript = desktopRepoPath
        End If
    End If
    
    ' Uu tien 5: Kiem tra trong APPDATA\KangatangGuard
    If Len(scannerScript) = 0 Then
        Dim appDataScript As String
        appDataScript = Environ("APPDATA") & "\" & LOG_SUBFOLDER & "\Kangatang_FolderScanner.ps1"
        If fso.FileExists(appDataScript) Then
            scannerScript = appDataScript
        End If
    End If
    
    ' AUTO-REPAIR v3.9.0: Neu khong tim thay file scanner, tu dong tai tu Hub LAN truoc khi bao loi
    If Len(scannerScript) = 0 Then
        WriteLog "[AUTO_REPAIR] Scanner script not found locally, attempting download from LAN Hub..."
        Dim hubSources As Variant
        hubSources = Array(DEFAULT_HUB_PRIMARY, DEFAULT_HUB_BACKUP)
        Dim repairTarget As String
        repairTarget = Environ("APPDATA") & "\" & LOG_SUBFOLDER & "\Kangatang_FolderScanner.ps1"
        
        ' Dam bao thu muc ton tai
        Dim repairDir As String
        repairDir = Environ("APPDATA") & "\" & LOG_SUBFOLDER
        If Not fso.FolderExists(repairDir) Then
            fso.CreateFolder repairDir
        End If
        
        Dim h As Long
        For h = LBound(hubSources) To UBound(hubSources)
            Dim hubScanner As String
            hubScanner = hubSources(h) & "\Kangatang_FolderScanner.ps1"
            On Error Resume Next
            If IsUncReachable(CStr(hubSources(h))) Then
                If fso.FileExists(hubScanner) Then
                    fso.CopyFile hubScanner, repairTarget, True
                    If Err.Number = 0 And fso.FileExists(repairTarget) Then
                        scannerScript = repairTarget
                        ' Cap nhat Registry de lan sau khoi can repair
                        wsh.RegWrite "HKCU\Software\KangatangGuard\ScannerScript", repairTarget, "REG_SZ"
                        WriteLog "[AUTO_REPAIR] SUCCESS - Downloaded scanner from: " & hubSources(h)
                        On Error GoTo LaunchErr
                        GoTo RepairDone
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo LaunchErr
        Next h
    End If
    
RepairDone:
    ' Kiem tra cuoi cung: Neu van khong tim thay file script thi thong bao ro rang
    If Len(scannerScript) = 0 Or Not fso.FileExists(scannerScript) Then
        WriteLog "[ERROR] Scanner script not found after all fallbacks and auto-repair attempts"
        Set fso = Nothing
        Set wsh = Nothing
        MsgBoxW Uni("Kh\u00f4ng t\u00ecm th\u1ea5y t\u1ec7p tr\u00ecnh qu\u00e9t ng\u1ea7m: Kangatang_FolderScanner.ps1!" & vbCrLf & vbCrLf & _
                    "H\u1ec7 th\u1ed1ng \u0111\u00e3 th\u1eed t\u1ea3i l\u1ea1i t\u1eeb M\u00e1y ch\u1ee7 LAN nh\u01b0ng kh\u00f4ng th\u00e0nh c\u00f4ng." & vbCrLf & _
                    "Vui l\u00f2ng ki\u1ec3m tra k\u1ebft n\u1ed1i m\u1ea1ng v\u00e0 ch\u1ea1y l\u1ea1i t\u1ec7p C\u00e0i \u0111\u1eb7t Add-in ho\u1eb7c li\u00ean h\u1ec7 IT."), _
                vbCritical, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - Thi\u1ebfu t\u1ec7p h\u1ec7 th\u1ed1ng")
        Exit Sub
    End If
    Set fso = Nothing
    
    ' VACCINE v3.6.0: Ho tro tham so -Resume
    Dim cmd As String
    If bResume Then
        If Len(cleanFolder) > 0 Then
            cmd = "powershell.exe -NoExit -ExecutionPolicy Bypass -File """ & scannerScript & """ -TargetFolder """ & cleanFolder & """ -Resume"
        Else
            cmd = "powershell.exe -NoExit -ExecutionPolicy Bypass -File """ & scannerScript & """ -Resume"
        End If
    Else
        cmd = "powershell.exe -NoExit -ExecutionPolicy Bypass -File """ & scannerScript & """ -TargetFolder """ & cleanFolder & """"
    End If
    
    ' VACCINE v3.8.3: Don sach DocumentRecovery truoc khi khoi chay quet
    Call CleanDocumentRecoveryRegistry
    
    ' Tham so 1 = Normal window, False = Asynchronous non-blocking (Khong cho, thoat ngay lap tuc!)
    wsh.Run cmd, 1, False
    Set wsh = Nothing
    
    WriteLog "[BG_LAUNCH] Da khoi chay tien trinh quet ngam cho: " & cleanFolder & " (Resume: " & bResume & ", Script: " & scannerScript & ")"
    Exit Sub
    
LaunchErr:
    WriteLog "[BG_LAUNCH_ERROR] Loi khoi chay worker: " & Err.Description
    On Error GoTo 0
End Sub

' ===========================================================================
' HAM BACKUP: Tao ban sao file truoc khi lam sach (v3.8.5 Centralized Quarantine)
' Chuyen toan bo ban sao ve Kho cach ly tap trung tren May chu LAN
' Khong tao thu muc _Backup_Kangatang tai vi tri lam viec cua nguoi dung
' ===========================================================================
Public Function BackupBeforeClean(ByVal wb As Workbook) As Boolean
    On Error GoTo BackupError
    BackupBeforeClean = False
    
    Dim filePath As String
    filePath = wb.FullName
    
    If Len(filePath) = 0 Or InStrRev(filePath, "\") = 0 Then
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' v3.10.0: Xac dinh kho cach ly qua GetQuarantineDir (co ping-guard, khong treo khi mat LAN)
    Dim backupDir As String
    Dim isCentral As Boolean
    backupDir = GetQuarantineDir(isCentral)
    If Len(backupDir) = 0 Then GoTo BackupError
    
    Dim timestamp As String
    timestamp = Format(Now, "yyyyMMdd_HHmmss")
    
    Dim compName As String
    compName = Environ("COMPUTERNAME")
    If Len(compName) = 0 Then compName = "UNKNOWN_PC"
    
    Dim originalName As String
    originalName = wb.Name
    
    Dim baseName As String, extName As String
    baseName = fso.GetBaseName(originalName)
    extName = fso.GetExtensionName(originalName)
    
    Dim backupName As String
    backupName = baseName & "_backup_" & compName & "_" & timestamp & "." & extName
    
    Dim destPath As String
    destPath = backupDir & backupName
    
    ' v3.10.0: KHONG BAO GIO xoa ban sao luu da ton tai - them hau to so neu trung ten
    Dim suffixN As Long
    suffixN = 1
    Do While fso.FileExists(destPath) And suffixN < 100
        suffixN = suffixN + 1
        destPath = backupDir & baseName & "_backup_" & compName & "_" & timestamp & "_" & suffixN & "." & extName
    Loop
    
    fso.CopyFile filePath, destPath, False
    
    ' Xac minh ban sao thuc su ton tai truoc khi cho phep lam sach
    If Not fso.FileExists(destPath) Then
        Err.Raise vbObjectError + 701, "BackupBeforeClean", "Backup copy not found after CopyFile: " & destPath
    End If
    
    ' Chuan hoa thuoc tinh ve Archive/Normal
    On Error Resume Next
    fso.GetFile(destPath).Attributes = 0
    On Error GoTo BackupError
    
    Set fso = Nothing
    
    If isCentral Then
        WriteLog "[BACKUP_CENTRAL] Da cach ly file nhiem ve Hub tap trung: " & destPath
    Else
        WriteLog "[BACKUP_OFFLINE_QUARANTINE] Mang LAN offline, da luu cach ly cuc bo tai: " & destPath
    End If
    
    BackupBeforeClean = True
    Exit Function
    
BackupError:
    WriteLog "[BACKUP_ERROR] " & Err.Description & " | File: " & filePath
    BackupBeforeClean = False
End Function

' ===========================================================================
' THU TUC PHIM TAT NATIVE COPY / PASTE / CUT (FAIL-SAFE DUAL ROUTING v3.12.2)
' Dam bao Ctrl+C, Ctrl+V, Ctrl+X luon hoat dong 100% tren moi workbook/sheet
' Khong dung ExecuteMso de tranh xoa buffer CutCopyMode trong tien trinh VBA
' ===========================================================================
Public Sub Kangatang_ShortcutCopy()
    On Error Resume Next
    If Not Selection Is Nothing Then
        Selection.Copy
    ElseIf Not ActiveCell Is Nothing Then
        ActiveCell.Copy
    End If
    On Error GoTo 0
End Sub

Public Sub Kangatang_ShortcutPaste()
    On Error Resume Next
    
    ' Truong hop 1: Dang copy vung o tinh trong Excel (CutCopyMode = xlCopy = 1)
    If Application.CutCopyMode = 1 Then
        ' Buoc 1a: Dan cong thuc & gia tri (tranh loi 1004 cua xlPasteAll tren Office 2016/2019/365)
        Selection.PasteSpecial -4123 ' xlPasteFormulas
        If Err.Number <> 0 Then
            Err.Clear
            Selection.PasteSpecial -4163 ' xlPasteValues
        End If
        ' Buoc 1b: Dan toan bo dinh dang (mau sac, font chu, vien ke, number format)
        Selection.PasteSpecial -4122 ' xlPasteFormats
        Err.Clear
        Exit Sub
    End If
    
    ' Truong hop 2: Dang cut vung o tinh (CutCopyMode = xlCut = 2)
    If Application.CutCopyMode = 2 Then
        ActiveSheet.Paste
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If
    
    ' Truong hop 3: Du lieu tu clipboard ben ngoai (Text, Web, Notepad...) hoac hinh anh/shape/chart
    ActiveSheet.PasteSpecial Format:="Unicode Text"
    If Err.Number <> 0 Then
        Err.Clear
        ActiveSheet.Paste
    End If
    If Err.Number <> 0 Then
        Err.Clear
        ActiveCell.PasteSpecial -4163 ' xlPasteValues fallback
    End If
    
    On Error GoTo 0
End Sub

Public Sub Kangatang_ShortcutCut()
    On Error Resume Next
    If Not Selection Is Nothing Then
        Selection.Cut
    ElseIf Not ActiveCell Is Nothing Then
        ActiveCell.Cut
    End If
    On Error GoTo 0
End Sub

' ===========================================================================
' KHOI PHUC CLIPBOARD, PHIM TAT VA GIAO DIEN EXCEL (v3.12.1)
' Khac phuc triet de loi khong the Copy/Paste hoac mat chuot phai do tan du virus
' ===========================================================================
Public Sub RestoreExcelClipboardAndUI(Optional ByVal targetWb As Workbook = Nothing)
    On Error Resume Next
    
    Dim sAddinPrefix As String
    sAddinPrefix = "'" & ThisWorkbook.Name & "'!"
    
    ' 1. Go bo toan bo cac phim tat bi virus chiem dung hoac vo hieu hoa
    Application.OnKey "^c"
    Application.OnKey "^C"
    Application.OnKey "^{c}"
    Application.OnKey "^{C}"
    Application.OnKey "^+{c}"
    Application.OnKey "^+{C}"
    Application.OnKey "^{INSERT}"
    
    Application.OnKey "^v"
    Application.OnKey "^V"
    Application.OnKey "^{v}"
    Application.OnKey "^{V}"
    Application.OnKey "^+{v}"
    Application.OnKey "^+{V}"
    Application.OnKey "+{INSERT}"
    
    Application.OnKey "^x"
    Application.OnKey "^X"
    Application.OnKey "^{x}"
    Application.OnKey "^{X}"
    Application.OnKey "^+{x}"
    Application.OnKey "^+{X}"
    Application.OnKey "+{DELETE}"
    
    Application.OnKey "^z"
    Application.OnKey "^Z"
    Application.OnKey "^y"
    Application.OnKey "^Y"
    Application.OnKey "^d"
    Application.OnKey "^D"
    Application.OnKey "^r"
    Application.OnKey "^R"
    Application.OnKey "%{F11}"
    Application.OnKey "%{F8}"
    
    ' 2. Enforced Shortcut Routing (v3.12.1):
    ' Dinh tuyen truc tiep Ctrl+C, Ctrl+V, Ctrl+X vao quy trinh sao chep cua Add-in
    ' Khac phuc triet de tinh trang liet phim tat do hong ban accelerator cua Office
    Application.OnKey "^c", sAddinPrefix & "Kangatang_ShortcutCopy"
    Application.OnKey "^C", sAddinPrefix & "Kangatang_ShortcutCopy"
    Application.OnKey "^{c}", sAddinPrefix & "Kangatang_ShortcutCopy"
    Application.OnKey "^{C}", sAddinPrefix & "Kangatang_ShortcutCopy"
    Application.OnKey "^{INSERT}", sAddinPrefix & "Kangatang_ShortcutCopy"
    
    Application.OnKey "^v", sAddinPrefix & "Kangatang_ShortcutPaste"
    Application.OnKey "^V", sAddinPrefix & "Kangatang_ShortcutPaste"
    Application.OnKey "^{v}", sAddinPrefix & "Kangatang_ShortcutPaste"
    Application.OnKey "^{V}", sAddinPrefix & "Kangatang_ShortcutPaste"
    Application.OnKey "+{INSERT}", sAddinPrefix & "Kangatang_ShortcutPaste"
    
    Application.OnKey "^x", sAddinPrefix & "Kangatang_ShortcutCut"
    Application.OnKey "^X", sAddinPrefix & "Kangatang_ShortcutCut"
    Application.OnKey "^{x}", sAddinPrefix & "Kangatang_ShortcutCut"
    Application.OnKey "^{X}", sAddinPrefix & "Kangatang_ShortcutCut"
    Application.OnKey "+{DELETE}", sAddinPrefix & "Kangatang_ShortcutCut"
    
    ' 3. Khoi phuc tinh nang keo tha o tinh
    Application.CellDragAndDrop = True
    
    ' 4. Bat lai va Reset TOAN BO cac CommandBars (Cell, Row, Column, etc.)
    Dim iCb As Long, cbItem As Object, ctrlItem As Object
    For iCb = 1 To Application.CommandBars.Count
        Set cbItem = Nothing
        Set cbItem = Application.CommandBars(iCb)
        If Not cbItem Is Nothing Then
            Select Case cbItem.Name
                Case "Cell", "Row", "Column", "XLM Cell", "Standard", "Worksheet Menu Bar"
                    cbItem.Enabled = True
                    cbItem.Reset
                    For Each ctrlItem In cbItem.Controls
                        ctrlItem.Enabled = True
                    Next ctrlItem
            End Select
        End If
    Next iCb
    
    ' Bat cac nut Copy, Cut, Paste tren he thong theo Office Control ID
    Dim sysIds As Variant, vId As Variant, fCtrl As Object
    sysIds = Array(19, 21, 22, 21437, 3624, 755, 369)
    For Each vId In sysIds
        Set fCtrl = Nothing
        Set fCtrl = Application.CommandBars.FindControl(Id:=CLng(vId))
        If Not fCtrl Is Nothing Then fCtrl.Enabled = True
    Next vId
    
    ' 5. Xoa triet de cac hook su kien mo coi cua Excel Application
    Application.OnSheetActivate = ""
    Application.OnSheetDeactivate = ""
    Application.OnWindow = ""
    Application.OnCalculate = ""
    Application.OnDoubleClick = ""
    Application.OnEntry = ""
    
    ' 6. Khoi phuc quyen chon o tinh va bo khoa tren cac sheet (xlNoRestrictions = -4142)
    Dim wbToFix As Workbook, sItem As Object
    Set wbToFix = targetWb
    If wbToFix Is Nothing Then Set wbToFix = ActiveWorkbook
    If Not wbToFix Is Nothing Then
        For Each sItem In wbToFix.Sheets
            On Error Resume Next
            Call TryUnprotectSheet(sItem)
            sItem.EnableSelection = -4142
        Next sItem
    End If
    
    WriteLog "[UI_RESTORE] Excel clipboard, shortcuts, context menus, and event hooks restored (v" & CURRENT_VERSION & ")."
    On Error GoTo 0
End Sub

' ===========================================================================
' HAM MO KHOA WORKBOOK VA SHEET CO MAT KHAU TRONG (v3.11.0)
' Giup lam sach virus trong cac file/sheet bi bao ve cau truc
' ===========================================================================
Public Function TryUnprotectWorkbook(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then Exit Function
    If Not wb.ProtectStructure And Not wb.ProtectWindows Then
        TryUnprotectWorkbook = True
        Exit Function
    End If
    
    ' Thu go bao ve cau truc voi mat khau rong
    wb.Unprotect ""
    Err.Clear
    
    TryUnprotectWorkbook = (Not wb.ProtectStructure And Not wb.ProtectWindows)
End Function

Public Function TryUnprotectSheet(ByVal sh As Object) As Boolean
    On Error Resume Next
    If sh Is Nothing Then Exit Function
    If Not sh.ProtectContents Then
        TryUnprotectSheet = True
        Exit Function
    End If
    
    ' Thu go bao ve sheet voi mat khau rong
    sh.Unprotect ""
    Err.Clear
    
    TryUnprotectSheet = (Not sh.ProtectContents)
End Function

' ===========================================================================
' HAM LAM SACH (v3.11.0 - Surgical Clean & Protection Handling)
' - Component mang ten virus (Type 1/2/3)  -> xoa ca component (luu mau ma doc truoc khi xoa)
' - Component khac co IOC trong code      -> CHI xoa procedure doc hai (luu mau ma doc)
' - Sheet / Name khop chu ky               -> xoa; moi thao tac co lap loi rieng
' - Ho tro go bao ve cau truc workbook va sheet (blank password)
' - Xac minh lai sau khi lam sach. bSave = False khi dang trong BeforeSave hoac file Read-Only.
' ===========================================================================
Public Function CleanInfectedWorkbook(ByVal wb As Workbook, Optional ByVal bSave As Boolean = True, Optional ByVal bForceSave As Boolean = False) As Boolean
    On Error GoTo CleanError
    CleanInfectedWorkbook = False
    
    Dim i As Long
    Dim changed As Boolean
    Dim failures As Long
    Dim nModuleSheetsDeferred As Long
    
    ' Go hook su kien cua virus trong RAM truoc khi xoa code (tranh loi "Cannot run macro" va tai nhiem)
    Call NeutralizeVirusHooks
    
    ' Kiem tra va go bao ve cau truc workbook neu co (v3.11.0)
    If wb.ProtectStructure Then
        If Not TryUnprotectWorkbook(wb) Then
            WriteLog "[CLEAN_WARN] Workbook structure protected by password."
        End If
    End If
    
    ' --- 1. Sheets TRUOC (v3.10.0): VeryHidden / ban sao "Kangatang (2)" ---
    ' Module-sheet Excel 5 KHONG xoa tai day: xoa tu VBA se huy call stack + reset project add-in
    ' -> giao cho Two-phase purge (OnTime) o cuoi ham.
    Dim shtName As String
    For i = wb.Sheets.Count To 1 Step -1
        shtName = wb.Sheets(i).Name
        Call TryUnprotectSheet(wb.Sheets(i))
        If IsVirusName(shtName) Then
            If TypeName(wb.Sheets(i)) = "Module" Then
                nModuleSheetsDeferred = nModuleSheetsDeferred + 1
            Else
                ' Thu thap mau ma doc cua sheet truoc khi xoa
                Call CollectThreatSample(wb, shtName, "[Sheet: " & shtName & ", Type: " & TypeName(wb.Sheets(i)) & "]", "InfectedSheet")
                If DeleteSheetSafe(wb, i) Then
                    changed = True
                    WriteLog "[CLEAN] Removed sheet: " & shtName
                Else
                    failures = failures + 1
                End If
            End If
        End If
    Next i
    
    ' --- 2 + 3. VBA components ---
    Dim vbProj As Object, comp As Object
    Dim compName As String, compType As Long, nRemoved As Long
    Dim vbOK As Boolean
    If LegacyModuleSheetCount(wb) <> 0 Then
        ' Crash guard: con module-sheet (khong mang ten virus) -> khong mo VBProject
        WriteLog "[VBP_SKIPPED] " & wb.FullName & " - legacy module sheet remains, VBA components not cleaned."
        vbOK = False
    Else
        On Error Resume Next
        Set vbProj = wb.VBProject
        vbOK = (Err.Number = 0 And Not vbProj Is Nothing)
        Err.Clear
        If vbOK Then
            If vbProj.Protection = 1 Then
                vbOK = False
                WriteLog "[CLEAN_WARN] VBProject is password-protected (VBA components locked)."
            End If
            Err.Clear
        End If
        On Error GoTo CleanError
    End If
    
    If vbOK Then
        For i = vbProj.VBComponents.Count To 1 Step -1
            Set comp = vbProj.VBComponents.Item(i)
            compName = comp.Name
            compType = comp.Type
            If IsVirusName(compName) And (compType = 1 Or compType = 2 Or compType = 3) Then
                ' Thu thap mau ma doc truoc khi xoa component
                Dim rawCompCode As String
                rawCompCode = ""
                On Error Resume Next
                If Not comp.CodeModule Is Nothing Then
                    If comp.CodeModule.CountOfLines > 0 Then
                        rawCompCode = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                    End If
                End If
                Err.Clear
                Call CollectThreatSample(wb, compName, rawCompCode, "VirusModule")
                
                On Error Resume Next
                vbProj.VBComponents.Remove comp
                If Err.Number <> 0 Then
                    WriteLog "[CLEAN_WARN] Cannot remove component " & compName & ": " & Err.Description
                    failures = failures + 1
                Else
                    WriteLog "[CLEAN] Removed component: " & compName
                    changed = True
                End If
                Err.Clear
                On Error GoTo CleanError
            Else
                ' Thu thap mau ma doc truoc khi xoa procedures co IOC
                Dim rawProcCode As String
                rawProcCode = ""
                On Error Resume Next
                If Not comp.CodeModule Is Nothing Then
                    If comp.CodeModule.CountOfLines > 0 Then
                        rawProcCode = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                        If CodeHasIOC(rawProcCode) Then
                            Call CollectThreatSample(wb, compName, rawProcCode, "InfectedProcedure")
                        End If
                    End If
                End If
                Err.Clear
                nRemoved = RemoveInfectedProcedures(comp)
                If nRemoved > 0 Then
                    changed = True
                    WriteLog "[CLEAN] Removed " & nRemoved & " infected procedure(s) from: " & compName
                    ' Module chuan chi chua ma doc (khong con procedure nao) -> xoa component rong
                    If (compType = 1 Or compType = 2) Then
                        If CountProcedures(comp) = 0 Then
                            On Error Resume Next
                            vbProj.VBComponents.Remove comp
                            If Err.Number = 0 Then WriteLog "[CLEAN] Removed empty component: " & compName
                            Err.Clear
                            On Error GoTo CleanError
                        End If
                    End If
                End If
            End If
        Next i
    End If
    
    ' --- 4. Named Ranges ---
    Dim nm As Object, refText As String, nmName As String
    For i = wb.Names.Count To 1 Step -1
        On Error Resume Next
        Set nm = wb.Names(i)
        nmName = nm.Name
        refText = ""
        refText = nm.RefersTo
        Err.Clear
        If IsVirusName(nmName) Or IsVirusReference(refText) Then
            nm.Delete
            If Err.Number = 0 Then
                changed = True
                WriteLog "[CLEAN] Removed name: " & nmName
            Else
                WriteLog "[CLEAN_WARN] Cannot remove name " & nmName & ": " & Err.Description
                failures = failures + 1
            End If
            Err.Clear
        End If
        On Error GoTo CleanError
    Next i
    
    ' --- 4b. Module-sheet virus -> Two-phase purge (OnTime), dung xac minh/luu tai day ---
    If nModuleSheetsDeferred > 0 Then
        ' Phase 2 se luu neu tep ghi duoc (ke ca khi goi tu BeforeSave: nguoi dung dang muon luu,
        ' ban luu hien tai van con module-sheet -> Phase 2 ghi de ban sach ngay sau do).
        Call DeferModuleSheetPurge(wb, (Not wb.ReadOnly) And (Len(wb.Path) > 0))
        WriteLog "[CLEAN_DEFERRED] " & wb.FullName & " - " & nModuleSheetsDeferred & " legacy module sheet(s) handed to two-phase purge."
        Call RestoreExcelClipboardAndUI(wb)
        CleanInfectedWorkbook = False
        Exit Function
    End If
    
    ' --- 5. Xac minh lai ---
    Dim remain As String, stillInfected As Boolean
    stillInfected = CheckWorkbookInfection(wb, remain)
    If stillInfected Then
        WriteLog "[CLEAN_PARTIAL] " & wb.FullName & " (failures=" & failures & ") still has: " & Replace(remain, vbCrLf, " ")
    End If
    
    ' --- 6. Luu (tat hop thoai Compatibility Checker cua .xls) ---
    If (changed Or bForceSave) And bSave Then
        Dim prevAlerts As Boolean
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        On Error Resume Next
        wb.Save
        If Err.Number <> 0 Then
            WriteLog "[CLEAN_SAVE_ERROR] " & wb.FullName & ": " & Err.Description
            Err.Clear
            Application.DisplayAlerts = prevAlerts
            Call RestoreExcelClipboardAndUI(wb)
            On Error GoTo CleanError
            Exit Function
        End If
        Application.DisplayAlerts = prevAlerts
        On Error GoTo CleanError
    End If
    
    ' Khoi phuc clipboard va giao dien sau khi lam sach
    Call RestoreExcelClipboardAndUI(wb)
    
    CleanInfectedWorkbook = Not stillInfected
    Exit Function
    
CleanError:
    WriteLog "[CLEAN_ERROR] " & wb.FullName & ": " & Err.Description
    Application.DisplayAlerts = True
    Call RestoreExcelClipboardAndUI(wb)
    CleanInfectedWorkbook = False
End Function

' ===========================================================================
' HAM MENU: Quet file hien tai (goi tu menu)
' ===========================================================================
Public Sub ScanActiveWorkbook()
    If ActiveWorkbook Is Nothing Then
        MsgBoxW Uni("Kh\u00f4ng c\u00f3 t\u1ec7p Excel n\u00e0o \u0111ang m\u1edf."), vbInformation, Uni("KangatangGuard v") & CURRENT_VERSION
        Exit Sub
    End If
    
    Dim wbName As String
    wbName = ActiveWorkbook.Name
    
    Call ScanWorkbook(ActiveWorkbook)
    
    MsgBoxW Uni("\u0110\u00e3 ho\u00e0n th\u00e0nh qu\u00e9t t\u1ec7p: ") & wbName & vbCrLf & _
            Uni("Chi ti\u1ebft \u0111\u01b0\u1ee3c ghi t\u1ea1i th\u01b0 m\u1ee5c nh\u1eadt k\u00fd (Log)."), _
            vbInformation, Uni("KangatangGuard v") & CURRENT_VERSION
End Sub

' ===========================================================================
' HAM MENU: Quet thu muc moi (goi tu menu, hien hop thoai chon thu muc)
' v3.6.0: Tach tien trinh quet ngam doc lap khong lam treo Excel
' ===========================================================================
Public Sub ScanFolderDialog()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = Uni("KangatangGuard - Ch\u1ecdn th\u01b0 m\u1ee5c c\u1ea7n qu\u00e9t")
    fd.ButtonName = Uni("Qu\u00e9t th\u01b0 m\u1ee5c n\u00e0y")
    
    If fd.Show = -1 Then
        Dim folderPath As String
        folderPath = fd.SelectedItems(1)
        
        ' v3.6.0: Goi worker ngam doc lap (non-blocking)
        Call LaunchBackgroundScanner(folderPath, bForce:=True)
        
        ' Thong bao nhanh cho nguoi dung (Excel tiep tuc hoat dong ngay lap tuc)
        MsgBoxW Uni("\u0110\u00e3 kh\u1edfi ch\u1ea1y ti\u1ebfn tr\u00ecnh qu\u00e9t ng\u1ea7m \u0111\u1ed9c l\u1eadp cho th\u01b0 m\u1ee5c:" & vbCrLf & vbCrLf) & _
                folderPath & vbCrLf & vbCrLf & _
                Uni("Ti\u1ebfn tr\u00ecnh \u0111ang ch\u1ea1y tr\u00ean c\u1eeda s\u1ed5 ri\u00eang bi\u1ec7t v\u1edbi thanh ti\u1ebfn \u0111\u1ed9 % th\u1eddi gian th\u1ef1c." & vbCrLf & _
                    "B\u1ea1n c\u00f3 th\u1ec3 TI\u1ebeP T\u1ee4C L\u00c0M VI\u1ec6C tr\u00ean Excel b\u00ecnh th\u01b0\u1eddng m\u00e0 kh\u00f4ng lo b\u1ecb treo m\u00e1y!"), _
                vbInformation, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - Ti\u1ebfn tr\u00ecnh ng\u1ea7m")
    End If
End Sub

' ===========================================================================
' HAM MENU: Tiep tuc phien quet truoc do (v3.6.0 - Resume In-Progress Scan)
' ===========================================================================
Public Sub ResumeScanDialog()
    On Error GoTo ResumeErr
    
    Dim lastSessionPath As String
    lastSessionPath = Environ("APPDATA") & "\" & LOG_SUBFOLDER & "\last_session.json"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(lastSessionPath) Then
        Set fso = Nothing
        MsgBoxW Uni("Kh\u00f4ng t\u00ecm th\u1ea5y phi\u00ean qu\u00e9t n\u00e0o tr\u01b0\u1edbc \u0111\u00f3 \u0111\u1ec3 ti\u1ebfp t\u1ee5c."), _
                vbInformation, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - Ti\u1ebfp t\u1ee5c phi\u00ean")
        Exit Sub
    End If
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(lastSessionPath, 1, False, -2)
    Dim jsonText As String
    jsonText = ts.ReadAll
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Dim statusVal As String, targetVal As String, scannedVal As String, cleanedVal As String, timeVal As String
    statusVal = ExtractJsonValue(jsonText, "Status")
    targetVal = ExtractJsonValue(jsonText, "TargetFolder")
    scannedVal = ExtractJsonValue(jsonText, "TotalScanned")
    cleanedVal = ExtractJsonValue(jsonText, "TotalCleaned")
    timeVal = ExtractJsonValue(jsonText, "StartTime")
    
    If UCase(statusVal) = "COMPLETED" Then
        MsgBoxW Uni("Phi\u00ean qu\u00e9t g\u1ea7n nh\u1ea5t \u0111\u00e3 HO\u00c0N T\u1ea4T 100%!" & vbCrLf & _
                    "Th\u01b0 m\u1ee5c: ") & targetVal & vbCrLf & _
                Uni("\u0110\u00e3 qu\u00e9t: ") & scannedVal & Uni(" t\u1ec7p | \u0110\u00e3 di\u1ec7t: ") & cleanedVal & Uni(" t\u1ec7p" & vbCrLf & vbCrLf & _
                    "B\u1ea1n c\u00f3 th\u1ec3 ch\u1ecdn 'Qu\u00e9t th\u01b0 m\u1ee5c m\u1edbi...' n\u1ebfu mu\u1ed1n qu\u00e9t l\u1ea1i ho\u1eb7c ch\u1ecdn th\u01b0 m\u1ee5c kh\u00e1c."), _
                vbInformation, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - Phi\u00ean \u0111\u00e3 ho\u00e0n t\u1ea5t")
        Exit Sub
    End If
    
    Dim promptMsg As String
    promptMsg = Uni("T\u00ecm th\u1ea5y phi\u00ean qu\u00e9t d\u1edf dang ch\u01b0a ho\u00e0n th\u00e0nh:" & vbCrLf & vbCrLf) & _
                Uni("\u2022 Th\u01b0 m\u1ee5c: ") & targetVal & vbCrLf & _
                Uni("\u2022 Th\u1eddi gian: ") & timeVal & vbCrLf & _
                Uni("\u2022 Ti\u1ebfn \u0111\u1ed9: \u0110\u00e3 qu\u00e9t ") & scannedVal & Uni(" t\u1ec7p (\u0110\u00e3 di\u1ec7t: ") & cleanedVal & Uni(" t\u1ec7p)" & vbCrLf & vbCrLf) & _
                Uni("B\u1ea1n c\u00f3 mu\u1ed1n TI\u1ebeP T\u1ee4C qu\u00e9t t\u1eeb v\u1ecb tr\u00ed n\u00e0y kh\u00f4ng?" & vbCrLf & _
                    "(C\u00e1c t\u1ec7p \u0111\u00e3 ki\u1ec3m tra s\u1ebd \u0111\u01b0\u1ee3c b\u1ecf qua si\u00eau t\u1ed1c)")
    
    Dim res As VbMsgBoxResult
    res = MsgBoxW(promptMsg, vbYesNo + vbQuestion, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - Ti\u1ebfp t\u1ee5c phi\u00ean qu\u00e9t"))
    
    If res = vbYes Then
        Call LaunchBackgroundScanner(targetVal, bResume:=True)
        MsgBoxW Uni("\u0110\u00e3 k\u00edch ho\u1ea1t ti\u1ebfp t\u1ee5c phi\u00ean qu\u00e9t ng\u1ea7m tr\u00ean c\u1eeda s\u1ed5 ri\u00eang bi\u1ec7t." & vbCrLf & _
                    "B\u1ea1n c\u00f3 th\u1ec3 ti\u1ebfp t\u1ee5c l\u00e0m vi\u1ec7c tr\u00ean Excel b\u00ecnh th\u01b0\u1eddng!"), _
                vbInformation, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - Ti\u1ebfp t\u1ee5c phi\u00ean")
    End If
    Exit Sub
    
ResumeErr:
    MsgBoxW Uni("L\u1ed7i khi \u0111\u1ecdc th\u00f4ng tin phi\u00ean qu\u00e9t: ") & Err.Description, vbCritical, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - L\u1ed7i")
    On Error GoTo 0
End Sub

Private Function ExtractJsonValue(ByVal json As String, ByVal key As String) As String
    Dim pattern As String, p1 As Long, p2 As Long
    pattern = """" & key & """"
    p1 = InStr(1, json, pattern, vbTextCompare)
    If p1 = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    
    p1 = InStr(p1 + Len(pattern), json, ":")
    If p1 = 0 Then Exit Function
    p1 = p1 + 1
    
    Do While Mid(json, p1, 1) = " " Or Mid(json, p1, 1) = vbTab Or Mid(json, p1, 1) = vbCr Or Mid(json, p1, 1) = vbLf
        p1 = p1 + 1
    Loop
    
    If Mid(json, p1, 1) = """" Then
        p1 = p1 + 1
        p2 = InStr(p1, json, """")
        If p2 > p1 Then
            ExtractJsonValue = Mid(json, p1, p2 - p1)
        End If
    Else
        p2 = p1
        Do While p2 <= Len(json) And Mid(json, p2, 1) <> "," And Mid(json, p2, 1) <> "}" And Mid(json, p2, 1) <> vbCr And Mid(json, p2, 1) <> vbLf
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function

' ===========================================================================
' VACCINE v3.8.3: Tu dong don dep Document Recovery gay phien nguoi dung
' ===========================================================================
Public Sub CleanDocumentRecoveryRegistry()
    On Error Resume Next
    Dim reg As Object
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    If reg Is Nothing Then Exit Sub
    
    Const HKEY_CURRENT_USER = &H80000001
    Dim regPath As String
    regPath = "Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery"
    
    Dim subKeys As Variant
    reg.EnumKey HKEY_CURRENT_USER, regPath, subKeys
    If Not IsArray(subKeys) Then Exit Sub
    
    Dim k As Variant
    For Each k In subKeys
        Dim valNames As Variant
        reg.EnumValues HKEY_CURRENT_USER, regPath & "\" & k, valNames
        Dim bDelete As Boolean
        bDelete = False
        
        If IsArray(valNames) Then
            Dim vn As Variant
            For Each vn In valNames
                Dim binVal As Variant
                reg.GetBinaryValue HKEY_CURRENT_USER, regPath & "\" & k, CStr(vn), binVal
                If IsArray(binVal) Then
                    Dim strVal As String
                    strVal = ""
                    Dim i As Long
                    For i = 0 To UBound(binVal) Step 2
                        If i <= UBound(binVal) Then
                            If binVal(i) > 31 And binVal(i) < 127 Then
                                strVal = strVal & Chr(binVal(i))
                            End If
                        End If
                    Next i
                    ' v3.10.0: Thu hep pham vi - KHONG xoa entry khoi phuc cua moi file LAN (192.168./file_shared)
                    ' vi lam mat du lieu khoi phuc sau crash cua nguoi dung. Chi xoa entry tro toi kho cach ly/backup.
                    If InStr(1, strVal, "_Backup_Kangatang", vbTextCompare) > 0 Or _
                       InStr(1, strVal, "Virus backupfile", vbTextCompare) > 0 Or _
                       InStr(1, strVal, "Quarantine_Backup", vbTextCompare) > 0 Or _
                       InStr(1, strVal, "TBEX", vbTextCompare) > 0 Then
                        bDelete = True
                        Exit For
                    End If
                End If
            Next vn
        End If
        
        If bDelete Then
            reg.DeleteKey HKEY_CURRENT_USER, regPath & "\" & k
        End If
    Next k
    On Error GoTo 0
End Sub

' ===========================================================================
' HAM MENU: Mo thu muc Log
' ===========================================================================
Public Sub OpenLogFolder()
    Dim logDir As String
    logDir = Environ("APPDATA") & "\" & LOG_SUBFOLDER
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(logDir) Then
        fso.CreateFolder logDir
    End If
    Set fso = Nothing
    
    Shell "explorer.exe " & Chr(34) & logDir & Chr(34), vbNormalFocus
End Sub

' ===========================================================================
' HAM MENU: Hien thi thong tin Add-in
' ===========================================================================
Public Sub ShowAbout()
    Dim regVer As String, regSrc As String
    Dim wsh As Object
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    regVer = wsh.RegRead("HKCU\Software\KangatangGuard\InstalledVersion")
    regSrc = wsh.RegRead("HKCU\Software\KangatangGuard\UpdateSource")
    Set wsh = Nothing
    On Error GoTo 0
    
    If Len(regVer) = 0 Then regVer = CURRENT_VERSION
    If Len(regVer) = 0 Then regVer = CURRENT_VERSION
    If Len(regSrc) = 0 Then regSrc = DEFAULT_HUB_PRIMARY
    
    MsgBoxW Uni("KangatangGuard v" & CURRENT_VERSION & vbCrLf & vbCrLf & _
                "H\u1ec7 th\u1ed1ng b\u1ea3o v\u1ec7 Excel chuy\u00ean d\u1ee5ng ch\u1ed1ng virus macro Kangatang / Laroux / mypersonnel." & vbCrLf & vbCrLf & _
                "- T\u1ef1 \u0111\u1ed9ng qu\u00e9t th\u1eddi gian th\u1ef1c khi m\u1edf t\u1ec7p." & vbCrLf & _
                "- ScanCache th\u00f4ng minh ch\u1ed1ng lag khi l\u01b0u v\u00e0 AutoSave." & vbCrLf & _
                "- Ki\u1ebfn tr\u00fac Out-of-Process Worker: Excel kh\u00f4ng bao gi\u1edd b\u1ecb treo (Not Responding)!" & vbCrLf & _
                "- T\u00ednh n\u0103ng Fast Resume: Ti\u1ebfp t\u1ee5c phi\u00ean qu\u00e9t d\u1edf dang si\u00eau t\u1ed1c." & vbCrLf & _
                "- Trung t\u00e2m Ph\u00e2n ph\u1ed1i LAN Hub (\\192.168.223.7) & T\u1ef1 \u0111\u1ed9ng C\u1eadp nh\u1eadt." & vbCrLf & vbCrLf & _
                "\u2022 M\u00e1y ch\u1ee7 ngu\u1ed3n: " & regSrc & vbCrLf & _
                "\u2022 Phi\u00ean b\u1ea3n ki\u1ebfn tr\u00fac: v" & CURRENT_VERSION & " Production-grade"), _
            vbInformation, Uni("Gi\u1edbi thi\u1ec7u KangatangGuard v" & CURRENT_VERSION)
End Sub

' ===========================================================================
' KIEM TRA VA DONG BO CAP NHAT TU TRUNG TAM LAN HUB (v3.8.6)
' ===========================================================================
Public Sub CheckForLanUpdatesManual()
    Call CheckForLanUpdates(bSilent:=False)
End Sub

Public Sub CheckForLanUpdatesSilent()
    Call CheckForLanUpdates(bSilent:=True)
End Sub

Public Sub CheckForLanUpdates(Optional ByVal bSilent As Boolean = True)
    On Error GoTo UpdateErr
    
    Dim wsh As Object, fso As Object
    Dim updateSource As String
    Dim versionFile As String
    Dim serverVer As String
    Dim jsonText As String
    
    Set wsh = CreateObject("WScript.Shell")
    
    ' v3.8.6: Rate-limit 24h khi chay ngam de Excel khoi dong tuc thi, khong bao gio bi timeout mang
    If bSilent Then
        Dim lastCheck As String
        On Error Resume Next
        lastCheck = wsh.RegRead("HKCU\Software\KangatangGuard\LastUpdateCheck")
        On Error GoTo UpdateErr
        If Len(lastCheck) > 0 Then
            If IsDate(lastCheck) Then
                If DateDiff("h", CDate(lastCheck), Now) < UPDATE_CHECK_HOURS Then
                    Set wsh = Nothing
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Doc UpdateSource tu Registry HKCU\Software\KangatangGuard
    On Error Resume Next
    updateSource = wsh.RegRead("HKCU\Software\KangatangGuard\UpdateSource")
    On Error GoTo UpdateErr
    
    ' v3.10.0: Chon Hub qua ping-guard (Registry -> 223.7 -> 223.176), khong cham UNC khi may chu khong phan hoi
    If Right(updateSource, 1) = "\" Then updateSource = Left(updateSource, Len(updateSource) - 1)
    Dim candidates As Variant, c As Long
    candidates = Array(Trim(updateSource), DEFAULT_HUB_PRIMARY, DEFAULT_HUB_BACKUP)
    updateSource = ""
    For c = LBound(candidates) To UBound(candidates)
        If Len(candidates(c)) > 0 Then
            If IsUncReachable(CStr(candidates(c))) Then
                If fso.FileExists(candidates(c) & "\version.json") Then
                    updateSource = candidates(c)
                    Exit For
                End If
            End If
        End If
    Next c
    
    If Len(updateSource) = 0 Then
        If Not bSilent Then
            MsgBoxW Uni("Kh\u00f4ng th\u1ec3 k\u1ebft n\u1ed1i \u0111\u1ebfn M\u00e1y ch\u1ee7 LAN (223.7 / 223.176) ho\u1eb7c kh\u00f4ng t\u00ecm th\u1ea5y version.json." & vbCrLf & vbCrLf & _
                        "Vui l\u00f2ng ki\u1ec3m tra k\u1ebft n\u1ed1i m\u1ea1ng LAN."), _
                    vbExclamation, Uni("KangatangGuard v" & CURRENT_VERSION & " - C\u1eadp nh\u1eadt")
        Else
            On Error Resume Next
            wsh.RegWrite "HKCU\Software\KangatangGuard\LastUpdateCheck", Format(Now, "yyyy-MM-dd HH:mm:ss"), "REG_SZ"
        End If
        Exit Sub
    End If
    
    versionFile = updateSource & "\version.json"
    
    If Not fso.FileExists(versionFile) Then
        If bSilent Then
            On Error Resume Next
            wsh.RegWrite "HKCU\Software\KangatangGuard\LastUpdateCheck", Format(Now, "yyyy-MM-dd HH:mm:ss"), "REG_SZ"
        Else
            MsgBoxW Uni("Kh\u00f4ng th\u1ec3 k\u1ebft n\u1ed1i \u0111\u1ebfn M\u00e1y ch\u1ee7 LAN ho\u1eb7c kh\u00f4ng t\u00ecm th\u1ea5y th\u00f4ng tin phi\u00ean b\u1ea3n:" & vbCrLf & vbCrLf) & _
                    updateSource & vbCrLf & vbCrLf & _
                    Uni("Vui l\u00f2ng ki\u1ec3m tra k\u1ebft n\u1ed1i m\u1ea1ng LAN ho\u1eb7c m\u00e1y ch\u1ee7 c\u00f3 \u0111ang b\u1eadt kh\u00f4ng."), _
                    vbExclamation, Uni("KangatangGuard v" & CURRENT_VERSION & " - K\u1ebft n\u1ed1i th\u1ea5t b\u1ea1i")
        End If
        Exit Sub
    End If
    
    ' Doc noi dung version.json
    Dim ts As Object
    Set ts = fso.OpenTextFile(versionFile, 1, False)
    jsonText = ts.ReadAll
    ts.Close
    Set ts = Nothing
    
    serverVer = ExtractJsonValue(jsonText, "version")
    If Len(serverVer) = 0 Then
        If Not bSilent Then
            MsgBoxW Uni("Kh\u00f4ng th\u1ec3 \u0111\u1ecdc th\u00f4ng tin phi\u00ean b\u1ea3n t\u1eeb t\u1ec7p version.json tr\u00ean m\u00e1y ch\u1ee7."), vbExclamation, Uni("KangatangGuard v" & CURRENT_VERSION)
        End If
        Exit Sub
    End If
    
    ' So sanh phien ban
    If IsNewerVersion(serverVer, CURRENT_VERSION) Then
        Dim promptMsg As String
        Dim changelog As String
        changelog = ExtractJsonValue(jsonText, "changelog")
        
        promptMsg = Uni("M\u00e1y ch\u1ee7 ph\u00e1t h\u00e0nh phi\u00ean b\u1ea3n M\u1edaI: v") & serverVer & "!" & vbCrLf & _
                    Uni("(Phi\u00ean b\u1ea3n hi\u1ec7n t\u1ea1i tr\u00ean m\u00e1y b\u1ea1n: v") & CURRENT_VERSION & ")" & vbCrLf & vbCrLf
        If Len(changelog) > 0 Then
            promptMsg = promptMsg & Uni("N\u1ed9i dung c\u1eadp nh\u1eadt:") & vbCrLf & changelog & vbCrLf & vbCrLf
        End If
        promptMsg = promptMsg & Uni("B\u1ea1n c\u00f3 mu\u1ed1n C\u1eacP NH\u1eacT NGAY kh\u00f4ng?")
        
        ' v3.10.0: Ban cap nhat bao mat bat buoc (version.json: "mandatory": true) -> ap dung khong hoi khi chay ngam
        Dim isMandatory As Boolean
        isMandatory = (LCase(ExtractJsonValue(jsonText, "mandatory")) = "true")
        If isMandatory And bSilent Then
            WriteLog "[UPDATE_MANDATORY] Auto-applying v" & serverVer & " from " & updateSource
            Call PerformLanUpdate(updateSource, serverVer, True)
        Else
            Dim ans As VbMsgBoxResult
            ans = MsgBoxW(promptMsg, vbYesNo + vbInformation, Uni("KangatangGuard - C\u00f3 phi\u00ean b\u1ea3n m\u1edbi v") & serverVer)
            If ans = vbYes Then
                Call PerformLanUpdate(updateSource, serverVer, False)
            End If
        End If
    Else
        If Not bSilent Then
            MsgBoxW Uni("B\u1ea1n \u0111ang s\u1eed d\u1ee5ng phi\u00ean b\u1ea3n M\u1edaI NH\u1ea4T (v") & CURRENT_VERSION & ")!" & vbCrLf & vbCrLf & _
                    Uni("M\u00e1y ch\u1ee7 ph\u00e2n ph\u1ed1i: ") & updateSource, _
                    vbInformation, Uni("KangatangGuard v" & CURRENT_VERSION & " - H\u1ec7 th\u1ed1ng \u0111\u00e3 c\u1eadp nh\u1eadt")
        End If
    End If
    
    ' Luu moc thoi gian kiem tra
    On Error Resume Next
    wsh.RegWrite "HKCU\Software\KangatangGuard\LastUpdateCheck", Format(Now, "yyyy-MM-dd HH:mm:ss"), "REG_SZ"
    On Error GoTo 0
    
    Set fso = Nothing
    Set wsh = Nothing
    Exit Sub
    
UpdateErr:
    If bSilent Then
        On Error Resume Next
        wsh.RegWrite "HKCU\Software\KangatangGuard\LastUpdateCheck", Format(Now, "yyyy-MM-dd HH:mm:ss"), "REG_SZ"
    Else
        MsgBoxW Uni("L\u1ed7i khi ki\u1ec3m tra c\u1eadp nh\u1eadt: ") & Err.Description, vbCritical, Uni("KangatangGuard v" & CURRENT_VERSION & " - L\u1ed7i")
    End If
    On Error GoTo 0
End Sub

Public Function IsNewerVersion(ByVal vServer As String, ByVal vLocal As String) As Boolean
    IsNewerVersion = False
    On Error Resume Next
    
    ' Loai bo ky tu 'v' hoac khoang trang
    vServer = Replace(Replace(vServer, "v", ""), "V", "")
    vLocal  = Replace(Replace(vLocal, "v", ""), "V", "")
    
    Dim sParts() As String, lParts() As String
    sParts = Split(vServer, ".")
    lParts = Split(vLocal, ".")
    
    Dim i As Long, maxParts As Long
    maxParts = UBound(sParts)
    If UBound(lParts) > maxParts Then maxParts = UBound(lParts)
    
    For i = 0 To maxParts
        Dim sNum As Long, lNum As Long
        sNum = 0: lNum = 0
        If i <= UBound(sParts) Then sNum = CLng(Val(sParts(i)))
        If i <= UBound(lParts) Then lNum = CLng(Val(lParts(i)))
        
        If sNum > lNum Then
            IsNewerVersion = True
            Exit Function
        ElseIf sNum < lNum Then
            IsNewerVersion = False
            Exit Function
        End If
    Next i
    
    On Error GoTo 0
End Function

Public Sub PerformLanUpdate(ByVal updateSource As String, ByVal serverVer As String, Optional ByVal bSilent As Boolean = False)
    On Error GoTo InstallErr
    
    Dim fso As Object, wsh As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    
    Dim stagedDir As String
    stagedDir = Environ("APPDATA") & "\KangatangGuard\staged_update"
    If Not fso.FolderExists(stagedDir) Then
        fso.CreateFolder stagedDir
    End If
    
    ' Sao chep tap tin tu Hub vao bo dem tam staged_update
    Dim srcXlam As String, srcPs1 As String
    Dim targetXlam As String, targetPs1 As String
    srcXlam = updateSource & "\KangatangGuard.xlam"
    srcPs1 = updateSource & "\Kangatang_FolderScanner.ps1"
    targetXlam = stagedDir & "\KangatangGuard.xlam"
    targetPs1 = stagedDir & "\Kangatang_FolderScanner.ps1"
    
    ' Phong thu triet de: Go bo thuoc tinh Read-Only neu ton tai va xoa file cu
    If fso.FileExists(targetXlam) Then
        On Error Resume Next
        fso.GetFile(targetXlam).Attributes = 0
        fso.DeleteFile targetXlam, True
        On Error GoTo InstallErr
    End If
    If fso.FileExists(targetPs1) Then
        On Error Resume Next
        fso.GetFile(targetPs1).Attributes = 0
        fso.DeleteFile targetPs1, True
        On Error GoTo InstallErr
    End If
    
    ' Sao chep va chuan hoa thuoc tinh ve Normal/Archive (khong co Read-Only)
    If fso.FileExists(srcXlam) Then
        fso.CopyFile srcXlam, targetXlam, True
        On Error Resume Next
        If fso.FileExists(targetXlam) Then fso.GetFile(targetXlam).Attributes = 0
        On Error GoTo InstallErr
    End If
    If fso.FileExists(srcPs1) Then
        fso.CopyFile srcPs1, targetPs1, True
        On Error Resume Next
        If fso.FileExists(targetPs1) Then fso.GetFile(targetPs1).Attributes = 0
        On Error GoTo InstallErr
    End If
    
    ' Tao script ap dung cap nhat ngoai tien trinh (Out-of-Process Updater)
    Dim updaterBat As String
    updaterBat = stagedDir & "\apply_update.cmd"
    
    Dim xlStartDir As String, addInsDir As String, guardDir As String
    xlStartDir = Environ("APPDATA") & "\Microsoft\Excel\XLSTART"
    addInsDir = Environ("APPDATA") & "\Microsoft\AddIns"
    guardDir = Environ("APPDATA") & "\KangatangGuard"
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(updaterBat, True)
    ' v3.11.0: Polling nhanh 2 giay, tu dong don zombie process sau 60s, assert AccessVBOM
    ts.WriteLine "@echo off"
    ts.WriteLine "setlocal"
    ts.WriteLine "set ""LOG=" & guardDir & "\update_apply.log"""
    ts.WriteLine "echo [%date% %time%] Staged v" & serverVer & ". Waiting for all Excel processes to exit... >> ""%LOG%"""
    ts.WriteLine "set /a WAITED=0"
    ts.WriteLine ":WAIT_EXCEL"
    ts.WriteLine "tasklist /FI ""IMAGENAME eq EXCEL.EXE"" /NH 2>nul | find /I ""EXCEL.EXE"" >nul"
    ts.WriteLine "if errorlevel 1 goto APPLY"
    ts.WriteLine "set /a WAITED+=2"
    ts.WriteLine "if %WAITED% GEQ 60 goto FORCE_CLOSE"
    ts.WriteLine "ping -n 3 127.0.0.1 >nul"
    ts.WriteLine "goto WAIT_EXCEL"
    ts.WriteLine ":FORCE_CLOSE"
    ts.WriteLine "echo [%date% %time%] Force-closing orphan Excel processes... >> ""%LOG%"""
    ts.WriteLine "taskkill /F /IM EXCEL.EXE >nul 2>&1"
    ts.WriteLine "ping -n 3 127.0.0.1 >nul"
    ts.WriteLine ":APPLY"
    ts.WriteLine "ping -n 2 127.0.0.1 >nul"
    ts.WriteLine "attrib -r """ & xlStartDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "attrib -r """ & addInsDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "attrib -r """ & stagedDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "attrib -r """ & guardDir & "\Kangatang_FolderScanner.ps1"" >nul 2>&1"
    ts.WriteLine "copy /y """ & stagedDir & "\KangatangGuard.xlam"" """ & xlStartDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "if errorlevel 1 goto FAIL"
    ts.WriteLine "fc /b """ & stagedDir & "\KangatangGuard.xlam"" """ & xlStartDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "if errorlevel 1 goto FAIL"
    ts.WriteLine "if exist """ & addInsDir & "\KangatangGuard.xlam"" copy /y """ & stagedDir & "\KangatangGuard.xlam"" """ & addInsDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "if exist """ & stagedDir & "\Kangatang_FolderScanner.ps1"" copy /y """ & stagedDir & "\Kangatang_FolderScanner.ps1"" """ & guardDir & "\Kangatang_FolderScanner.ps1"" >nul 2>&1"
    ts.WriteLine "reg add ""HKCU\Software\KangatangGuard"" /v ""InstalledVersion"" /t REG_SZ /d """ & serverVer & """ /f >nul 2>&1"
    ts.WriteLine "reg add ""HKCU\Software\KangatangGuard"" /v ""ScannerScript"" /t REG_SZ /d """ & guardDir & "\Kangatang_FolderScanner.ps1"" /f >nul 2>&1"
    ts.WriteLine "reg add ""HKCU\Software\Microsoft\Office\16.0\Excel\Security"" /v ""AccessVBOM"" /t REG_DWORD /d 1 /f >nul 2>&1"
    ts.WriteLine "reg add ""HKCU\Software\Microsoft\Office\15.0\Excel\Security"" /v ""AccessVBOM"" /t REG_DWORD /d 1 /f >nul 2>&1"
    ts.WriteLine "reg add ""HKCU\Software\Microsoft\Office\14.0\Excel\Security"" /v ""AccessVBOM"" /t REG_DWORD /d 1 /f >nul 2>&1"
    ts.WriteLine "reg delete ""HKCU\Software\Microsoft\Office\16.0\Excel\Add-in Manager"" /v ""KangatangGuard.xlam"" /f >nul 2>&1"
    ts.WriteLine "reg delete ""HKCU\Software\Microsoft\Office\15.0\Excel\Add-in Manager"" /v ""KangatangGuard.xlam"" /f >nul 2>&1"
    ts.WriteLine "reg delete ""HKCU\Software\Microsoft\Office\14.0\Excel\Add-in Manager"" /v ""KangatangGuard.xlam"" /f >nul 2>&1"
    ts.WriteLine "echo [%date% %time%] SUCCESS: applied v" & serverVer & " >> ""%LOG%"""
    ts.WriteLine "exit /b 0"
    ts.WriteLine ":FAIL"
    ts.WriteLine "echo [%date% %time%] FAIL: copy/verify failed, InstalledVersion NOT changed >> ""%LOG%"""
    ts.WriteLine "exit /b 1"
    ts.Close
    Set ts = Nothing
    
    ' Chay updater script ngam
    wsh.Run Chr(34) & updaterBat & Chr(34), 0, False
    
    WriteLog "[UPDATE_STAGED] v" & serverVer & " from Hub: " & updateSource & " (will apply after Excel exits)"
    
    If Not bSilent Then
        Dim askClose As VbMsgBoxResult
        askClose = MsgBoxW(Uni("\u0110\u00e3 t\u1ea3i b\u1ea3n c\u1eadp nh\u1eadt v") & serverVer & Uni(" th\u00e0nh c\u00f4ng!" & vbCrLf & vbCrLf & _
                               "B\u1ea1n c\u00f3 mu\u1ed1n \u0110\u00d3NG EXCEL NGAY B\u00c2Y GI\u1edc \u0111\u1ec3 ho\u00e0n t\u1ea5t c\u1eadp nh\u1eadt kh\u00f4ng?" & vbCrLf & _
                               "(Ch\u1ecdn 'Yes' \u0111\u1ec3 \u00e1p d\u1ee5ng ngay, ho\u1eb7c 'No' \u0111\u1ec3 t\u1ef1 \u0111\u1ed9ng c\u1eadp nh\u1eadt sau khi b\u1ea1n l\u00e0m vi\u1ec7c xong)."), _
                           vbYesNo + vbQuestion, Uni("KangatangGuard - C\u1eadp nh\u1eadt th\u00e0nh c\u00f4ng"))
        If askClose = vbYes Then
            Application.Quit
        End If
    End If
            
    Set fso = Nothing
    Set wsh = Nothing
    Exit Sub
    
InstallErr:
    MsgBoxW Uni("L\u1ed7i khi c\u00e0i \u0111\u1eb7t b\u1ea3n c\u1eadp nh\u1eadt: ") & Err.Description, vbCritical, Uni("KangatangGuard - L\u1ed7i c\u1eadp nh\u1eadt")
    On Error GoTo 0
End Sub

'### END_SECTION: modKangatangScanner ###


'### SECTION: modKangatangShield ###
'--- Standard Module: Runtime Shield v3.10.0 - signatures, quick check, hook neutralizer, startup quarantine, vaccine ---

Option Explicit

' ---------------------------------------------------------------------------
' CHU KY VIRUS (rut ra tu mau that trong kho cach ly 223.7 ngay 06/10/2026):
'   Module "Kangatang": Sub Auto_Open -> ThisWorkbook.SaveCopyAs Application.StartupPath & "\mypersonnel1.xls"
'                       Application.OnSheetActivate = "mypersonnel1.xls!allocated"
'   Sub allocated     -> ThisWorkbook.Sheets("Kangatang").Copy before:=ActiveWorkbook.Sheets(1)
'   Sheet "Kangatang" la Module-sheet Excel 5 (VeryHidden) -> chep sheet = chep luon ma doc.
' ---------------------------------------------------------------------------
Private Const SIG_NAME_PREFIXES As String = "kangatang|kangaatang"
Private Const SIG_NAME_CONTAINS As String = "mypersonnel|mypersonel"
Private Const SIG_CODE_STRONG   As String = "mypersonnel|mypersonel|""kangatang""|""kangaatang"""
Private Const SIG_HOOK_TERMS    As String = "mypersonnel|mypersonel|kangatang|kangaatang|!allocated"
Private Const VACCINE_NAMES     As String = "mypersonnel.xls|mypersonnel1.xls|mypersonel.xls|mypersonel1.xls"
Private Const NET_CACHE_MINUTES As Long = 5
Private Const PING_TIMEOUT_MS   As Long = 800

Private mShieldPending As Boolean
' True khi dang chay ngoai ngu canh su kien (OnTime) -> an toan de dong workbook
Public bShieldContext As Boolean
Private mVbomWarned As Boolean
Private mRandomized As Boolean
Private dicNetCache As Object
' v3.10.0 Two-phase purge (module-sheet Excel 5): trang thai cho luu o Registry
Public bPurgeDeferred As Boolean
Private Const PURGE_KEY As String = "PendingPurge"
Private Const PURGE_MAX_ATTEMPTS As Long = 5

' ===========================================================================
' SIGNATURE HELPERS
' ===========================================================================
Public Function IsVirusName(ByVal s As String) As Boolean
    Dim t As String, parts() As String, i As Long
    t = LCase$(Trim$(s))
    t = Replace(t, "'", "")
    If Left$(t, 2) = "~$" Then t = Mid$(t, 3)
    If Len(t) = 0 Then Exit Function
    If Left$(t, 14) = "kangatangguard" Then Exit Function
    parts = Split(SIG_NAME_PREFIXES, "|")
    For i = 0 To UBound(parts)
        If Left$(t, Len(parts(i))) = parts(i) Then
            IsVirusName = True
            Exit Function
        End If
    Next i
    parts = Split(SIG_NAME_CONTAINS, "|")
    For i = 0 To UBound(parts)
        If InStr(1, t, parts(i), vbBinaryCompare) > 0 Then
            IsVirusName = True
            Exit Function
        End If
    Next i
End Function

Public Function IsVirusReference(ByVal refersTo As String) As Boolean
    Dim t As String
    t = LCase$(refersTo)
    If Len(t) = 0 Then Exit Function
    IsVirusReference = (InStr(1, t, "kangatang", vbBinaryCompare) > 0) Or _
                       (InStr(1, t, "kangaatang", vbBinaryCompare) > 0) Or _
                       (InStr(1, t, "mypersonnel", vbBinaryCompare) > 0) Or _
                       (InStr(1, t, "mypersonel", vbBinaryCompare) > 0)
End Function

Public Function CodeHasIOC(ByVal code As String) As Boolean
    Dim t As String, parts() As String, i As Long
    If Len(code) = 0 Then Exit Function
    t = LCase$(code)
    parts = Split(SIG_CODE_STRONG, "|")
    For i = 0 To UBound(parts)
        If InStr(1, t, parts(i), vbBinaryCompare) > 0 Then
            CodeHasIOC = True
            Exit Function
        End If
    Next i
    ' Heuristic 1 (persistence): tu sao chep vao thu muc khoi dong Excel
    If InStr(1, t, "savecopyas", vbBinaryCompare) > 0 Then
        If InStr(1, t, "startuppath", vbBinaryCompare) > 0 Or InStr(1, t, "xlstart", vbBinaryCompare) > 0 Then
            CodeHasIOC = True
            Exit Function
        End If
    End If
    ' Heuristic 2 (propagation): hook OnSheetActivate + chep sheet sang workbook dang hoat dong
    If InStr(1, t, "onsheetactivate", vbBinaryCompare) > 0 Then
        If InStr(1, t, ".copy before:=activeworkbook", vbBinaryCompare) > 0 Then
            CodeHasIOC = True
        End If
    End If
End Function

Private Function IsVirusHook(ByVal v As String) As Boolean
    Dim t As String, parts() As String, i As Long
    t = LCase$(v)
    If Len(t) = 0 Then Exit Function
    parts = Split(SIG_HOOK_TERMS, "|")
    For i = 0 To UBound(parts)
        If InStr(1, t, parts(i), vbBinaryCompare) > 0 Then
            IsVirusHook = True
            Exit Function
        End If
    Next i
End Function

' ===========================================================================
' CRASH GUARD: dem module-sheet Excel 5 (TypeName = "Module").
' Da tai hien (06/10/2026): file .xls chua module-sheet, mo o che do TAT macro
' (AutomationSecurity=3 hoac Trust Center "Disable with notification"),
' chi can truy cap wb.VBProject la Excel crash (RPC_E_SERVERFAULT) hoac treo.
' Xoa module-sheet qua Sheets API van an toan -> sau khi xoa moi cham VBProject.
' Tra ve -1 neu khong xac dinh duoc (caller coi nhu KHONG an toan).
' ===========================================================================
Public Function LegacyModuleSheetCount(ByVal wb As Workbook) As Long
    On Error GoTo LmErr
    Dim sh As Object, n As Long
    For Each sh In wb.Sheets
        If TypeName(sh) = "Module" Then n = n + 1
    Next sh
    LegacyModuleSheetCount = n
    Exit Function
LmErr:
    LegacyModuleSheetCount = -1
End Function

' ===========================================================================
' QUICK CHECK: chi doc TEN sheet / component (< 5ms) - dung cho BeforeSave va SheetActivate
' ===========================================================================
Public Function QuickCheckWorkbook(ByVal wb As Workbook, ByRef reason As String) As Boolean
    On Error Resume Next
    reason = ""
    If wb Is Nothing Then Exit Function
    If wb Is ThisWorkbook Then Exit Function
    
    If IsVirusName(wb.Name) Then
        reason = "filename:" & wb.Name
        QuickCheckWorkbook = True
        Exit Function
    End If
    
    Dim sh As Object, nmx As String
    For Each sh In wb.Sheets
        nmx = ""
        nmx = sh.Name
        If IsVirusName(nmx) Then
            reason = "sheet:" & nmx
            QuickCheckWorkbook = True
            Exit Function
        End If
    Next sh
    Err.Clear
    
    ' Crash guard (v3.10.0): khong cham VBProject khi con module-sheet Excel 5
    If LegacyModuleSheetCount(wb) <> 0 Then Exit Function
    
    Dim vbProj As Object, comp As Object, prot As Long
    Set vbProj = Nothing
    Set vbProj = wb.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        Err.Clear
        Exit Function
    End If
    prot = -1
    prot = vbProj.Protection
    If prot <> 0 Then
        Err.Clear
        Exit Function
    End If
    For Each comp In vbProj.VBComponents
        nmx = ""
        nmx = comp.Name
        If IsVirusName(nmx) Then
            reason = "module:" & nmx
            QuickCheckWorkbook = True
            Exit Function
        End If
    Next comp
    Err.Clear
End Function

' ===========================================================================
' GO HOOK SU KIEN CU (Application.OnSheetActivate...) MA VIRUS DA CAI TRONG RAM
' Late-bound de khong loi bien dich neu thuoc tinh an bi go bo trong ban Excel tuong lai
' ===========================================================================
Public Function NeutralizeVirusHooks() As Long
    On Error Resume Next
    Dim app As Object, v As String
    Set app = Application
    
    v = ""
    v = app.OnSheetActivate
    If IsVirusHook(v) Then
        app.OnSheetActivate = ""
        WriteLog "[HOOK_REMOVED] Application.OnSheetActivate = " & v
        NeutralizeVirusHooks = NeutralizeVirusHooks + 1
    End If
    
    v = ""
    v = app.OnSheetDeactivate
    If IsVirusHook(v) Then
        app.OnSheetDeactivate = ""
        WriteLog "[HOOK_REMOVED] Application.OnSheetDeactivate = " & v
        NeutralizeVirusHooks = NeutralizeVirusHooks + 1
    End If
    
    v = ""
    v = app.OnWindow
    If IsVirusHook(v) Then
        app.OnWindow = ""
        WriteLog "[HOOK_REMOVED] Application.OnWindow = " & v
        NeutralizeVirusHooks = NeutralizeVirusHooks + 1
    End If
    Err.Clear
End Function

' ===========================================================================
' KIEM TRA TRE (DEBOUNCE) SAU SheetActivate / WorkbookActivate / WorkbookOpen
' ===========================================================================
Public Sub ScheduleShieldCheck()
    On Error Resume Next
    If mShieldPending Then Exit Sub
    If bIsFolderScanning Then Exit Sub
    mShieldPending = True
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!ShieldDeferredCheck"
    If Err.Number <> 0 Then
        mShieldPending = False
        Err.Clear
    End If
End Sub

Public Sub ShieldDeferredCheck()
    On Error GoTo ShieldErr
    mShieldPending = False
    If bIsFolderScanning Then Exit Sub
    
    Call NeutralizeVirusHooks
    
    Dim hits As New Collection
    Dim wb As Workbook, reason As String
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If QuickCheckWorkbook(wb, reason) Then
                WriteLog "[SHIELD_HIT] " & wb.FullName & " | " & reason
                hits.Add wb
            End If
        End If
    Next wb
    
    ' Xu ly sau vong lap (ScanWorkbook co the dong workbook -> khong sua collection khi dang duyet)
    Dim item As Variant
    bShieldContext = True
    For Each item In hits
        Call ScanWorkbook(item)
    Next item
    bShieldContext = False
    Exit Sub
    
ShieldErr:
    WriteLog "[SHIELD_ERROR] ShieldDeferredCheck: " & Err.Description
    mShieldPending = False
    bShieldContext = False
End Sub

' ===========================================================================
' QUET KHI KHOI DONG: workbook da mo truoc add-in (XLSTART, PERSONAL.XLSB) + add-in dang nap
' + don thu muc khoi dong + vaccine. Chi I/O cuc bo, khong cham mang.
' ===========================================================================
Public Sub ShieldStartupSweep()
    On Error GoTo SweepErr
    WriteLog "[SHIELD_STARTUP] Sweep started (v" & CURRENT_VERSION & ")"
    
    Call NeutralizeVirusHooks
    
    ' Dong va cach ly ngay cac workbook mang ten virus dang mo (mypersonnel*, kangatang*)
    Dim curWb As Workbook
    For Each curWb In Application.Workbooks
        If Not curWb Is ThisWorkbook Then
            If IsVirusName(curWb.Name) Then
                WriteLog "[SHIELD_STARTUP] Closing and quarantining virus carrier: " & curWb.FullName
                Call NeutralizeVirusHooks
                Call QuarantineLoadedStartupWorkbook(curWb)
            End If
        End If
    Next curWb
    
    Dim wbs As New Collection, addinWbs As New Collection
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then wbs.Add wb
    Next wb
    
    ' Add-in dang nap khong nam trong Application.Workbooks -> lay qua AddIns2 (Excel 2010+)
    Dim app As Object, ai As Object, aiWb As Workbook, aiName As String, aiOpen As Boolean
    Set app = Application
    On Error Resume Next
    For Each ai In app.AddIns2
        aiOpen = False
        aiOpen = ai.IsOpen
        aiName = ""
        aiName = ai.Name
        If aiOpen And Len(aiName) > 0 And UCase(aiName) <> UCase(ThisWorkbook.Name) Then
            Set aiWb = Nothing
            Set aiWb = Application.Workbooks(aiName)
            If Not aiWb Is Nothing Then addinWbs.Add aiWb
        End If
    Next ai
    Err.Clear
    On Error GoTo SweepErr
    
    Dim item As Variant, reason As String
    bShieldContext = True
    ' Workbook thuong / XLSTART / PERSONAL.XLSB: quet day du (ten + noi dung)
    For Each item In wbs
        Call ScanWorkbook(item)
    Next item
    ' Add-in cua ben thu 3: chi quet theo ten (tranh false-positive heuristic tren add-in hop le)
    For Each item In addinWbs
        If QuickCheckWorkbook(item, reason) Then
            WriteLog "[SHIELD_HIT_ADDIN] " & item.FullName & " | " & reason
            Call ScanWorkbook(item)
        End If
    Next item
    
    bShieldContext = False
    Call SweepStartupFolders
    Call ApplyStartupVaccine
    Call CleanDocumentRecoveryRegistry
    Call CleanAddInCollisions
    Call RestoreExcelClipboardAndUI
    WriteLog "[SHIELD_STARTUP] Sweep finished: " & wbs.Count & " workbook(s), " & addinWbs.Count & " add-in(s)"
    Exit Sub
    
SweepErr:
    bShieldContext = False
    Call RestoreExcelClipboardAndUI
    WriteLog "[SHIELD_ERROR] ShieldStartupSweep: " & Err.Description
End Sub

' ===========================================================================
' THU MUC KHOI DONG EXCEL
' ===========================================================================
Private Function GetStartupFolders() As Collection
    Dim col As New Collection, p As String
    On Error Resume Next
    p = ""
    p = Application.StartupPath
    If Len(p) > 0 Then col.Add LCase$(p), LCase$(p)
    p = ""
    p = Application.AltStartupPath
    If Len(p) > 0 Then col.Add LCase$(p), LCase$(p)
    p = ""
    p = Application.Path & "\XLSTART"
    If Len(p) > 8 Then col.Add LCase$(p), LCase$(p)
    Err.Clear
    Set GetStartupFolders = col
End Function

Public Function IsInStartupFolder(ByVal fullName As String) As Boolean
    Dim pos As Long, parent As String, f As Variant
    pos = InStrRev(fullName, "\")
    If pos <= 1 Then Exit Function
    parent = LCase$(Left$(fullName, pos - 1))
    For Each f In GetStartupFolders()
        If parent = CStr(f) Then
            IsInStartupFolder = True
            Exit Function
        End If
    Next f
End Function

' ===========================================================================
' CACH LY TEP NGUON VIRUS DANG MO TU XLSTART (mypersonnel1.xls)
' ===========================================================================
Public Sub QuarantineLoadedStartupWorkbook(ByVal wb As Workbook)
    On Error Resume Next
    Dim p As String
    p = wb.FullName
    Call NeutralizeVirusHooks
    
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = prevAlerts
    If Err.Number <> 0 Then
        WriteLog "[QUARANTINE_ERROR] Cannot close startup workbook " & p & ": " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    If QuarantineFile(p) Then
        WriteLog "[XLSTART_QUARANTINED] " & p
        Call ApplyStartupVaccine
        MsgBoxW Uni("\u0110\u00c3 C\u00c1CH LY NGU\u1ed2N L\u00c2Y VIRUS KANGATANG!" & vbCrLf & vbCrLf & "T\u1ec7p: ") & p & vbCrLf & vbCrLf & _
                Uni("T\u1ec7p ngu\u1ed3n virus trong th\u01b0 m\u1ee5c kh\u1edfi \u0111\u1ed9ng Excel (XLSTART) \u0111\u00e3 \u0111\u01b0\u1ee3c \u0111\u00f3ng v\u00e0 chuy\u1ec3n v\u00e0o kho c\u00e1ch ly." & vbCrLf & _
                    "H\u1ec7 th\u1ed1ng \u0111\u00e3 t\u1ea1o 'vaccine' ch\u1eb7n virus t\u00e1i t\u1ea1o t\u1ec7p n\u00e0y."), _
                vbExclamation, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - C\u00e1ch ly ngu\u1ed3n l\u00e2y")
    End If
End Sub

' ===========================================================================
' KHO CACH LY: Hub trung tam (neu ping duoc) hoac APPDATA cuc bo. Tra ve duong dan co dau "\" cuoi.
' ===========================================================================
Public Function GetQuarantineDir(ByRef isCentral As Boolean) As String
    On Error Resume Next
    Dim fso As Object, d As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    isCentral = False
    If IsUncReachable(CENTRAL_BACKUP_HUB) Then
        If fso.FolderExists(CENTRAL_BACKUP_HUB) Then
            GetQuarantineDir = CENTRAL_BACKUP_HUB & "\"
            isCentral = True
            Exit Function
        End If
    End If
    d = Environ("APPDATA") & "\" & LOG_SUBFOLDER
    If Not fso.FolderExists(d) Then fso.CreateFolder d
    d = d & "\Quarantine_Backup"
    If Not fso.FolderExists(d) Then fso.CreateFolder d
    If fso.FolderExists(d) Then GetQuarantineDir = d & "\"
    Err.Clear
End Function

' Sao chep tep vao kho cach ly (duoi .quarantine de khong ai mo nham), xac minh, roi moi xoa ban goc.
Public Function QuarantineFile(ByVal srcPath As String) As Boolean
    On Error GoTo QErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcPath) Then Exit Function
    
    If Not mRandomized Then
        Randomize
        mRandomized = True
    End If
    
    Dim isCentral As Boolean, destDir As String, destPath As String, pc As String
    destDir = GetQuarantineDir(isCentral)
    If Len(destDir) = 0 Then Exit Function
    pc = Environ("COMPUTERNAME")
    If Len(pc) = 0 Then pc = "UNKNOWN_PC"
    destPath = destDir & fso.GetFileName(srcPath) & "_" & pc & "_" & Format(Now, "yyyyMMdd_HHmmss") & "_" & CStr(Int(Rnd * 9000) + 1000) & ".quarantine"
    
    fso.CopyFile srcPath, destPath, False
    If Not fso.FileExists(destPath) Then Exit Function
    
    On Error Resume Next
    fso.GetFile(srcPath).Attributes = 0
    fso.DeleteFile srcPath, True
    If fso.FileExists(srcPath) Then
        WriteLog "[QUARANTINE_WARN] Copied to " & destPath & " but could not remove original: " & srcPath & " | " & Err.Description
        Err.Clear
        Exit Function
    End If
    WriteLog "[QUARANTINE] " & srcPath & " -> " & destPath
    QuarantineFile = True
    Exit Function
    
QErr:
    WriteLog "[QUARANTINE_ERROR] " & srcPath & " | " & Err.Description
End Function

' Don cac tep mang ten virus (mypersonnel*.xls, Kangatang*.xls, ~$mypersonnel1.xls) chua mo trong thu muc khoi dong
Public Sub SweepStartupFolders()
    On Error Resume Next
    Dim fso As Object, f As Variant, fl As Object, p As Variant
    Dim found As Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In GetStartupFolders()
        If fso.FolderExists(CStr(f)) Then
            Set found = New Collection
            For Each fl In fso.GetFolder(CStr(f)).Files
                Dim nmLower As String
                nmLower = LCase$(fl.Name)
                If nmLower <> LCase$(ADDIN_NAME) And nmLower <> LCase$("~$" & ADDIN_NAME) And nmLower <> LCase$(ThisWorkbook.Name) Then
                    If IsVirusName(fso.GetBaseName(fl.Name)) Then found.Add fl.Path
                End If
            Next fl
            For Each p In found
                If Not IsWorkbookOpenByPath(CStr(p)) Then
                    If QuarantineFile(CStr(p)) Then WriteLog "[XLSTART_QUARANTINED] " & p
                Else
                    ' v3.11.0: Neu dang mo trong Excel: tim va dong workbook roi cach ly ngay lap tuc
                    Dim openWb As Workbook
                    For Each openWb In Application.Workbooks
                        If LCase$(openWb.FullName) = LCase$(CStr(p)) Or LCase$(openWb.Name) = LCase$(fso.GetFileName(CStr(p))) Then
                            Call NeutralizeVirusHooks
                            Call QuarantineLoadedStartupWorkbook(openWb)
                            Exit For
                        End If
                    Next openWb
                End If
            Next p
        End If
    Next f
    Err.Clear
End Sub

Private Function IsWorkbookOpenByPath(ByVal fullPath As String) As Boolean
    On Error Resume Next
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = LCase$(fullPath) Then
            IsWorkbookOpenByPath = True
            Exit Function
        End If
    Next wb
End Function

' ===========================================================================
' VACCINE: tao THU MUC an trung ten tep virus trong XLSTART -> SaveCopyAs cua virus that bai
' (Excel bo qua thu muc con trong XLSTART khi khoi dong)
' ===========================================================================
Public Sub ApplyStartupVaccine()
    On Error Resume Next
    Dim fso As Object, sp As String, names() As String, i As Long, target As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    sp = ""
    sp = Application.StartupPath
    If Len(sp) = 0 Then Exit Sub
    If Not fso.FolderExists(sp) Then Exit Sub
    names = Split(VACCINE_NAMES, "|")
    For i = 0 To UBound(names)
        target = sp & "\" & names(i)
        If Not fso.FileExists(target) And Not fso.FolderExists(target) Then
            fso.CreateFolder target
            If fso.FolderExists(target) Then
                fso.GetFolder(target).Attributes = 2 ' Hidden
                WriteLog "[VACCINE] Created blocker folder: " & target
            End If
        End If
        Err.Clear
    Next i
End Sub

' ===========================================================================
' VACCINE v3.10.0: Xoa bo trung lap Add-in va khoa OPEN gay xung dot khoi dong
' ===========================================================================
Public Sub CleanAddInCollisions()
    On Error Resume Next
    Dim reg As Object, fso As Object
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const HKEY_CURRENT_USER = &H80000001
    Dim officeVers As Variant, ver As Variant
    officeVers = Array("16.0", "15.0", "14.0")
    
    ' 1. Xoa tep trung lap trong %APPDATA%\Microsoft\AddIns de khong bi nap 2 lan
    Dim addInsPath As String, dupFile As String
    addInsPath = Environ("APPDATA") & "\Microsoft\AddIns"
    If fso.FolderExists(addInsPath) Then
        dupFile = addInsPath & "\" & ADDIN_NAME
        If fso.FileExists(dupFile) Then
            fso.GetFile(dupFile).Attributes = 0
            fso.DeleteFile dupFile, True
            WriteLog "[COLLISION_FIX] Removed duplicate in AddIns folder: " & dupFile
        End If
    End If
    
    If reg Is Nothing Then Exit Sub
    
    For Each ver In officeVers
        Dim optPath As String, valNames As Variant, valTypes As Variant, vn As Variant
        optPath = "Software\Microsoft\Office\" & ver & "\Excel\Options"
        
        ' 2. Xoa cac khoa OPEN* tro toi KangatangGuard.xlam
        reg.EnumValues HKEY_CURRENT_USER, optPath, valNames, valTypes
        If IsArray(valNames) Then
            For Each vn In valNames
                If UCase$(Left$(CStr(vn), 4)) = "OPEN" Then
                    Dim strVal As String
                    strVal = ""
                    reg.GetStringValue HKEY_CURRENT_USER, optPath, CStr(vn), strVal
                    If InStr(1, strVal, "KangatangGuard.xlam", vbTextCompare) > 0 Then
                        reg.DeleteValue HKEY_CURRENT_USER, optPath, CStr(vn)
                        WriteLog "[COLLISION_FIX] Removed registry entry " & optPath & "\" & CStr(vn) & " (" & strVal & ")"
                    End If
                End If
            Next vn
        End If
        
        ' 3. Xoa AddInLoadTimes neu co
        Dim loadTimesPath As String
        loadTimesPath = "Software\Microsoft\Office\" & ver & "\Excel\AddInLoadTimes"
        reg.DeleteValue HKEY_CURRENT_USER, loadTimesPath, ADDIN_NAME
        reg.DeleteValue HKEY_CURRENT_USER, loadTimesPath, "KangatangGuard"
        
        ' 4. Xoa khoi Resiliency\DisabledItems neu Office vo tinh danh dau vo hieu hoa
        Dim disPath As String, disVals As Variant, dv As Variant
        disPath = "Software\Microsoft\Office\" & ver & "\Excel\Resiliency\DisabledItems"
        reg.EnumValues HKEY_CURRENT_USER, disPath, disVals
        If IsArray(disVals) Then
            For Each dv In disVals
                Dim binVal As Variant
                reg.GetBinaryValue HKEY_CURRENT_USER, disPath, CStr(dv), binVal
                If IsArray(binVal) Then
                    Dim textRepr As String, i As Long
                    textRepr = ""
                    For i = 0 To UBound(binVal) Step 2
                        If binVal(i) > 31 And binVal(i) < 127 Then
                            textRepr = textRepr & Chr(binVal(i))
                        End If
                    Next i
                    If InStr(1, textRepr, "KangatangGuard", vbTextCompare) > 0 Then
                        reg.DeleteValue HKEY_CURRENT_USER, disPath, CStr(dv)
                        WriteLog "[COLLISION_FIX] Removed from DisabledItems: " & disPath & "\" & CStr(dv)
                    End If
                End If
            Next dv
        End If
        
        ' 5. (v3.11.0) Xoa khoi Add-in Manager neu tro toi KangatangGuard hoac UNC path cu
        Dim aimPath As String, aimVals As Variant, av As Variant, avTypes As Variant
        aimPath = "Software\Microsoft\Office\" & ver & "\Excel\Add-in Manager"
        reg.EnumValues HKEY_CURRENT_USER, aimPath, aimVals, avTypes
        If IsArray(aimVals) Then
            For Each av In aimVals
                Dim aimStr As String
                aimStr = CStr(av)
                If InStr(1, aimStr, "KangatangGuard", vbTextCompare) > 0 Or InStr(1, aimStr, "addin_kangatang", vbTextCompare) > 0 Then
                    reg.DeleteValue HKEY_CURRENT_USER, aimPath, aimStr
                    WriteLog "[COLLISION_FIX] Removed from Add-in Manager: " & aimStr
                End If
            Next av
        End If
        
        ' 6. (v3.11.0) Dam bao AccessVBOM = 1 va AllowNetworkLocations = 1 trong HKCU Security
        Dim secPath As String
        secPath = "Software\Microsoft\Office\" & ver & "\Excel\Security"
        reg.SetDWORDValue HKEY_CURRENT_USER, secPath, "AccessVBOM", 1
        reg.SetDWORDValue HKEY_CURRENT_USER, secPath, "AllowNetworkLocations", 1
    Next ver
    
    ' 7. (v3.11.0) Go bo cac Add-in trung lap trong Application.AddIns (chi cho phep duy nhat XLSTART)
    Dim oAi As Object
    On Error Resume Next
    For Each oAi In Application.AddIns
        If InStr(1, oAi.Name, "KangatangGuard", vbTextCompare) > 0 Or InStr(1, oAi.FullName, "KangatangGuard", vbTextCompare) > 0 Then
            If InStr(1, oAi.FullName, "XLSTART", vbTextCompare) = 0 Then
                oAi.Installed = False
                WriteLog "[COLLISION_FIX] Uninstalled duplicate Add-in from Application.AddIns: " & oAi.FullName
            End If
        End If
    Next oAi
    Err.Clear
End Sub

' ===========================================================================
' KIEM TRA MAY CHU UNC CO PHAN HOI (ping WMI timeout 800ms, cache 5 phut)
' Chan treo Excel 20-60s khi fso.FolderExists cham UNC luc mat LAN
' ===========================================================================
Public Function IsUncReachable(ByVal uncPath As String) As Boolean
    On Error GoTo NetErr
    If Left$(uncPath, 2) <> "\\" Then
        IsUncReachable = True
        Exit Function
    End If
    Dim host As String, p As Long
    host = Mid$(uncPath, 3)
    p = InStr(host, "\")
    If p > 0 Then host = Left$(host, p - 1)
    If Len(host) = 0 Then Exit Function
    
    If dicNetCache Is Nothing Then Set dicNetCache = CreateObject("Scripting.Dictionary")
    Dim entry As Variant
    If dicNetCache.Exists(host) Then
        entry = dicNetCache(host)
        If DateDiff("n", entry(0), Now) < NET_CACHE_MINUTES Then
            IsUncReachable = entry(1)
            Exit Function
        End If
    End If
    
    Dim res As Boolean, svc As Object, rs As Object, it As Object, sc As Variant
    Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set rs = svc.ExecQuery("SELECT StatusCode FROM Win32_PingStatus WHERE Address='" & host & "' AND Timeout=" & PING_TIMEOUT_MS)
    For Each it In rs
        sc = it.StatusCode
        If Not IsNull(sc) Then
            If CLng(sc) = 0 Then res = True
        End If
    Next it
    
    dicNetCache(host) = Array(Now, res)
    If Not res Then WriteLog "[NET] Host unreachable (ping): " & host
    IsUncReachable = res
    Exit Function
    
NetErr:
    WriteLog "[NET_ERROR] Ping check failed for " & uncPath & ": " & Err.Description
    IsUncReachable = False
End Function

' ===========================================================================
' KHONG TRUY CAP DUOC VBProject (AccessVBOM = 0): khong im lang
' ===========================================================================
Public Sub ReportVbomBlocked(ByVal wb As Workbook)
    On Error Resume Next
    WriteLog "[NO_VBOM] Cannot access VBProject of " & wb.FullName & " - only sheets/names were checked."
    If mVbomWarned Then Exit Sub
    mVbomWarned = True
    
    Dim wsh As Object, ver As String, policyVal As Variant, byPolicy As Boolean
    Set wsh = CreateObject("WScript.Shell")
    ver = Application.Version
    policyVal = Empty
    policyVal = wsh.RegRead("HKCU\Software\Policies\Microsoft\Office\" & ver & "\Excel\Security\AccessVBOM")
    Err.Clear
    byPolicy = Not IsEmpty(policyVal)
    If Not byPolicy Then
        wsh.RegWrite "HKCU\Software\Microsoft\Office\" & ver & "\Excel\Security\AccessVBOM", 1, "REG_DWORD"
        If Err.Number = 0 Then WriteLog "[NO_VBOM] Re-enabled AccessVBOM in HKCU (effective after Excel restart)"
        Err.Clear
    Else
        WriteLog "[NO_VBOM] AccessVBOM is enforced by Group Policy - contact IT"
    End If
    
    ' Canh bao toi da 1 lan/ngay
    Dim lastWarn As String
    lastWarn = ""
    lastWarn = wsh.RegRead("HKCU\Software\KangatangGuard\LastVbomWarn")
    Err.Clear
    If lastWarn = Format(Date, "yyyy-mm-dd") Then Exit Sub
    wsh.RegWrite "HKCU\Software\KangatangGuard\LastVbomWarn", Format(Date, "yyyy-mm-dd"), "REG_SZ"
    
    MsgBoxW Uni("KangatangGuard KH\u00d4NG TH\u1ec2 ki\u1ec3m tra m\u00e3 VBA v\u00ec Excel \u0111ang ch\u1eb7n quy\u1ec1n 'Trust access to the VBA project object model'." & vbCrLf & vbCrLf & _
                "Add-in v\u1eabn ki\u1ec3m tra sheet \u1ea9n c\u1ee7a virus nh\u01b0ng c\u00f3 th\u1ec3 b\u1ecf s\u00f3t bi\u1ebfn th\u1ec3." & vbCrLf & _
                "H\u1ec7 th\u1ed1ng \u0111\u00e3 t\u1ef1 b\u1eadt l\u1ea1i quy\u1ec1n n\u00e0y - vui l\u00f2ng kh\u1edfi \u0111\u1ed9ng l\u1ea1i Excel." & vbCrLf & _
                "N\u1ebfu th\u00f4ng b\u00e1o l\u1eb7p l\u1ea1i, vui l\u00f2ng li\u00ean h\u1ec7 IT (c\u00f3 th\u1ec3 do ch\u00ednh s\u00e1ch GPO)."), _
            vbExclamation, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - C\u1ea3nh b\u00e1o b\u1ea3o m\u1eadt")
End Sub

' ===========================================================================
' LAM SACH PHAU THUAT: chi xoa procedure chua IOC, giu nguyen macro hop le
' ===========================================================================
Public Function RemoveInfectedProcedures(ByVal comp As Object) As Long
    On Error GoTo RipErr
    Dim cm As Object
    Set cm = comp.CodeModule
    Dim total As Long
    total = cm.CountOfLines
    If total = 0 Then Exit Function
    If Not CodeHasIOC(cm.Lines(1, total)) Then Exit Function
    
    Dim ln As Long, pName As String, pKind As Long, sLine As Long, nLines As Long, body As String
    ln = cm.CountOfDeclarationLines + 1
    Do While ln <= cm.CountOfLines
        pKind = 0
        pName = cm.ProcOfLine(ln, pKind)
        If Len(pName) = 0 Then
            ln = ln + 1
        Else
            sLine = cm.ProcStartLine(pName, pKind)
            nLines = cm.ProcCountLines(pName, pKind)
            If nLines <= 0 Or sLine + nLines <= ln Then
                ln = ln + 1
            Else
                body = cm.Lines(sLine, nLines)
                If CodeHasIOC(body) Then
                    cm.DeleteLines sLine, nLines
                    RemoveInfectedProcedures = RemoveInfectedProcedures + 1
                    ln = sLine
                Else
                    ln = sLine + nLines
                End If
            End If
        End If
    Loop
    Exit Function
    
RipErr:
    WriteLog "[CLEAN_WARN] RemoveInfectedProcedures(" & comp.Name & "): " & Err.Description
End Function

Public Function CountProcedures(ByVal comp As Object) As Long
    On Error GoTo CpErr
    Dim cm As Object, ln As Long, pName As String, pKind As Long, sLine As Long, nLines As Long
    Set cm = comp.CodeModule
    ln = cm.CountOfDeclarationLines + 1
    Do While ln <= cm.CountOfLines
        pKind = 0
        pName = cm.ProcOfLine(ln, pKind)
        If Len(pName) = 0 Then
            ln = ln + 1
        Else
            CountProcedures = CountProcedures + 1
            sLine = cm.ProcStartLine(pName, pKind)
            nLines = cm.ProcCountLines(pName, pKind)
            If nLines <= 0 Or sLine + nLines <= ln Then
                ln = ln + 1
            Else
                ln = sLine + nLines
            End If
        End If
    Loop
    Exit Function
CpErr:
    CountProcedures = -1
End Function

' Xoa sheet an toan: kiem tra bao ve cau truc, giu it nhat 1 sheet hien thi, co lap loi
Public Function DeleteSheetSafe(ByVal wb As Workbook, ByVal idx As Long) As Boolean
    On Error Resume Next
    Dim sh As Object, nm As String, s As Object, v As Long, visibleOthers As Long
    Set sh = wb.Sheets(idx)
    If sh Is Nothing Then Exit Function
    nm = sh.Name
    
    Dim isProtected As Boolean
    isProtected = False
    isProtected = wb.ProtectStructure
    If isProtected Then
        Call TryUnprotectWorkbook(wb)
        isProtected = wb.ProtectStructure
    End If
    If isProtected Then
        WriteLog "[CLEAN_WARN] Workbook structure is protected, cannot delete sheet " & nm & " in " & wb.Name
        Exit Function
    End If
    
    Call TryUnprotectSheet(sh)
    
    For Each s In wb.Sheets
        If Not s Is sh Then
            v = 0
            v = s.Visible
            If v = -1 Then visibleOthers = visibleOthers + 1
        End If
    Next s
    If visibleOthers = 0 Then
        WriteLog "[CLEAN_WARN] Sheet " & nm & " is the only visible sheet, skipped in " & wb.Name
        Exit Function
    End If
    
    Dim prevAlerts As Boolean, delErr As Long, delMsg As String
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Err.Clear
    sh.Visible = -1
    Err.Clear
    sh.Delete
    delErr = Err.Number
    delMsg = Err.Description
    Err.Clear
    Application.DisplayAlerts = prevAlerts
    
    If delErr <> 0 Then
        WriteLog "[CLEAN_WARN] Cannot delete sheet " & nm & " in " & wb.Name & ": " & delMsg
        Exit Function
    End If
    DeleteSheetSafe = True
End Function

' ===========================================================================
' TWO-PHASE PURGE CHO MODULE-SHEET EXCEL 5 (v3.10.0)
' Bang chung thuc nghiem 06/10/2026 tren mau that:
'   - Xoa module-sheet tu COM ben ngoai: OK, khong anh huong gi.
'   - Xoa module-sheet tu ma VBA: call stack bi huy (RPC_E_SERVERFAULT) va
'     TOAN BO bien toan cuc cua project GOI LENH bi reset (project khac khong anh huong).
' => Khong xoa trong luong su kien. Phase 1 (OnTime rieng) xoa sheet - chap nhan bi reset;
'    Phase 2 (OnTime rieng) tai kich hoat Guard, lam sach phan con lai, luu, thong bao.
'    Trang thai cho duoc luu trong Registry (song sot qua reset).
' ===========================================================================

Public Function VirusModuleSheetCount(ByVal wb As Workbook) As Long
    On Error GoTo VmErr
    Dim sh As Object, n As Long
    For Each sh In wb.Sheets
        If TypeName(sh) = "Module" Then
            If IsVirusName(sh.Name) Then n = n + 1
        End If
    Next sh
    VirusModuleSheetCount = n
    Exit Function
VmErr:
    VirusModuleSheetCount = 0
End Function

Public Function IsPurgePending(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    Dim fn As String, ts As String
    fn = GetSetting(LOG_SUBFOLDER, PURGE_KEY, "FullName", "")
    If Len(fn) = 0 Then Exit Function
    If StrComp(fn, wb.FullName, vbTextCompare) <> 0 Then Exit Function
    ts = GetSetting(LOG_SUBFOLDER, PURGE_KEY, "Ts", "")
    If Len(ts) = 0 Then Exit Function
    ' Het han sau 2 phut (phong truong hop Phase 2 khong bao gio chay)
    IsPurgePending = (DateDiff("s", CDate(ts), Now) < 120)
End Function

Public Sub DeferModuleSheetPurge(ByVal wb As Workbook, ByVal bSaveAfter As Boolean)
    On Error Resume Next
    bPurgeDeferred = True
    If IsPurgePending(wb) Then Exit Sub
    SaveSetting LOG_SUBFOLDER, PURGE_KEY, "FullName", wb.FullName
    SaveSetting LOG_SUBFOLDER, PURGE_KEY, "Save", IIf(bSaveAfter, "1", "0")
    SaveSetting LOG_SUBFOLDER, PURGE_KEY, "Attempts", "0"
    SaveSetting LOG_SUBFOLDER, PURGE_KEY, "Ts", CStr(Now)
    WriteLog "[PURGE_SCHEDULED] " & wb.FullName & " - legacy module sheet will be removed outside the event stack."
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!ShieldPurgeModuleSheets"
    Application.OnTime Now + TimeSerial(0, 0, 3), "'" & ThisWorkbook.Name & "'!ShieldAfterPurge"
End Sub

Private Function GetPendingPurgeWorkbook() As Workbook
    On Error Resume Next
    Dim fn As String, w As Workbook
    fn = GetSetting(LOG_SUBFOLDER, PURGE_KEY, "FullName", "")
    If Len(fn) = 0 Then Exit Function
    For Each w In Application.Workbooks
        If StrComp(w.FullName, fn, vbTextCompare) = 0 Then
            Set GetPendingPurgeWorkbook = w
            Exit Function
        End If
    Next w
End Function

Private Sub ClearPendingPurge()
    On Error Resume Next
    DeleteSetting LOG_SUBFOLDER, PURGE_KEY
End Sub

' PHASE 1 (OnTime): xoa module-sheet virus. CANH BAO: lenh Delete se huy call stack nay
' va reset project add-in -> moi dong sau Delete co the khong chay. Phase 2 lo phan con lai.
Public Sub ShieldPurgeModuleSheets()
    On Error Resume Next
    Dim wb As Workbook, i As Long, sh As Object, s As Object, nm As String, visibleOthers As Long
    Set wb = GetPendingPurgeWorkbook()
    If wb Is Nothing Then Exit Sub
    Call NeutralizeVirusHooks
    If wb.ProtectStructure Then
        Call TryUnprotectWorkbook(wb)
    End If
    If wb.ProtectStructure Then
        WriteLog "[PURGE_FAIL] Workbook structure is protected: " & wb.FullName
        Exit Sub
    End If
    For i = wb.Sheets.Count To 1 Step -1
        Set sh = wb.Sheets(i)
        If TypeName(sh) = "Module" Then
            nm = sh.Name
            If IsVirusName(nm) Then
                visibleOthers = 0
                For Each s In wb.Sheets
                    If Not s Is sh Then
                        If s.Visible = -1 Then visibleOthers = visibleOthers + 1
                    End If
                Next s
                If visibleOthers = 0 Then
                    WriteLog "[PURGE_FAIL] No other visible sheet, cannot delete '" & nm & "' in " & wb.FullName
                    Exit Sub
                End If
                
                ' Thu thap mau ma doc truoc khi xoa module-sheet (v3.11.0)
                Dim rawModCode As String
                rawModCode = ""
                On Error Resume Next
                If Not wb.VBProject Is Nothing Then
                    Dim mComp As Object
                    Set mComp = wb.VBProject.VBComponents(nm)
                    If Not mComp Is Nothing Then
                        If Not mComp.CodeModule Is Nothing Then
                            If mComp.CodeModule.CountOfLines > 0 Then
                                rawModCode = mComp.CodeModule.Lines(1, mComp.CodeModule.CountOfLines)
                            End If
                        End If
                    End If
                End If
                Err.Clear
                If Len(rawModCode) = 0 Then
                    rawModCode = "[Legacy Excel 5.0 Module Sheet: " & nm & "]"
                End If
                Call CollectThreatSample(wb, nm, rawModCode, "LegacyModuleSheet")
                
                WriteLog "[PURGE] Removing legacy module sheet '" & nm & "' from " & wb.FullName & " (VBA project reset expected)"
                Application.DisplayAlerts = False
                sh.Visible = -1
                sh.Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next i
End Sub

' PHASE 2 (OnTime): tai kich hoat Guard sau reset, lam not (VBA component / name), luu, thong bao.
Public Sub ShieldAfterPurge()
    On Error GoTo AfterErr
    ' Khoi phuc trang thai ung dung co the bi bo do khi Phase 1 bi huy giua chung
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Call EnsureGuardAlive
    
    Dim wb As Workbook, attempts As Long, bSave As Boolean, ok As Boolean, fn As String
    Set wb = GetPendingPurgeWorkbook()
    If wb Is Nothing Then
        ClearPendingPurge
        Exit Sub
    End If
    fn = wb.FullName
    attempts = CLng(Val(GetSetting(LOG_SUBFOLDER, PURGE_KEY, "Attempts", "0"))) + 1
    SaveSetting LOG_SUBFOLDER, PURGE_KEY, "Attempts", CStr(attempts)
    
    If VirusModuleSheetCount(wb) > 0 Then
        If attempts < PURGE_MAX_ATTEMPTS Then
            ' Con module-sheet (vd "Kangatang (2)") -> lap lai chu trinh
            SaveSetting LOG_SUBFOLDER, PURGE_KEY, "Ts", CStr(Now)
            Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!ShieldPurgeModuleSheets"
            Application.OnTime Now + TimeSerial(0, 0, 3), "'" & ThisWorkbook.Name & "'!ShieldAfterPurge"
            Exit Sub
        End If
        WriteLog "[PURGE_FAIL] " & fn & " still has a legacy module sheet after " & attempts & " attempts."
        ClearPendingPurge
        If Application.Visible Then
            MsgBoxW Uni("C\u1ea2NH B\u00c1O: Ph\u00e1t hi\u1ec7n virus nh\u01b0ng KH\u00d4NG TH\u1ec2 l\u00e0m s\u1ea1ch ho\u00e0n to\u00e0n t\u1ef1 \u0111\u1ed9ng!" & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & _
                    Uni("Vui l\u00f2ng ch\u1ea1y Chay_Diet_Virus_Ngoai.bat ho\u1eb7c li\u00ean h\u1ec7 IT."), _
                vbCritical, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - L\u1ed7i")
        End If
        Exit Sub
    End If
    
    bSave = (GetSetting(LOG_SUBFOLDER, PURGE_KEY, "Save", "0") = "1")
    ClearPendingPurge
    WriteLog "[PURGE_DONE] Legacy module sheet removed: " & fn
    
    ' Module-sheet da het -> VBProject an toan; lam not va LUU (bat buoc luu vi sheet da bi xoa)
    bPurgeDeferred = False
    ok = CleanInfectedWorkbook(wb, bSave, True)
    If ok Then
        WriteLog "[CLEANED] " & fn & " (two-phase purge)"
    Else
        WriteLog "[ERROR] Clean failed after purge: " & fn
    End If
    
    If Application.Visible Then
        If ok Then
            MsgBoxW Uni("C\u1ea2NH B\u00c1O: PH\u00c1T HI\u1ec6N VIRUS KANGATANG!" & vbCrLf & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & vbCrLf & _
                    Uni("(D\u1ea1ng virus: Module-sheet Excel 5 \u1ea9n - \u0111\u00e3 g\u1ee1 an to\u00e0n ngo\u00e0i lu\u1ed3ng s\u1ef1 ki\u1ec7n)" & vbCrLf & _
                        "Tr\u1ea1ng th\u00e1i: \u0110\u00c3 TI\u00caU DI\u1ec6T TH\u00c0NH C\u00d4NG!" & vbCrLf & _
                        "B\u1ea3n sao l\u01b0u g\u1ed1c \u0111\u00e3 \u0111\u01b0\u1ee3c c\u00e1ch ly an to\u00e0n v\u1ec1 Kho M\u00e1y ch\u1ee7 LAN."), _
                    vbExclamation, Uni("KangatangGuard - Ti\u00eau di\u1ec7t th\u00e0nh c\u00f4ng")
        Else
            MsgBoxW Uni("C\u1ea2NH B\u00c1O: Ph\u00e1t hi\u1ec7n virus nh\u01b0ng KH\u00d4NG TH\u1ec2 l\u00e0m s\u1ea1ch ho\u00e0n to\u00e0n t\u1ef1 \u0111\u1ed9ng!" & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & _
                    Uni("Vui l\u00f2ng ch\u1ea1y Chay_Diet_Virus_Ngoai.bat ho\u1eb7c li\u00ean h\u1ec7 IT."), _
                vbCritical, Uni("KangatangGuard v") & CURRENT_VERSION & Uni(" - L\u1ed7i")
        End If
    End If
    
    If ok And bSave And Len(wb.Path) > 0 Then
        If Not IsInStartupFolder(fn) Then Call LaunchBackgroundScanner(wb.Path)
    End If
    Exit Sub
    
AfterErr:
    WriteLog "[PURGE_ERROR] ShieldAfterPurge: " & Err.Description
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

'### END_SECTION: modKangatangShield ###


'### SECTION: modLogger ###
'--- Standard Module: Ghi log kiem toan ra file ---

Option Explicit

Private Const LOG_FOLDER As String = "KangatangGuard"

Public Sub WriteLog(ByVal msg As String)
    On Error Resume Next
    
    Dim logDir As String
    logDir = Environ("APPDATA") & "\" & LOG_FOLDER
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(logDir) Then
        fso.CreateFolder logDir
    End If
    
    Dim logFile As String
    logFile = logDir & "\scan_log_" & Format(Now, "yyyyMMdd") & ".txt"
    
    Dim fNum As Integer
    fNum = FreeFile
    Open logFile For Append As #fNum
    Print #fNum, Format(Now, "yyyy-MM-dd HH:mm:ss") & " | " & msg
    Close #fNum
    
    Set fso = Nothing
    On Error GoTo 0
End Sub

'### END_SECTION: modLogger ###


'### SECTION: modThreatCollector ###
'--- Standard Module: Thu thap va luu tru mau ma doc Excel macro ---

Option Explicit

' Duong dan thu muc mau ma doc
Private Const LOCAL_SAMPLE_DIR As String = "D:\7. AI tools\kangatang\Threat_Samples"
Private Const LOCAL_BASE_DIR As String = "D:\7. AI tools\kangatang"
Private Const LAN_SAMPLE_DIR As String = "\\192.168.223.176\KangatangGuard_Hub\Threat_Samples"
Private Const LAN_BASE_DIR As String = "\\192.168.223.176\KangatangGuard_Hub"
Private Const FORBIDDEN_SERVER As String = "192.168.223.7"

' ===========================================================================
' XAC DINH THU MUC LUU MAU MA DOC
' Uu tien:
'   1. Cuc bo: D:\7. AI tools\kangatang\Threat_Samples\
'   2. Mang LAN may ca nhan: \\192.168.223.176\KangatangGuard_Hub\Threat_Samples\
'   3. Fallback: %APPDATA%\KangatangGuard\Threat_Samples\
' TUYET DOI KHONG LUU TREN 223.7!
' ===========================================================================
Private Function GetThreatSampleFolder() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim targetDir As String
    
    ' 1. Kiem tra thu muc cuc bo D:\7. AI tools\kangatang\Threat_Samples
    If fso.FolderExists(LOCAL_BASE_DIR) Then
        If Not fso.FolderExists(LOCAL_SAMPLE_DIR) Then
            fso.CreateFolder LOCAL_SAMPLE_DIR
        End If
        If fso.FolderExists(LOCAL_SAMPLE_DIR) Then
            GetThreatSampleFolder = LOCAL_SAMPLE_DIR
            Exit Function
        End If
    End If
    
    ' 2. Kiem tra mang LAN may ca nhan 192.168.223.176 (TUYET DOI KHONG LUU TREN 223.7!)
    If fso.FolderExists(LAN_BASE_DIR) Then
        If Not fso.FolderExists(LAN_SAMPLE_DIR) Then
            fso.CreateFolder LAN_SAMPLE_DIR
        End If
        If fso.FolderExists(LAN_SAMPLE_DIR) Then
            ' Chan tuyet doi may chu 223.7
            If InStr(1, LAN_SAMPLE_DIR, FORBIDDEN_SERVER, vbTextCompare) = 0 Then
                GetThreatSampleFolder = LAN_SAMPLE_DIR
                Exit Function
            End If
        End If
    End If
    
    ' 3. Fallback cuc bo: %APPDATA%\KangatangGuard\Threat_Samples
    Dim appData As String
    appData = Environ("APPDATA")
    If Len(appData) > 0 Then
        targetDir = appData & "\KangatangGuard"
        If Not fso.FolderExists(targetDir) Then
            fso.CreateFolder targetDir
        End If
        targetDir = targetDir & "\Threat_Samples"
        If Not fso.FolderExists(targetDir) Then
            fso.CreateFolder targetDir
        End If
        If fso.FolderExists(targetDir) Then
            GetThreatSampleFolder = targetDir
            Exit Function
        End If
    End If
    
    GetThreatSampleFolder = ""
End Function

' ===========================================================================
' TINH HASH MD5 DUA TREN .NET COM HOAC FALLBACK CHUOI AN TOAN
' ===========================================================================
Public Function ComputeThreatHash(ByVal rawText As String) As String
    On Error GoTo HashFallback
    Dim enc As Object, md5 As Object
    Set enc = CreateObject("System.Text.UTF8Encoding")
    Set md5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    Dim bytes As Variant, hashBytes As Variant
    bytes = enc.GetBytes_4(rawText)
    hashBytes = md5.ComputeHash_2((bytes))
    
    Dim hexStr As String, i As Long, b As Byte, h As String
    hexStr = ""
    For i = 1 To LenB(hashBytes)
        b = AscB(MidB(hashBytes, i, 1))
        h = Hex(b)
        If Len(h) = 1 Then h = "0" & h
        hexStr = hexStr & LCase$(h)
    Next i
    
    If Len(hexStr) > 0 Then
        ComputeThreatHash = hexStr
        Exit Function
    End If

HashFallback:
    ' Fallback hash neu COM .NET bi chan tren he thong
    Dim hashVal As Long, codeLen As Long
    codeLen = Len(rawText)
    hashVal = 5381
    For i = 1 To codeLen
        hashVal = ((hashVal * 33) Xor AscW(Mid$(rawText, i, 1))) And &H7FFFFFFF
    Next i
    ComputeThreatHash = "fb_" & Hex(codeLen) & "_" & Hex(hashVal)
End Function

' ===========================================================================
' ESCAPE CHUOI AN TOAN CHO DINH DANG JSON
' ===========================================================================
Private Function EscapeJsonString(ByVal s As String) As String
    Dim res As String
    res = Replace(s, "\", "\\")
    res = Replace(res, """", "\""")
    res = Replace(res, vbCrLf, "\r\n")
    res = Replace(res, vbCr, "\r")
    res = Replace(res, vbLf, "\n")
    res = Replace(res, vbTab, "\t")
    EscapeJsonString = res
End Function

' ===========================================================================
' THU THAP VA LUU TRU MAU MA DOC (Threat Collector)
' ===========================================================================
Public Sub CollectThreatSample(ByVal wb As Workbook, ByVal compOrSheetName As String, ByVal rawCode As String, ByVal threatType As String)
    On Error GoTo CollectErr
    
    ' Bo qua neu khong co du lieu de thu thap
    If Len(Trim$(rawCode)) = 0 And Len(compOrSheetName) = 0 Then Exit Sub
    
    Dim targetFolder As String
    targetFolder = GetThreatSampleFolder()
    If Len(targetFolder) = 0 Then
        WriteLog "[THREAT_COLLECT_ERR] Cannot determine threat sample directory."
        Exit Sub
    End If
    
    ' Chan bao mat nghiem ngat: TUYET DOI KHONG LUU TREN 223.7
    If InStr(1, targetFolder, FORBIDDEN_SERVER, vbTextCompare) > 0 Then
        WriteLog "[THREAT_COLLECT_ABORT] Target folder contains forbidden server 223.7: " & targetFolder
        Exit Sub
    End If
    
    ' Tinh ma hash cua raw code
    Dim codeHash As String
    codeHash = ComputeThreatHash(rawCode)
    
    ' Kiem tra trung lap: Neu file hash da ton tai trong thu muc, khong ghi de de tiet kiem I/O
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sampleFileName As String
    sampleFileName = targetFolder & "\" & codeHash & ".sample.json"
    
    If fso.FileExists(sampleFileName) Then
        WriteLog "[THREAT_COLLECT] Sample already exists (hash: " & codeHash & ", comp: " & compOrSheetName & ")"
        Exit Sub
    End If
    
    ' Thong tin ngu canh
    Dim wbPath As String
    wbPath = "Unknown"
    If Not wb Is Nothing Then
        On Error Resume Next
        wbPath = wb.FullName
        Err.Clear
        On Error GoTo CollectErr
    End If
    
    Dim compName As String
    compName = Environ("COMPUTERNAME")
    If Len(compName) = 0 Then compName = "Unknown"
    
    Dim timeStr As String
    timeStr = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Xay dung noi dung file mau defanged dang JSON
    Dim jsonContent As String
    jsonContent = "{" & vbCrLf & _
        "  ""timestamp"": """ & timeStr & """," & vbCrLf & _
        "  ""computer_name"": """ & EscapeJsonString(compName) & """," & vbCrLf & _
        "  ""file_path"": """ & EscapeJsonString(wbPath) & """," & vbCrLf & _
        "  ""component_name"": """ & EscapeJsonString(compOrSheetName) & """," & vbCrLf & _
        "  ""threat_type"": """ & EscapeJsonString(threatType) & """," & vbCrLf & _
        "  ""code_length"": " & Len(rawCode) & "," & vbCrLf & _
        "  ""code_hash"": """ & codeHash & """," & vbCrLf & _
        "  ""raw_code"": """ & EscapeJsonString(rawCode) & """" & vbCrLf & _
        "}"
        
    ' Ghi file UTF-8 qua ADODB.Stream, fallback sang FSO
    Dim writeOk As Boolean
    writeOk = False
    
    On Error Resume Next
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    If Not stm Is Nothing Then
        stm.Type = 2 ' adTypeText
        stm.Charset = "utf-8"
        stm.Open
        stm.WriteText jsonContent
        stm.SaveToFile sampleFileName, 2 ' adSaveCreateOverWrite
        stm.Close
        writeOk = (Err.Number = 0)
        Err.Clear
    End If
    
    If Not writeOk Then
        ' Fallback FSO
        Dim ts As Object
        Set ts = fso.CreateTextFile(sampleFileName, True, True) ' Unicode
        If Not ts Is Nothing Then
            ts.Write jsonContent
            ts.Close
            writeOk = True
        End If
        Err.Clear
    End If
    On Error GoTo CollectErr
    
    If writeOk Then
        WriteLog "[THREAT_COLLECT] Saved threat sample " & codeHash & " (" & threatType & " - " & compOrSheetName & ") to " & sampleFileName
    Else
        WriteLog "[THREAT_COLLECT_ERR] Failed to write sample file " & sampleFileName
    End If
    Exit Sub

CollectErr:
    WriteLog "[THREAT_COLLECT_ERR] Exception during collection: " & Err.Description
End Sub

'### END_SECTION: modThreatCollector ###

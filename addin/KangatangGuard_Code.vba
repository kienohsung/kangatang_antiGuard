'==============================================================================
' KangatangGuard - Excel Add-in Diet Virus Macro Kangatang
' Phien ban: v3.8.0 (Dedicated LAN File Server & Dual-Mirror Architecture)
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
    Call CleanDocumentRecoveryRegistry
    Call InitializeGuard
    ' Kiem tra cap nhat tu May chu LAN ngam sau 3 giay (Async / Non-blocking)
    ' Giup Excel mo tuc thi trong 0.05 giay ma khong bi cham tre du chi 1 ms
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:03"), "CheckForLanUpdatesSilent"
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
    On Error GoTo 0
End Sub

Private Sub xlApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' v3.4.0: Chi quet neu file chua duoc quet gan day (ScanCache chong lag)
    On Error Resume Next
    If bIsFolderScanning Then Exit Sub
    If Not IsScanCacheExpired(Wb) Then Exit Sub
    Call ScanWorkbook(Wb)
    On Error GoTo 0
End Sub

Private Sub xlApp_NewWorkbook(ByVal Wb As Workbook)
    On Error Resume Next
    If bIsFolderScanning Then Exit Sub
    Call ScanWorkbook(Wb)
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

Public Const CURRENT_VERSION As String = "3.8.2"
Private Const DEFAULT_HUB_PRIMARY As String = "\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang"
Private Const DEFAULT_HUB_BACKUP  As String = "\\192.168.223.176\KangatangGuard_Hub"

Private Const ADDIN_NAME As String = "KangatangGuard.xlam"
Private Const BACKUP_FOLDER_NAME As String = "_Backup_Kangatang"
Private Const LOG_SUBFOLDER As String = "KangatangGuard"
Private Const SCAN_CACHE_MINUTES As Long = 5

' --- Danh sach tu khoa virus de quet ---
Private Const VIRUS_PATTERN_NAMES As String = "Kangatang,Kangaatang,Kanga,mypersonnel"

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
        title = Uni("KangatangGuard v3.8.0")
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
' Ham khoi tao va huy event handler
' ===========================================================================
Public Sub InitializeGuard()
    On Error Resume Next
    Set dicScanCache = CreateObject("Scripting.Dictionary")
    bIsFolderScanning = False
    
    Set oAppEvents = New clsAppEvents
    Set oAppEvents.xlApp = Application
    
    Call CreateMenu
    
    WriteLog "KangatangGuard v3.8.0 da khoi dong thanh cong."
    
    Dim openWb As Workbook
    For Each openWb In Application.Workbooks
        Call ScanWorkbook(openWb)
    Next openWb
    
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

' ===========================================================================
' Ham tao va xoa menu tren thanh cong cu (Add-ins tab)
' ===========================================================================
Private Sub CreateMenu()
    On Error Resume Next
    Call RemoveMenu
    
    Dim cmdBar As CommandBar
    Set cmdBar = Application.CommandBars("Worksheet Menu Bar")
    
    Dim menuItem As CommandBarPopup
    Set menuItem = cmdBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    menuItem.Caption = "KangatangGuard"
    menuItem.Tag = "KangatangGuardMenu"
    
    ' Nut 1: Quet tep hien tai
    Dim btn1 As CommandBarButton
    Set btn1 = menuItem.Controls.Add(Type:=msoControlButton)
    btn1.Caption = Uni("Qu\u00e9t t\u1ec7p hi\u1ec7n t\u1ea1i")
    btn1.FaceId = 1088
    btn1.OnAction = "ScanActiveWorkbook"
    btn1.Tag = "KG_ScanCurrent"
    
    ' Nut 2: Quet thu muc moi (Out-of-process)
    Dim btn2 As CommandBarButton
    Set btn2 = menuItem.Controls.Add(Type:=msoControlButton)
    btn2.Caption = Uni("Qu\u00e9t th\u01b0 m\u1ee5c m\u1edbi...")
    btn2.FaceId = 23
    btn2.OnAction = "ScanFolderDialog"
    btn2.Tag = "KG_ScanFolder"
    
    ' Nut 2b: Tiep tuc phien quet truoc do (v3.6.0 Fast Resume)
    Dim btn2b As CommandBarButton
    Set btn2b = menuItem.Controls.Add(Type:=msoControlButton)
    btn2b.Caption = Uni("Ti\u1ebfp t\u1ee5c phi\u00ean qu\u00e9t tr\u01b0\u1edbc...")
    btn2b.FaceId = 38
    btn2b.OnAction = "ResumeScanDialog"
    btn2b.Tag = "KG_ResumeScan"
    
    ' Nut 3: Mo thu muc Nhat ky
    Dim btn3 As CommandBarButton
    Set btn3 = menuItem.Controls.Add(Type:=msoControlButton)
    btn3.Caption = Uni("M\u1edf th\u01b0 m\u1ee5c nh\u1eadt k\u00fd (Log)")
    btn3.FaceId = 40
    btn3.OnAction = "OpenLogFolder"
    btn3.Tag = "KG_OpenLog"
    
    ' Nut 3b: Kiem tra cap nhat tu May chu LAN (v3.8.0)
    Dim btn3b As CommandBarButton
    Set btn3b = menuItem.Controls.Add(Type:=msoControlButton)
    btn3b.Caption = Uni("Ki\u1ec3m tra c\u1eadp nh\u1eadt t\u1eeb M\u00e1y ch\u1ee7...")
    btn3b.FaceId = 463
    btn3b.OnAction = "CheckForLanUpdatesManual"
    btn3b.Tag = "KG_CheckUpdate"
    
    ' Nut 4: Thong tin
    Dim btn4 As CommandBarButton
    Set btn4 = menuItem.Controls.Add(Type:=msoControlButton)
    btn4.Caption = Uni("Th\u00f4ng tin KangatangGuard v3.8.0")
    btn4.FaceId = 487
    btn4.OnAction = "ShowAbout"
    btn4.Tag = "KG_About"
    
    On Error GoTo 0
End Sub

Private Sub RemoveMenu()
    On Error Resume Next
    Dim ctrl As CommandBarControl
    For Each ctrl In Application.CommandBars("Worksheet Menu Bar").Controls
        If ctrl.Tag = "KangatangGuardMenu" Then
            ctrl.Delete
        End If
    Next ctrl
    On Error GoTo 0
End Sub

' ===========================================================================
' Ham kiem tra whitelist
' ===========================================================================
Private Function IsWhitelisted(ByVal wb As Workbook) As Boolean
    IsWhitelisted = False
    On Error Resume Next
    
    If UCase(wb.Name) = UCase(ADDIN_NAME) Then
        IsWhitelisted = True
        Exit Function
    End If
    
    If LCase(Right(wb.Name, 5)) = ".xlam" Then
        IsWhitelisted = True
        Exit Function
    End If
    
    If UCase(wb.Name) = "PERSONAL.XLSB" Then
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
' HAM KIEM TRA MA DOC (INSPECTION ENGINE)
' ===========================================================================
Public Function CheckWorkbookInfection(ByVal wb As Workbook, ByRef detailMsg As String) As Boolean
    On Error GoTo CheckError
    CheckWorkbookInfection = False
    detailMsg = ""
    
    If IsWhitelisted(wb) Then Exit Function
    
    Dim virusKeywords() As String
    virusKeywords = Split(VIRUS_PATTERN_NAMES, ",")
    
    Dim vbProj As Object
    On Error Resume Next
    Set vbProj = wb.VBProject
    If Err.Number <> 0 Then
        WriteLog "[SKIP] " & wb.FullName & " - Khong truy cap duoc VBProject."
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo CheckError
    
    Dim comp As Object
    Dim compName As String
    Dim i As Long
    Dim kw As Long
    
    ' 1a. Quet ten Module
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents.Item(i)
        compName = comp.Name
        
        For kw = LBound(virusKeywords) To UBound(virusKeywords)
            If InStr(1, compName, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                detailMsg = detailMsg & Uni("  - Module \u0111\u1ed9c h\u1ea1i: ") & compName & vbCrLf
                CheckWorkbookInfection = True
                Exit For
            End If
        Next kw
    Next i
    
    ' 1b. Quet noi dung ma nguon ben trong moi component
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents.Item(i)
        On Error Resume Next
        If comp.CodeModule.CountOfLines > 0 Then
            Dim codeContent As String
            codeContent = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            
            For kw = LBound(virusKeywords) To UBound(virusKeywords)
                If InStr(1, codeContent, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                    If InStr(1, detailMsg, comp.Name, vbTextCompare) = 0 Then
                        detailMsg = detailMsg & Uni("  - M\u00e3 \u0111\u1ed9c trong: ") & comp.Name & vbCrLf
                        CheckWorkbookInfection = True
                    End If
                    Exit For
                End If
            Next kw
        End If
        On Error GoTo CheckError
    Next i
    
    ' 2. Quet Sheet an
    Dim sht As Object
    For i = wb.Sheets.Count To 1 Step -1
        Set sht = wb.Sheets(i)
        For kw = LBound(virusKeywords) To UBound(virusKeywords)
            If InStr(1, sht.Name, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                detailMsg = detailMsg & Uni("  - Sheet \u1ea9n \u0111\u1ed9c h\u1ea1i: ") & sht.Name & vbCrLf
                CheckWorkbookInfection = True
                Exit For
            End If
        Next kw
    Next i
    
    ' 3. Quet Hidden Named Ranges
    Dim nm As Name
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        For kw = LBound(virusKeywords) To UBound(virusKeywords)
            If InStr(1, nm.Name, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                detailMsg = detailMsg & Uni("  - Named Range \u0111\u1ed9c h\u1ea1i: ") & nm.Name & vbCrLf
                CheckWorkbookInfection = True
                Exit For
            End If
        Next kw
        
        On Error Resume Next
        If Not nm.Visible Then
            Dim refStr As String
            refStr = nm.RefersTo
            For kw = LBound(virusKeywords) To UBound(virusKeywords)
                If InStr(1, refStr, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                    detailMsg = detailMsg & Uni("  - Named Range \u1ea9n (tham chi\u1ebfu \u0111\u1ed9c): ") & nm.Name & vbCrLf
                    CheckWorkbookInfection = True
                    Exit For
                End If
            Next kw
        End If
        On Error GoTo CheckError
    Next i
    
    Exit Function
    
CheckError:
    WriteLog "[ERROR] Loi khi kiem tra " & wb.Name & ": " & Err.Description
    On Error GoTo 0
End Function

' ===========================================================================
' HAM QUET CHINH: Quet 1 Workbook khi mo/luu (Thuc thi cuc nhanh < 0.2s)
' ===========================================================================
Public Sub ScanWorkbook(ByVal wb As Workbook)
    On Error GoTo ScanError
    
    If IsWhitelisted(wb) Then Exit Sub
    
    Dim virusFound As Boolean
    Dim detailMsg As String
    virusFound = CheckWorkbookInfection(wb, detailMsg)
    
    UpdateScanCache wb
    
    If virusFound Then
        WriteLog "[DETECTED] " & wb.FullName & vbCrLf & detailMsg
        
        If wb.ReadOnly Then
            WriteLog "[READONLY_DETECTED] " & wb.FullName & " la file Read-Only."
            MsgBoxW Uni("C\u1ea2NH B\u00c1O: PH\u00c1T HI\u1ec6N VIRUS KANGATANG TRONG T\u1ec6P CH\u1ec8 \u0110\u1eccC (READ-ONLY)!" & vbCrLf & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & vbCrLf & _
                        Uni("Chi ti\u1ebft:") & vbCrLf & detailMsg & vbCrLf & _
                        Uni("L\u01b0u \u00fd: T\u1ec7p \u0111ang \u1edf ch\u1ebf \u0111\u1ed9 Ch\u1ec8 \u0111\u1ecdc (Read-Only) n\u00ean kh\u00f4ng th\u1ec3 t\u1ef1 \u0111\u1ed9ng ghi \u0111\u00e8." & vbCrLf & _
                        "Vui l\u00f2ng m\u1edf kh\u00f3a t\u1ec7p ho\u1eb7c ch\u1ecdn 'Save As' sang b\u1ea3n sao m\u1edbi."), _
                    vbCritical, Uni("KangatangGuard v3.6.0 - C\u1ea3nh b\u00e1o")
            Exit Sub
        End If
        
        Dim backupOK As Boolean
        backupOK = BackupBeforeClean(wb)
        
        If backupOK Then
            Dim cleanResult As Boolean
            cleanResult = CleanInfectedWorkbook(wb)
            
            If cleanResult Then
                WriteLog "[CLEANED] " & wb.FullName
                
                MsgBoxW Uni("C\u1ea2NH B\u00c1O: PH\u00c1T HI\u1ec6N VIRUS KANGATANG!" & vbCrLf & vbCrLf & _
                            "T\u1ec7p: ") & wb.Name & vbCrLf & vbCrLf & _
                            Uni("Chi ti\u1ebft:") & vbCrLf & detailMsg & vbCrLf & _
                            Uni("Tr\u1ea1ng th\u00e1i: \u0110\u00c3 TI\u00caU DI\u1ec6T TH\u00c0NH C\u00d4NG!" & vbCrLf & _
                            "B\u1ea3n sao l\u01b0u g\u1ed1c \u0111\u00e3 \u0111\u01b0\u1ee3c t\u1ea1o t\u1ea1i: ") & BACKUP_FOLDER_NAME & "\" & vbCrLf & vbCrLf & _
                            Uni("Ti\u1ebfn tr\u00ecnh qu\u00e9t ng\u1ea7m \u0111\u1ed9c l\u1eadp s\u1ebd T\u1ef0 \u0110\u1ed8NG QU\u00c9T TO\u00c0N B\u1ed8 TH\u01af M\u1ee4C ch\u1ee9a t\u1ec7p n\u00e0y m\u00e0 kh\u00f4ng l\u00e0m gi\u00e1n \u0111o\u1ea1n Excel c\u1ee7a b\u1ea1n!"), _
                        vbExclamation, Uni("KangatangGuard v3.6.0 - Th\u00e0nh c\u00f4ng")
                
                ' v3.6.0: Khoi chay tien trinh quet ngam doc lap (khong lam treo Excel)
                Dim parentDir As String
                parentDir = Left(wb.FullName, InStrRev(wb.FullName, "\"))
                If Len(parentDir) > 0 Then
                    Call LaunchBackgroundScanner(parentDir)
                End If
            Else
                WriteLog "[ERROR] Khong the lam sach: " & wb.FullName
                MsgBoxW Uni("C\u1ea2NH B\u00c1O: Ph\u00e1t hi\u1ec7n virus nh\u01b0ng KH\u00d4NG TH\u1ec2 l\u00e0m s\u1ea1ch t\u1ef1 \u0111\u1ed9ng!" & vbCrLf & _
                            "T\u1ec7p: ") & wb.Name & vbCrLf & _
                            Uni("Vui l\u00f2ng ch\u1ea1y Diet_Virus_Kangatang.bat \u0111\u1ec3 x\u1eed l\u00fd th\u1ee7 c\u00f4ng."), _
                        vbCritical, Uni("KangatangGuard v3.6.0 - L\u1ed7i")
            End If
        Else
            WriteLog "[ERROR] Khong the tao backup cho: " & wb.FullName
            MsgBoxW Uni("C\u1ea2NH B\u00c1O: Ph\u00e1t hi\u1ec7n virus nh\u01b0ng KH\u00d4NG TH\u1ec2 t\u1ea1o b\u1ea3n sao l\u01b0u (Backup)!" & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & _
                        Uni("\u0110\u1ec3 b\u1ea3o to\u00e0n d\u1eef li\u1ec7u, h\u1ec7 th\u1ed1ng kh\u00f4ng t\u1ef1 \u0111\u1ed9ng s\u1eeda t\u1ec7p khi ch\u01b0a sao l\u01b0u \u0111\u01b0\u1ee3c." & vbCrLf & _
                        "Vui l\u00f2ng sao l\u01b0u th\u1ee7 c\u00f4ng v\u00e0 ch\u1ea1y Diet_Virus_Kangatang.bat."), _
                    vbCritical, Uni("KangatangGuard v3.6.0 - C\u1ea3nh b\u00e1o an to\u00e0n")
        End If
    Else
        WriteLog "[SAFE] " & wb.FullName
    End If
    
    Exit Sub
    
ScanError:
    WriteLog "[ERROR] Loi khi quet " & wb.Name & ": " & Err.Description
    On Error GoTo 0
End Sub

' ===========================================================================
' HAM KHOI CHAY TIEN TRINH QUET NGAM DOC LAP (v3.6.0 - ASYNCHRONOUS WORKER)
' Chay ngoai tien trinh (Out-of-Process), thoat ngay trong 10ms
' Ho tro tham so bResume de tiep tuc phien quet truoc do
' ===========================================================================
Public Sub LaunchBackgroundScanner(ByVal folderPath As String, Optional ByVal bResume As Boolean = False)
    On Error GoTo LaunchErr
    
    ' VACCINE v3.5.4+: Cat bo toan bo dau \ o cuoi duong dan de tranh loi Command Line Escaping (\" trong Windows CLI)
    Dim cleanFolder As String
    cleanFolder = folderPath
    Do While Right(cleanFolder, 1) = "\" And Len(cleanFolder) > 0
        cleanFolder = Left(cleanFolder, Len(cleanFolder) - 1)
    Loop
    
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
    
    ' Kiem tra cuoi cung: Neu van khong tim thay file script thi thong bao ro rang
    If Len(scannerScript) = 0 Or Not fso.FileExists(scannerScript) Then
        Set fso = Nothing
        Set wsh = Nothing
        MsgBoxW Uni("Kh\u00f4ng t\u00ecm th\u1ea5y t\u1ec7p tr\u00ecnh qu\u00e9t ng\u1ea7m: Kangatang_FolderScanner.ps1!" & vbCrLf & vbCrLf & _
                    "Vui l\u00f2ng ch\u1ea1y l\u1ea1i t\u1ec7p C\u00e0i \u0111\u1eb7t Add-in ho\u1eb7c li\u00ean h\u1ec7 IT \u0111\u1ec3 \u0111\u01b0\u1ee3c h\u1ed7 tr\u1ee3."), _
                vbCritical, Uni("KangatangGuard v3.6.0 - Thi\u1ebfu t\u1ec7p h\u1ec7 th\u1ed1ng")
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
' HAM BACKUP: Tao ban sao file truoc khi lam sach (v3.5.4)
' Ho tro ten file Unicode khong bi loi Bad file name or number
' ===========================================================================
Public Function BackupBeforeClean(ByVal wb As Workbook) As Boolean
    On Error GoTo BackupError
    BackupBeforeClean = False
    
    Dim filePath As String
    filePath = wb.FullName
    
    If Len(filePath) = 0 Or InStrRev(filePath, "\") = 0 Then
        Exit Function
    End If
    
    Dim parentFolder As String
    parentFolder = Left(filePath, InStrRev(filePath, "\"))
    
    Dim backupDir As String
    backupDir = parentFolder & BACKUP_FOLDER_NAME & "\"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(backupDir) Then
        fso.CreateFolder backupDir
    End If
    
    Dim timestamp As String
    timestamp = Format(Now, "yyyyMMdd_HHmmss")
    
    Dim originalName As String
    originalName = wb.Name
    
    Dim baseName As String
    Dim extName As String
    baseName = fso.GetBaseName(originalName)
    extName = fso.GetExtensionName(originalName)
    
    Dim backupName As String
    backupName = baseName & "_backup_" & timestamp & "." & extName
    
    Dim destPath As String
    destPath = backupDir & backupName
    
    fso.CopyFile filePath, destPath, True
    Set fso = Nothing
    
    WriteLog "[BACKUP] Da tao ban sao: " & destPath
    BackupBeforeClean = True
    Exit Function
    
BackupError:
    WriteLog "[BACKUP_ERROR] " & Err.Description & " | File: " & filePath
    BackupBeforeClean = False
End Function

' ===========================================================================
' HAM LAM SACH: Xoa toan bo thanh phan nhiem virus
' ===========================================================================
Public Function CleanInfectedWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo CleanError
    CleanInfectedWorkbook = False
    
    Dim virusKeywords() As String
    virusKeywords = Split(VIRUS_PATTERN_NAMES, ",")
    
    Dim vbProj As Object
    Set vbProj = wb.VBProject
    
    Dim i As Long
    Dim kw As Long
    Dim comp As Object
    Dim cleaned As Boolean
    cleaned = False
    
    ' 1. Xoa cac VBA Module doc hai (chi xoa Module=1 va Class=2)
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents.Item(i)
        
        If comp.Type = 1 Or comp.Type = 2 Then
            For kw = LBound(virusKeywords) To UBound(virusKeywords)
                If InStr(1, comp.Name, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                    vbProj.VBComponents.Remove comp
                    cleaned = True
                    Exit For
                End If
            Next kw
        End If
    Next i
    
    ' 2. Xoa noi dung ma doc ben trong Document modules (ThisWorkbook, Sheet...)
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents.Item(i)
        On Error Resume Next
        If comp.CodeModule.CountOfLines > 0 Then
            Dim codeText As String
            codeText = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            
            For kw = LBound(virusKeywords) To UBound(virusKeywords)
                If InStr(1, codeText, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                    comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
                    cleaned = True
                    Exit For
                End If
            Next kw
        End If
        On Error GoTo CleanError
    Next i
    
    ' 3. Xoa Sheet an doc hai
    Application.DisplayAlerts = False
    For i = wb.Sheets.Count To 1 Step -1
        If wb.Sheets.Count > 1 Then
            Dim shtName As String
            shtName = wb.Sheets(i).Name
            
            For kw = LBound(virusKeywords) To UBound(virusKeywords)
                If InStr(1, shtName, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                    wb.Sheets(i).Visible = xlSheetVisible
                    wb.Sheets(i).Delete
                    cleaned = True
                    Exit For
                End If
            Next kw
        End If
    Next i
    Application.DisplayAlerts = True
    
    ' 4. Xoa Hidden Named Ranges doc hai
    Dim nm As Name
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        Dim shouldDelete As Boolean
        shouldDelete = False
        
        For kw = LBound(virusKeywords) To UBound(virusKeywords)
            If InStr(1, nm.Name, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                shouldDelete = True
                Exit For
            End If
        Next kw
        
        If Not shouldDelete Then
            On Error Resume Next
            If Not nm.Visible Then
                Dim refText As String
                refText = nm.RefersTo
                For kw = LBound(virusKeywords) To UBound(virusKeywords)
                    If InStr(1, refText, Trim(virusKeywords(kw)), vbTextCompare) > 0 Then
                        shouldDelete = True
                        Exit For
                    End If
                Next kw
            End If
            On Error GoTo CleanError
        End If
        
        If shouldDelete Then
            On Error Resume Next
            nm.Delete
            cleaned = True
            On Error GoTo CleanError
        End If
    Next i
    
    If cleaned Then
        wb.Save
    End If
    
    CleanInfectedWorkbook = True
    Exit Function
    
CleanError:
    Application.DisplayAlerts = True
    CleanInfectedWorkbook = False
End Function

' ===========================================================================
' HAM MENU: Quet file hien tai (goi tu menu)
' ===========================================================================
Public Sub ScanActiveWorkbook()
    If ActiveWorkbook Is Nothing Then
        MsgBoxW Uni("Kh\u00f4ng c\u00f3 t\u1ec7p Excel n\u00e0o \u0111ang m\u1edf."), vbInformation, Uni("KangatangGuard v3.6.0")
        Exit Sub
    End If
    
    Dim wbName As String
    wbName = ActiveWorkbook.Name
    
    Call ScanWorkbook(ActiveWorkbook)
    
    MsgBoxW Uni("\u0110\u00e3 ho\u00e0n th\u00e0nh qu\u00e9t t\u1ec7p: ") & wbName & vbCrLf & _
            Uni("Chi ti\u1ebft \u0111\u01b0\u1ee3c ghi t\u1ea1i th\u01b0 m\u1ee5c nh\u1eadt k\u00fd (Log)."), _
            vbInformation, Uni("KangatangGuard v3.6.0")
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
        Call LaunchBackgroundScanner(folderPath)
        
        ' Thong bao nhanh cho nguoi dung (Excel tiep tuc hoat dong ngay lap tuc)
        MsgBoxW Uni("\u0110\u00e3 kh\u1edfi ch\u1ea1y ti\u1ebfn tr\u00ecnh qu\u00e9t ng\u1ea7m \u0111\u1ed9c l\u1eadp cho th\u01b0 m\u1ee5c:" & vbCrLf & vbCrLf) & _
                folderPath & vbCrLf & vbCrLf & _
                Uni("Ti\u1ebfn tr\u00ecnh \u0111ang ch\u1ea1y tr\u00ean c\u1eeda s\u1ed5 ri\u00eang bi\u1ec7t v\u1edbi thanh ti\u1ebfn \u0111\u1ed9 % th\u1eddi gian th\u1ef1c." & vbCrLf & _
                    "B\u1ea1n c\u00f3 th\u1ec3 TI\u1ebeP T\u1ee4C L\u00c0M VI\u1ec6C tr\u00ean Excel b\u00ecnh th\u01b0\u1eddng m\u00e0 kh\u00f4ng lo b\u1ecb treo m\u00e1y!"), _
                vbInformation, Uni("KangatangGuard v3.6.0 - Ti\u1ebfn tr\u00ecnh ng\u1ea7m")
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
                vbInformation, Uni("KangatangGuard v3.6.0 - Ti\u1ebfp t\u1ee5c phi\u00ean")
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
                vbInformation, Uni("KangatangGuard v3.6.0 - Phi\u00ean \u0111\u00e3 ho\u00e0n t\u1ea5t")
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
    res = MsgBoxW(promptMsg, vbYesNo + vbQuestion, Uni("KangatangGuard v3.6.0 - Ti\u1ebfp t\u1ee5c phi\u00ean qu\u00e9t"))
    
    If res = vbYes Then
        Call LaunchBackgroundScanner(targetVal, bResume:=True)
        MsgBoxW Uni("\u0110\u00e3 k\u00edch ho\u1ea1t ti\u1ebfp t\u1ee5c phi\u00ean qu\u00e9t ng\u1ea7m tr\u00ean c\u1eeda s\u1ed5 ri\u00eang bi\u1ec7t." & vbCrLf & _
                    "B\u1ea1n c\u00f3 th\u1ec3 ti\u1ebfp t\u1ee5c l\u00e0m vi\u1ec7c tr\u00ean Excel b\u00ecnh th\u01b0\u1eddng!"), _
                vbInformation, Uni("KangatangGuard v3.6.0 - Ti\u1ebfp t\u1ee5c phi\u00ean")
    End If
    Exit Sub
    
ResumeErr:
    MsgBoxW Uni("L\u1ed7i khi \u0111\u1ecdc th\u00f4ng tin phi\u00ean qu\u00e9t: ") & Err.Description, vbCritical, Uni("KangatangGuard v3.6.0 - L\u1ed7i")
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
                    If InStr(1, strVal, "192.168.", vbTextCompare) > 0 Or _
                       InStr(1, strVal, "file_shared", vbTextCompare) > 0 Or _
                       InStr(1, strVal, "_Backup_Kangatang", vbTextCompare) > 0 Or _
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
' KIEM TRA VA DONG BO CAP NHAT TU TRUNG TAM LAN HUB (v3.8.0)
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
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Doc UpdateSource tu Registry HKCU\Software\KangatangGuard
    On Error Resume Next
    updateSource = wsh.RegRead("HKCU\Software\KangatangGuard\UpdateSource")
    On Error GoTo UpdateErr
    
    If Len(Trim(updateSource)) = 0 Then
        ' Uu tien May chu tep chuyen dung 223.7, sau do den may du phong 223.176
        If fso.FolderExists(DEFAULT_HUB_PRIMARY) Then
            updateSource = DEFAULT_HUB_PRIMARY
        ElseIf fso.FolderExists(DEFAULT_HUB_BACKUP) Then
            updateSource = DEFAULT_HUB_BACKUP
        End If
    End If
    
    If Len(Trim(updateSource)) = 0 Then
        If Not bSilent Then
            MsgBoxW Uni("Ch\u01b0a c\u1ea5u h\u00ecnh \u0111\u01b0\u1eddng d\u1eabn M\u00e1y ch\u1ee7 ph\u00e2n ph\u1ed1i (UpdateSource)!" & vbCrLf & vbCrLf & _
                        "Vui l\u00f2ng ch\u1ea1y file 'Install_Client_Kangatang.bat' t\u1eeb M\u00e1y ch\u1ee7 \u0111\u1ec3 \u0111\u0103ng k\u00fd."), _
                    vbExclamation, Uni("KangatangGuard v3.8.0 - C\u1eadp nh\u1eadt")
        End If
        Exit Sub
    End If
    
    ' Loai bo dau \ o cuoi neu co
    If Right(updateSource, 1) = "\" Then updateSource = Left(updateSource, Len(updateSource) - 1)
    
    versionFile = updateSource & "\version.json"
    
    ' Kiem tra tep version.json voi co che fallback
    If Not fso.FileExists(versionFile) Then
        If fso.FileExists(DEFAULT_HUB_PRIMARY & "\version.json") Then
            updateSource = DEFAULT_HUB_PRIMARY
            versionFile = updateSource & "\version.json"
        ElseIf fso.FileExists(DEFAULT_HUB_BACKUP & "\version.json") Then
            updateSource = DEFAULT_HUB_BACKUP
            versionFile = updateSource & "\version.json"
        End If
    End If
    
    If Not fso.FileExists(versionFile) Then
        If Not bSilent Then
            MsgBoxW Uni("Kh\u00f4ng th\u1ec3 k\u1ebft n\u1ed1i \u0111\u1ebfn M\u00e1y ch\u1ee7 LAN ho\u1eb7c kh\u00f4ng t\u00ecm th\u1ea5y th\u00f4ng tin phi\u00ean b\u1ea3n:" & vbCrLf & vbCrLf) & _
                    updateSource & vbCrLf & vbCrLf & _
                    Uni("Vui l\u00f2ng ki\u1ec3m tra k\u1ebft n\u1ed1i m\u1ea1ng LAN ho\u1eb7c m\u00e1y ch\u1ee7 c\u00f3 \u0111ang b\u1eadt kh\u00f4ng."), _
                    vbExclamation, Uni("KangatangGuard v3.8.0 - K\u1ebft n\u1ed1i th\u1ea5t b\u1ea1i")
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
            MsgBoxW Uni("Kh\u00f4ng th\u1ec3 \u0111\u1ecdc th\u00f4ng tin phi\u00ean b\u1ea3n t\u1eeb t\u1ec7p version.json tr\u00ean m\u00e1y ch\u1ee7."), vbExclamation, Uni("KangatangGuard v3.7.0")
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
        
        Dim ans As VbMsgBoxResult
        ans = MsgBoxW(promptMsg, vbYesNo + vbInformation, Uni("KangatangGuard - C\u00f3 phi\u00ean b\u1ea3n m\u1edbi v") & serverVer)
        If ans = vbYes Then
            Call PerformLanUpdate(updateSource, serverVer)
        End If
    Else
        If Not bSilent Then
            MsgBoxW Uni("B\u1ea1n \u0111ang s\u1eed d\u1ee5ng phi\u00ean b\u1ea3n M\u1edaI NH\u1ea4T (v") & CURRENT_VERSION & ")!" & vbCrLf & vbCrLf & _
                    Uni("M\u00e1y ch\u1ee7 ph\u00e2n ph\u1ed1i: ") & updateSource, _
                    vbInformation, Uni("KangatangGuard v3.8.0 - H\u1ec7 th\u1ed1ng \u0111\u00e3 c\u1eadp nh\u1eadt")
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
    If Not bSilent Then
        MsgBoxW Uni("L\u1ed7i khi ki\u1ec3m tra c\u1eadp nh\u1eadt: ") & Err.Description, vbCritical, Uni("KangatangGuard v3.8.0 - L\u1ed7i")
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

Public Sub PerformLanUpdate(ByVal updateSource As String, ByVal serverVer As String)
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
    srcXlam = updateSource & "\KangatangGuard.xlam"
    srcPs1 = updateSource & "\Kangatang_FolderScanner.ps1"
    
    If fso.FileExists(srcXlam) Then
        fso.CopyFile srcXlam, stagedDir & "\KangatangGuard.xlam", True
    End If
    If fso.FileExists(srcPs1) Then
        fso.CopyFile srcPs1, stagedDir & "\Kangatang_FolderScanner.ps1", True
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
    ts.WriteLine "@echo off"
    ts.WriteLine "timeout /t 2 /nobreak > nul"
    ts.WriteLine "attrib -r """ & xlStartDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "copy /y """ & stagedDir & "\KangatangGuard.xlam"" """ & xlStartDir & "\KangatangGuard.xlam"" > nul"
    ts.WriteLine "attrib +r """ & xlStartDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "attrib -r """ & addInsDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "copy /y """ & stagedDir & "\KangatangGuard.xlam"" """ & addInsDir & "\KangatangGuard.xlam"" > nul"
    ts.WriteLine "attrib +r """ & addInsDir & "\KangatangGuard.xlam"" >nul 2>&1"
    ts.WriteLine "copy /y """ & stagedDir & "\Kangatang_FolderScanner.ps1"" """ & guardDir & "\Kangatang_FolderScanner.ps1"" > nul"
    ts.WriteLine "reg add ""HKCU\Software\KangatangGuard"" /v ""InstalledVersion"" /t REG_SZ /d """ & serverVer & """ /f > nul"
    ts.Close
    Set ts = Nothing
    
    ' Chay updater script ngam
    wsh.Run Chr(34) & updaterBat & Chr(34), 0, False
    
    WriteLog "Da tai va kich hoat cap nhat len phien ban v" & serverVer & " tu Hub: " & updateSource
    
    MsgBoxW Uni("\u0110\u00e3 t\u1ea3i v\u00e0 k\u00edch ho\u1ea1t b\u1ea3n c\u1eadp nh\u1eadt v") & serverVer & Uni(" th\u00e0nh c\u00f4ng!" & vbCrLf & vbCrLf & _
                "B\u1ea3n m\u1edbi s\u1ebd \u0111\u01b0\u1ee3c \u00e1p d\u1ee5ng ho\u00e0n to\u00e0n khi b\u1ea1n kh\u1edfi \u0111\u1ed9ng l\u1ea1i Excel."), _
            vbInformation, Uni("KangatangGuard - C\u1eadp nh\u1eadt th\u00e0nh c\u00f4ng")
            
    Set fso = Nothing
    Set wsh = Nothing
    Exit Sub
    
InstallErr:
    MsgBoxW Uni("L\u1ed7i khi c\u00e0i \u0111\u1eb7t b\u1ea3n c\u1eadp nh\u1eadt: ") & Err.Description, vbCritical, Uni("KangatangGuard v3.8.0 - L\u1ed7i")
    On Error GoTo 0
End Sub

'### END_SECTION: modKangatangScanner ###


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

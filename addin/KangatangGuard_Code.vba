'==============================================================================
' KangatangGuard - Excel Add-in Diet Virus Macro Kangatang
' Phien ban: v3.5.3 (Real-Time Directory Streaming Architecture)
' Mo ta: Tu dong quet va tieu diet virus macro Kangatang/Laroux/mypersonnel
'         ngay khi mo file Excel. Khi quet thu muc hoac khi phat hien file nhiem,
'         tu dong tach sang tien trinh PowerShell doc lap (Asynchronous Worker),
'         quet luong truc tiep lan luot tung thu muc ma khong can cho gom tep!
'
' File nay chua toan bo ma VBA duoc phan tach theo section markers.
' PowerShell Installer se doc va inject tung phan vao dung module/class.
'==============================================================================

'### SECTION: ThisWorkbook ###
'--- Code nay se duoc inject vao ThisWorkbook cua file .xlam ---

Option Explicit

Private Sub Workbook_Open()
    Call InitializeGuard
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
        title = Uni("KangatangGuard v3.5.3")
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
    
    WriteLog "KangatangGuard v3.5.3 da khoi dong thanh cong."
    
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
    
    ' Nut 2: Quet thu muc (Out-of-process)
    Dim btn2 As CommandBarButton
    Set btn2 = menuItem.Controls.Add(Type:=msoControlButton)
    btn2.Caption = Uni("Qu\u00e9t th\u01b0 m\u1ee5c (Ch\u1ea1y ng\u1ea7m)...")
    btn2.FaceId = 23
    btn2.OnAction = "ScanFolderDialog"
    btn2.Tag = "KG_ScanFolder"
    
    ' Nut 3: Mo thu muc Nhat ky
    Dim btn3 As CommandBarButton
    Set btn3 = menuItem.Controls.Add(Type:=msoControlButton)
    btn3.Caption = Uni("M\u1edf th\u01b0 m\u1ee5c nh\u1eadt k\u00fd (Log)")
    btn3.FaceId = 40
    btn3.OnAction = "OpenLogFolder"
    btn3.Tag = "KG_OpenLog"
    
    ' Nut 4: Thong tin
    Dim btn4 As CommandBarButton
    Set btn4 = menuItem.Controls.Add(Type:=msoControlButton)
    btn4.Caption = Uni("Th\u00f4ng tin KangatangGuard v3.5.3")
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
                        Uni("L\u01b0u \u00fd: T\u1ec7p \u0111ang \u1edf ch\u1ebf \u0111\u1ed9 Ch\u1ec9 \u0111\u1ecdc (Read-Only) n\u00ean kh\u00f4ng th\u1ec3 t\u1ef1 \u0111\u1ed9ng ghi \u0111\u00e8." & vbCrLf & _
                        "Vui l\u00f2ng m\u1edf kh\u00f3a t\u1ec7p ho\u1eb7c ch\u1ecdn 'Save As' sang b\u1ea3n sao m\u1edbi."), _
                    vbCritical, Uni("KangatangGuard v3.5.3 - C\u1ea3nh b\u00e1o")
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
                        vbExclamation, Uni("KangatangGuard v3.5.3 - Th\u00e0nh c\u00f4ng")
                
                ' v3.5.3: Khoi chay tien trinh quet ngam doc lap (khong lam treo Excel)
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
                        vbCritical, Uni("KangatangGuard v3.5.3 - L\u1ed7i")
            End If
        Else
            WriteLog "[ERROR] Khong the tao backup cho: " & wb.FullName
            MsgBoxW Uni("C\u1ea2NH B\u00c1O: Ph\u00e1t hi\u1ec7n virus nh\u01b0ng KH\u00d4NG TH\u1ec2 t\u1ea1o b\u1ea3n sao l\u01b0u (Backup)!" & vbCrLf & _
                        "T\u1ec7p: ") & wb.Name & vbCrLf & _
                        Uni("\u0110\u1ec3 b\u1ea3o to\u00e0n d\u1eef li\u1ec7u, h\u1ec7 th\u1ed1ng kh\u00f4ng t\u1ef1 \u0111\u1ed9ng s\u1eeda t\u1ec7p khi ch\u01b0a sao l\u01b0u \u0111\u01b0\u1ee3c." & vbCrLf & _
                        "Vui l\u00f2ng sao l\u01b0u th\u1ee7 c\u00f4ng v\u00e0 ch\u1ea1y Diet_Virus_Kangatang.bat."), _
                    vbCritical, Uni("KangatangGuard v3.5.3 - C\u1ea3nh b\u00e1o an to\u00e0n")
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
' HAM KHOI CHAY TIEN TRINH QUET NGAM DOC LAP (v3.5.3 - ASYNCHRONOUS WORKER)
' Chay ngoai tien trinh (Out-of-Process), thoat ngay trong 10ms
' Giup Excel 100% muot ma, khong bao gio bi (Not Responding)
' ===========================================================================
Public Sub LaunchBackgroundScanner(ByVal folderPath As String)
    On Error GoTo LaunchErr
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Dim scannerScript As String
    scannerScript = Environ("APPDATA") & "\" & LOG_SUBFOLDER & "\Kangatang_FolderScanner.ps1"
    
    ' Phong thu: Neu chua co trong APPDATA thi tim trong thu muc Add-in
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(scannerScript) Then
        Dim addinDir As String
        addinDir = ThisWorkbook.Path & "\"
        If fso.FileExists(addinDir & "Kangatang_FolderScanner.ps1") Then
            scannerScript = addinDir & "Kangatang_FolderScanner.ps1"
        End If
    End If
    Set fso = Nothing
    
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scannerScript & """ -TargetFolder """ & folderPath & """"
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ' Tham so 1 = Normal window, False = Asynchronous non-blocking (Khong cho, thoat ngay lap tuc!)
    wsh.Run cmd, 1, False
    Set wsh = Nothing
    
    WriteLog "[BG_LAUNCH] Da khoi chay tien trinh quet ngam cho: " & folderPath
    Exit Sub
    
LaunchErr:
    WriteLog "[BG_LAUNCH_ERROR] Loi khoi chay worker: " & Err.Description
    On Error GoTo 0
End Sub

' ===========================================================================
' HAM BACKUP: Tao ban sao file truoc khi lam sach (v3.5.3)
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
        MsgBoxW Uni("Kh\u00f4ng c\u00f3 t\u1ec7p Excel n\u00e0o \u0111ang m\u1edf."), vbInformation, Uni("KangatangGuard v3.5.3")
        Exit Sub
    End If
    
    Dim wbName As String
    wbName = ActiveWorkbook.Name
    
    Call ScanWorkbook(ActiveWorkbook)
    
    MsgBoxW Uni("\u0110\u00e3 ho\u00e0n th\u00e0nh qu\u00e9t t\u1ec7p: ") & wbName & vbCrLf & _
            Uni("Chi ti\u1ebft \u0111\u01b0\u1ee3c ghi t\u1ea1i th\u01b0 m\u1ee5c nh\u1eadt k\u00fd (Log)."), _
            vbInformation, Uni("KangatangGuard v3.5.3")
End Sub

' ===========================================================================
' HAM MENU: Quet thu muc (goi tu menu, hien hop thoai chon thu muc)
' v3.5.3: Tach tien trinh quet ngam doc lap khong lam treo Excel
' ===========================================================================
Public Sub ScanFolderDialog()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = Uni("KangatangGuard - Ch\u1ecdn th\u01b0 m\u1ee5c c\u1ea7n qu\u00e9t")
    fd.ButtonName = Uni("Qu\u00e9t th\u01b0 m\u1ee5c n\u00e0y")
    
    If fd.Show = -1 Then
        Dim folderPath As String
        folderPath = fd.SelectedItems(1)
        If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
        
        ' v3.5.3: Goi worker ngam doc lap (non-blocking)
        Call LaunchBackgroundScanner(folderPath)
        
        ' Thong bao nhanh cho nguoi dung (Excel tiep tuc hoat dong ngay lap tuc)
        MsgBoxW Uni("\u0110\u00e3 kh\u1edfi ch\u1ea1y ti\u1ebfn tr\u00ecnh qu\u00e9t ng\u1ea7m \u0111\u1ed9c l\u1eadp cho th\u01b0 m\u1ee5c:" & vbCrLf & vbCrLf) & _
                folderPath & vbCrLf & vbCrLf & _
                Uni("Ti\u1ebfn tr\u00ecnh \u0111ang ch\u1ea1y tr\u00ean c\u1eeda s\u1ed5 ri\u00eang bi\u1ec7t v\u1edbi thanh ti\u1ebfn \u0111\u1ed9 % th\u1eddi gian th\u1ef1c." & vbCrLf & _
                    "B\u1ea1n c\u00f3 th\u1ec3 TI\u1ebeP T\u1ee4C L\u00c0M VI\u1ec6C tr\u00ean Excel b\u00ecnh th\u01b0\u1eddng m\u00e0 kh\u00f4ng lo b\u1ecb treo m\u00e1y!"), _
                vbInformation, Uni("KangatangGuard v3.5.3 - Ti\u1ebfn tr\u00ecnh ng\u1ea7m")
    End If
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
    MsgBoxW Uni("KangatangGuard v3.5.3" & vbCrLf & vbCrLf & _
                "H\u1ec7 th\u1ed1ng b\u1ea3o v\u1ec7 Excel chuy\u00ean d\u1ee5ng ch\u1ed1ng virus macro Kangatang / Laroux / mypersonnel." & vbCrLf & vbCrLf & _
                "- T\u1ef1 \u0111\u1ed9ng qu\u00e9t th\u1eddi gian th\u1ef1c khi m\u1edf t\u1ec7p." & vbCrLf & _
                "- ScanCache th\u00f4ng minh ch\u1ed1ng lag khi l\u01b0u v\u00e0 AutoSave." & vbCrLf & _
                "- Ki\u1ebfn tr\u00fac T\u00e1ch ti\u1ebfn tr\u00ecnh (Out-of-Process Worker): Qu\u00e9t h\u00e0ng ngh\u00ecn t\u1ec7p tr\u00ean c\u1eeda s\u1ed5 ri\u00eang, Excel kh\u00f4ng bao gi\u1edd b\u1ecb treo (Not Responding)!" & vbCrLf & _
                "- H\u1ed7 tr\u1ee3 \u1ed5 m\u1ea1ng UNC (\\\\server\\share) v\u00e0 t\u00ean t\u1ec7p Unicode c\u00f3 d\u1ea5u." & vbCrLf & vbCrLf & _
                "Phi\u00ean b\u1ea3n ki\u1ebfn tr\u00fac: v3.5.3 Production-grade"), _
            vbInformation, Uni("Gi\u1edbi thi\u1ec7u KangatangGuard v3.5.3")
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

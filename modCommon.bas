Attribute VB_Name = "modCommon"
Global NLoops As Integer, LoopDup As Integer, ListWithFocus As Boolean, Days As Byte
Global sRet As String, Ret As Long, MskErr1 As Boolean, MskErr2 As Boolean
Global DestinDir As String, NoIniArchive As Boolean, bDatedDir As Boolean, bCusDir As Boolean, bUseBoth As Boolean
Global WindowsDir As String, NLoopsTimer As Byte, Interval As Date, IniTime As Date, prevDir As String
Global Default As Boolean, LastBackup As Date, result As Long, Msg As Long, OpenError As Boolean
Global XDir(2) As New Collection, FromPath As String, BaseDir As String, tmpPath As String, newPath As String, bBakNow As Boolean

Public Const Arq = "PCBak.ini"
Public Const SW_SHOW = 5

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
    
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid As NOTIFYICONDATA

Public Type ListaArqs
    Nome As String
    Tamanho As Long
End Type

Public Files() As ListaArqs
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Option Explicit
'Menu item constants.
      Private Const SC_CLOSE       As Long = &HF060&

      'SetMenuItemInfo fMask constants.
      Private Const MIIM_STATE     As Long = &H1&
      Private Const MIIM_ID        As Long = &H2&

      'SetMenuItemInfo fState constants.
      Private Const MFS_GRAYED     As Long = &H3&
      Private Const MFS_CHECKED    As Long = &H8&

      'SendMessage constants.
      Private Const WM_NCACTIVATE  As Long = &H86

      'User-defined Types.
      Private Type MENUITEMINFO
          cbSize        As Long
          fMask         As Long
          fType         As Long
          fState        As Long
          wID           As Long
          hSubMenu      As Long
          hbmpChecked   As Long
          hbmpUnchecked As Long
          dwItemData    As Long
          dwTypeData    As String
          cch           As Long
      End Type

      'Declarations.
      Private Declare Function GetSystemMenu Lib "user32" ( _
          ByVal hwnd As Long, ByVal bRevert As Long) As Long

      Private Declare Function GetMenuItemInfo Lib "user32" Alias _
          "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
          ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

      Private Declare Function SetMenuItemInfo Lib "user32" Alias _
          "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
          ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

      Private Declare Function SendMessage Lib "user32" Alias _
          "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, lParam As Any) As Long

      'Application-specific constants and variables.
      Private Const xSC_CLOSE  As Long = -10
      Private Const SwapID     As Long = 1
      Private Const ResetID    As Long = 2

      Private hMenu  As Long
      Private MII    As MENUITEMINFO

Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long


Function ActivatePrevInstance()
    Dim OldTitle As String
    Dim PrevHndl As Long
    Dim result As Long
    'Save the title of the application.
    OldTitle = App.Title
    'Rename the title of this application so
    '     FindWindow
    'will not find this application instance
    '     .
    App.Title = "unwanted instance"
    'Attempt to get window handle using VB4
    '     class name.
    PrevHndl = FindWindow("ThunderRTMain", OldTitle)
    'Check for no success.


    If PrevHndl = 0 Then
        'Attempt to get window handle using VB5
        '     class name.
        PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
    End If
    'Check if found


    If PrevHndl = 0 Then
        'Attempt to get window handle using VB6
        '     class name
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    End If
    'Check if found


    If PrevHndl = 0 Then
        'No previous instance found.
        Exit Function
    End If
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    'Restore the program.
    result = OpenIcon(PrevHndl)
    'Activate the application.
    result = SetForegroundWindow(PrevHndl)
    'End the application.
    End
End Function

Function SetId(Action As Long) As Long
          Dim MenuID As Long
          Dim Ret As Long

          MenuID = MII.wID
          If MII.fState = (MII.fState Or MFS_GRAYED) Then
              If Action = SwapID Then
                  MII.wID = SC_CLOSE
              Else
                  MII.wID = xSC_CLOSE
              End If
          Else
              If Action = SwapID Then
                  MII.wID = xSC_CLOSE
              Else
                  MII.wID = SC_CLOSE
              End If
          End If

          MII.fMask = MIIM_ID
          Ret = SetMenuItemInfo(hMenu, MenuID, False, MII)
          If Ret = 0 Then
              MII.wID = MenuID
          End If
          SetId = Ret
      End Function



Function Initialize()
On Error GoTo erro

    Dim Lenght As Byte
    
    WindowsDir = String(255, 0)
    Lenght = GetWindowsDirectory(WindowsDir, 254)
    WindowsDir = Left(WindowsDir, Lenght)
    
    If Not Right(WindowsDir, 1) = "\" Then WindowsDir = WindowsDir & "\"
    
    If Dir(WindowsDir & "PCBak.ini") = "" Then
        If Dir(WindowsDir & "PCBak.bak") <> "" Then
            FileCopy WindowsDir & "PCBak.bak", WindowsDir & "PCBak.ini"
        Else
            NoIniArchive = True
        End If
    End If
        
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "AlwaysAt", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        If sRet = "???" Then
            IniTime = vbEmpty
        Else
            frmMain.MaskEdBox1.Text = sRet
            IniTime = TimeSerial(Hour(frmMain.MaskEdBox1.Text), Minute(frmMain.MaskEdBox1.Text), 0)
        End If
    End If

    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "Each", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        If sRet = "???" Then
            Interval = vbEmpty
        Else
            frmMain.MaskEdBox2.Text = sRet
            Interval = TimeSerial(Hour(frmMain.MaskEdBox2.Text), Minute(frmMain.MaskEdBox2.Text), 0)
        End If
    End If
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "Default", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        If sRet = "False" Then
            Default = False
        Else
            Default = True
        End If
    End If
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "Days", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        Dim BsRet As Byte
        BsRet = CByte(sRet)
        If Int(BsRet / 64) = 1 Then frmMain.chkDays(7).Value = True: BsRet = BsRet - 64
        If Int(BsRet / 32) = 1 Then frmMain.chkDays(6).Value = True: BsRet = BsRet - 32
        If Int(BsRet / 16) = 1 Then frmMain.chkDays(5).Value = True: BsRet = BsRet - 16
        If Int(BsRet / 8) = 1 Then frmMain.chkDays(4).Value = True: BsRet = BsRet - 8
        If Int(BsRet / 4) = 1 Then frmMain.chkDays(3).Value = True: BsRet = BsRet - 4
        If Int(BsRet / 2) = 1 Then frmMain.chkDays(2).Value = True: BsRet = BsRet - 2
        If Int(BsRet / 1) = 1 Then frmMain.chkDays(1).Value = True
    End If
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Log", "Save", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "False" Then frmMain.chkLog.Value = False
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Backup", "Incremental", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "True" Then frmMain.chkIncr.Value = True
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destination", "BaseDir", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        On Error GoTo erro1
        frmMain.dirDest.Path = sRet
        frmMain.driveDest.Drive = Left(sRet, 2)
        On Error GoTo erro
    End If
    DestinDir = sRet
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destination", "Custom", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "True" Then frmMain.chkCustom.Value = True
        
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destination", "CustomName", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
   If Not Ret = 0 Then frmMain.txtCusDir = sRet
    DestinDir = DestinDir & sRet
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destination", "And", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "True" Then frmMain.optAnd.Value = True
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destination", "Or", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "True" Then frmMain.optOr.Value = True
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destination", "DatedDirectories", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "True" Then frmMain.chkDated.Value = True
    'DestinDir = DestinDir & sRet
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destination", "DateSeperator", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then frmMain.cmbSep = sRet
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Options", "LoadWin", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "True" Then frmMain.chkServ.Value = True

    
cont:
    'DestinDir = sRet
    'frmMain.txtDest.Text = DestinDir
    NLoops = 0
    ReDim Files(0)
    
    
start:
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Entries", NLoops, "", sRet, 255, Arq)
    If Ret = 0 Then LastBackup = TimeSerial(Hour(Time), Minute(Time), 0): Exit Function
    sRet = Left(sRet, Ret)
    frmMain.lstSource.AddItem sRet
    NLoops = NLoops + 1
    GoTo start

recheckItems
Saída:
    Exit Function
    
erro:
    MsgBox Err.Number & vbLf & vbLf & Err.Description, vbCritical, "Initializing!"
    Resume Next
    
erro1:
    If Err.Number = 68 Or Err.Number = 76 Then
    
    Else
        MsgBox Err.Number & vbLf & Err.Description
    End If
    Resume cont
    
End Function

Function AddItem(OnlyFile As Boolean, Optional WithSubs As Boolean = False)
On Error GoTo erro

    Screen.MousePointer = vbHourglass

    Dim AddPath As String
    
    If Right(frmMain.dirSource.List(frmMain.dirSource.ListIndex), 1) = "\" Then
        AddPath = frmMain.dirSource.List(frmMain.dirSource.ListIndex)
    Else
        AddPath = frmMain.dirSource.List(frmMain.dirSource.ListIndex) & "\"
    End If
    
    If Not OnlyFile Then
        
        If WithSubs Then
            Dim i As Integer, d As String
            GetDirs (AddPath)
            For i = 1 To XDir(0).Count
                If VerificaDup(XDir(0).Item(i) & "\*.*") Then
                    MsgBox "This item is already on the list:" & vbLf & vbLf & XDir(0).Item(i) & "\*.*", vbExclamation
                Else
                    frmMain.lstSource.AddItem XDir(0).Item(i) & "\*.*"
                End If
            Next i
            For i = XDir(0).Count To 1 Step -1
                XDir(0).Remove (i)
            Next i
        End If
        
        If frmMain.lstSource.ListCount = 0 Then
            frmMain.lstSource.AddItem AddPath & "*.*"
            GoTo Saída
        Else
            If VerificaDup(AddPath & "*.*") Then
                MsgBox "This item is already on the list:" & vbLf & vbLf & AddPath & "*.*", vbExclamation
                GoTo Saída
            Else
                frmMain.lstSource.AddItem AddPath & "*.*"
                GoTo Saída
            End If
        End If
        
    Else
    
        Dim Entries As Integer
        For NLoops = 0 To frmMain.fileSource.ListCount - 1
            If frmMain.fileSource.Selected(NLoops) Then
                Entries = Entries + 1
                If Entries > 1 Then GoTo cont
            End If
        Next NLoops

cont:
        If Entries = 1 Then
            If VerificaDup(AddPath & frmMain.fileSource.FileName) Then
                MsgBox "This item is already on the list:" & vbLf & vbLf & AddPath & frmMain.fileSource.FileName, vbExclamation
                GoTo Saída
            Else
                frmMain.lstSource.AddItem AddPath & frmMain.fileSource.FileName
                GoTo Saída
            End If
        ElseIf Entries > 1 Then
            For NLoops = 0 To frmMain.fileSource.ListCount - 1
                If frmMain.fileSource.Selected(NLoops) Then
                    If VerificaDup(AddPath & frmMain.fileSource.List(NLoops)) Then
                        MsgBox "This item is already on the list:" & vbLf & vbLf & AddPath & frmMain.fileSource.List(NLoops), vbExclamation
                    Else
                        frmMain.lstSource.AddItem AddPath & frmMain.fileSource.List(NLoops)
                    End If
                End If
            Next NLoops
        End If
        
    End If
    
Saída:
    Screen.MousePointer = vbDefault
    Exit Function
    
erro:
    MsgBox Err.Number & vbLf & Err.Description, vbCritical
    Resume Saída
                    
End Function

Function recheckItems()
Dim i As Integer, d As String, n As Integer, strPath As String
For n = 0 To frmMain.lstSource.ListCount - 1
    strPath = frmMain.lstSource.List(n)
    If Right(strPath, 4) = "\*.*" Then
        strPath = Left(strPath, Len(strPath) - 3)
    End If
    GetDirs (strPath)
    For i = 1 To XDir(0).Count
        If Not VerificaDup(XDir(0).Item(i) & "\*.*") Then
            frmMain.lstSource.AddItem XDir(0).Item(i) & "\*.*"
        End If
    Next i
    For i = XDir(0).Count To 1 Step -1
        XDir(0).Remove (i)
    Next i
    Dim Entries As Integer
    For NLoops = 0 To frmMain.fileSource.ListCount - 1
        If frmMain.fileSource.Selected(NLoops) Then
            Entries = Entries + 1
            If Entries > 1 Then GoTo cont
        End If
    Next NLoops

cont:
    If Entries = 1 Then
        If Not VerificaDup(strPath & frmMain.fileSource.FileName) Then
            frmMain.lstSource.AddItem strPath & frmMain.fileSource.FileName
        End If
    ElseIf Entries > 1 Then
        For NLoops = 0 To frmMain.fileSource.ListCount - 1
            If frmMain.fileSource.Selected(NLoops) Then
                If Not VerificaDup(strPath & frmMain.fileSource.List(NLoops)) Then
                    frmMain.lstSource.AddItem strPath & frmMain.fileSource.List(NLoops)
                End If
            End If
        Next NLoops
    End If
 
Next n
End Function

Function GetDirs(Path As String)
    'on error Resume Next
    Dim vDirName As String, LastDir As String
    Dim i As Integer
    
    'Adjust so No Deletion of Drive
    If Len(Path$) < 3 Then Exit Function

    If Right(Path$, 1) <> "\" Then
        XDir(0).Add Path$
        Path$ = Path$ & "\"
    End If

    vDirName = Dir(Path, vbDirectory) ' Retrieve the first entry.

    Do While vDirName <> ""
        If vDirName <> "." And vDirName <> ".." Then
            If (GetAttr(Path & vDirName)) = vbDirectory Then
                LastDir = vDirName
                'Finds Directory Name then Repeats
                GetDirs (Path$ & vDirName)
                vDirName = Dir(Path$, vbDirectory)

                Do Until vDirName = LastDir Or vDirName = ""
                    vDirName = Dir
                Loop

                If vDirName = "" Then Exit Do
            End If
        End If
    
    vDirName = Dir
    
    Loop

End Function

Function ExtractText(FullText As String, token As String, Optional StartAtLeft = True, Optional IncludeLeftSide = True) As String
'ExtractText(Path$, ":", False, False)
    
    Dim i As Integer
    If StartAtLeft = True And IncludeLeftSide = True Then
        ExtractText = FullText
        For i = 1 To Len(FullText)
            If Mid(FullText, i, 1) = token Then
                ExtractText = Left(FullText, i - 1)
                Exit Function
            End If
        Next

    ElseIf StartAtLeft = True And IncludeLeftSide = False Then
        ExtractText = FullText
        For i = 1 To Len(FullText)
            If Mid(FullText, i, 1) = token Then
                ExtractText = Right(FullText, Len(FullText) - i)
                Exit Function
            End If
        Next
    
    ElseIf StartAtLeft = False And IncludeLeftSide = True Then
        ExtractText = ""
        For i = Len(FullText) To 1 Step -1
            If Mid(FullText, i, 1) = token Then
                ExtractText = Left(FullText, i - 1)
                Exit Function
            End If
        Next

    ElseIf StartAtLeft = False And IncludeLeftSide = False Then
        ExtractText = ""
        For i = Len(FullText) To 1 Step -1
            If Mid(FullText, i, 1) = token Then
                ExtractText = Right(FullText, Len(FullText) - i)
                Exit Function
            End If
        Next
    End If

End Function


Function MtxAdicionaArq(CamCompleto As String)
    
    If UBound(Files) = 1 Then
        Files(1).Nome = CamCompleto
        Files(1).Tamanho = FileLen(CamCompleto)
        ReDim Preserve Files(2)
    Else
        Files(UBound(Files)).Nome = CamCompleto
        Files(UBound(Files)).Tamanho = FileLen(CamCompleto)
        ReDim Preserve Files(UBound(Files) + 1)
    End If

End Function

Function MtxAdicionaDir(ByVal Caminho As String)
On Error GoTo erro

    Dim b As String, n As Integer, ShortPath As String
    
    If Not Right(Caminho, 1) = "*" Then Caminho = Caminho & "*.*"

    ShortPath = Left(Caminho, Len(Caminho) - 3)

    If Not UBound(Files) = 1 Then
        n = UBound(Files) + 1
        ReDim Preserve Files(n)
    End If
    
    b = Dir(Caminho)
    If b = "" Then
        Exit Function
    Else
        Files(UBound(Files) - 1).Nome = ShortPath & b
        Files(UBound(Files) - 1).Tamanho = FileLen(ShortPath & b)
    End If

    Do
    b = Dir
    If b = "" Then Exit Do
        
    With Files(n)
        .Nome = ShortPath & b
        .Tamanho = FileLen(ShortPath & b)
    End With
    n = n + 1
    ReDim Preserve Files(n)
    Loop

Saída:
    Exit Function
    
erro:
    MsgBox "MtxAddDir:" & vbLf & vbLf & Err.Number & ":" & Err.Description, vbCritical
    Resume Saída

End Function

Function Backup()
On Error GoTo erro

    Screen.MousePointer = vbHourglass
    
    Dim DateBak As Date, TimeBak As Date, ErrString As String
    Dim NDirs As Integer, File As String, TskID As Double, TotFiles As Long, TotalFilesCopied As Long
    Dim ErroDest As Byte, ArqAtr As Byte, Tam As Long, dirFolder As String, r As Boolean, srcTmp As String, destFile As String, destpath As String
    Dim FileCnt As Long, rdOnly As Boolean
    recheckItems
    frmMain.SSTab1.Tab = 6
    
    TimeBak = Now
    DateBak = Date
    
    frmMain.Caption = "Creating file list..."
    DestinDir = frmMain.txtDest
    dirFolder = Dir$(DestinDir, vbDirectory)
    If dirFolder = "" Then
        MkDir (DestinDir)
    End If
    If Not Right(DestinDir, 1) = "\" Then DestinDir = DestinDir & "\"

    For NLoops = 0 To frmMain.lstSource.ListCount - 1
        If Right(frmMain.lstSource.List(NLoops), 1) = "*" Then
            MtxAdicionaDir (Left(frmMain.lstSource.List(NLoops), Len(frmMain.lstSource.List(NLoops)) - 3))
        Else
            MtxAdicionaArq (frmMain.lstSource.List(NLoops))
        End If
    Next NLoops

    frmMain.Caption = "Doing the backup..."
    If frmMain.chkLog Then
        Open WindowsDir & "PCBak.log" For Output As #1
        Print #1, "Initializing backup at " & Now
        Print #1,
    End If
    
    frmMain.Label10.Caption = "Copying from"
    frmMain.Label12.Caption = "to"
    
    FileCnt = UBound(Files)
    TotFiles = UBound(Files) - 1
    For NLoops = 0 To TotFiles
        DoEvents
        If Not Files(NLoops).Nome = "" Then
            ArqAtr = GetAttr(Files(NLoops).Nome)


cont:
            srcTmp = Files(NLoops).Nome
            destFile = ReturnFileName(srcTmp)
            'destFile = ReturnFileName(srcTmp)
            If destFile = "" Then GoTo Saída
            If frmMain.chkIncr Then
                If ArqAtr = vbReadOnly Then rdOnly = True
                If ArqAtr And vbArchive <> 0 Then
                    destpath = GetParentDir(Left(srcTmp, (Len(srcTmp) - Len(destFile))))
                    FileCopy srcTmp, destpath & destFile
                    If frmMain.chkLog Then Print #1, srcTmp & " --> " & destpath & destFile & ", status: ";
                    If Not rdOnly Then
                        SetAttr srcTmp, (ArqAtr - vbArchive)
                    End If
                    If frmMain.chkLog Then Print #1, "Ok!"
                    Tam = Tam + FileLen(srcTmp)
                    TotalFilesCopied = TotalFilesCopied + 1
                End If
            Else
                destpath = GetParentDir(Left(srcTmp, (Len(srcTmp) - Len(destFile))))
                FileCopy srcTmp, destpath & destFile
                If frmMain.chkLog Then Print #1, srcTmp & " --> " & destpath & destFile & ", status: ";
                If frmMain.chkLog Then Print #1, "Ok!"
                Tam = Tam + FileLen(srcTmp)
                TotalFilesCopied = TotalFilesCopied + 1
            End If
            frmMain.Label11.Caption = srcTmp
            frmMain.Label13.Caption = destpath & destFile
            frmMain.Label14.Caption = "File " & NLoops & " of " & FileCnt
            frmMain.Label14.Caption = "File " & NLoops & " of " & FileCnt & ", total: " & _
                        Format(Tam / 1024 / 1024, "standard") & " Mb"
        End If
    Next NLoops

Saída:
    If srcTmp <> "" Then
        If frmMain.chkLog Then
            Print #1,
            Print #1, "Copied " & TotalFilesCopied & " files, " & Format(Tam / 1024 / 1024, "standard") & " Mb, From " & _
                Format(TimeBak, "short time") & " to " & Format(Time, "short time") & " on " & Format(DateBak, "short date") & "."
            Close #1
        End If
    End If
    frmMain.Label10.Caption = ""
    frmMain.Label11.Caption = ""
    frmMain.Label12.Caption = ""
    frmMain.Label13.Caption = ""
    frmMain.Label14.Caption = "Copied " & TotalFilesCopied & " files, " & Format(Tam / 1024 / 1024, "standard") & " Mb, From " & _
                Format(TimeBak, "short time") & " to " & Format(Time, "short time") & " on " & Format(DateBak, "short date") & "."
    ReDim Files(0)
    frmMain.Caption = "PC Backup ver 1.0"
    Screen.MousePointer = vbDefault
    Exit Function

erro:
    ErrString = vbLf & vbLf & "While trying to copy:" & vbLf & srcTmp & _
        vbLf & "to" & vbLf & destpath & destFile & vbLf & _
        vbLf & "Try again?"
    
    If frmMain.chkLog Then Print #1, "ERROR: " & Err.Number & " - " & Err.Description;
    
    Select Case Err.Number
        
        Case 5      'Invalid procedure call ???
            Resume Next
                    
        Case 52    'Bad filename
            MsgBox "Bad filename! (erro 52)" & vbLf & vbLf & srcTmp, vbExclamation
            Resume Next
            
        Case 53     'File not found
            MsgBox "File not found! (erro 53)" & vbLf & vbLf & srcTmp, vbExclamation
            Resume Next
                    
        Case 57     'Device I/O error
            If MsgBox("Destiny disk not ready! (erro 57)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
            
        Case 61     'Disk full
            If MsgBox("Destiny disk full! (error 61)" & ErrString, vbExclamation + vbYesNo) = vbYes Then Resume cont
                    
        Case 70    'Permission denied
            If MsgBox("Destiny directory or drive protected! (error 70)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
            
        Case 71    'Disk not ready
            If MsgBox("Destiny disk not ready! (error 71)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
                
        Case 75     'Path/file access error
            If destFile <> "" Then
                SetAttr destpath & destFile, (GetAttr(destpath & destFile) - vbReadOnly)
            End If
            Resume cont
        
        Case 76     'Path not found
            If MsgBox("Destiny directory unavailable! (error 76)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
        
        Case Else
            If MsgBox("PANIC!!" & vbLf & vbLf & Err.Number & ": " & Err.Description & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
    
    End Select
    
    Resume Saída
End Function

Function ReturnFileName(ByVal Arq As String) As String
'Arq is the full path, returns only the filename
    
    Dim n As Integer, X As Integer, i As Integer
    i = InStr(Arq, "\")
    For n = Len(Arq) To 1 Step -1
        If Mid(Arq, n, 1) = "\" Then
            
                ReturnFileName = Right(Arq, Len(Arq) - n)
                Exit Function
            
            
        End If
    Next n

End Function
Function CheckTime()
On Error GoTo erro
Dim mainCap
mainCap = frmMain.Caption
    If frmMain.optAlways And Not IniTime = vbEmpty Then
        If IniTime = TimeSerial(Hour(Time), Minute(Time), 0) Then
            recheckItems
            frmMain.Caption = "Doing the Backup..."
            frmMain.Refresh
            Backup
            LastBackup = TimeSerial(Hour(Time), Minute(Time), 0)
            frmMain.Caption = mainCap
            frmMain.Refresh
        End If
    End If
    
    If frmMain.optEach And Not Interval = vbEmpty Then
        If TimeSerial(Hour(Time), Minute(Time), 0) = TimeValue(Interval + LastBackup) Then
            recheckItems
            frmMain.Caption = "Doing the Backup..."
            frmMain.Refresh
            Backup
            LastBackup = TimeSerial(Hour(Time), Minute(Time), 0)
            frmMain.Caption = mainCap
            frmMain.Refresh
        End If
    End If
    
Saída:
    Exit Function
    
erro:
    If Not Err.Number = 13 Then MsgBox Err.Number & vbLf & Err.Description
    Resume Saída
    
End Function


Function SaveChanges()
On Error GoTo erro
Dim c
    Screen.MousePointer = vbHourglass
        
    On Error Resume Next
    c = Dir$(WindowsDir & "PCBak.bak", vbNormal)
    If c = "" Then
        Name WindowsDir & Arq As WindowsDir & "PCBak.bak"
        Kill WindowsDir & Arq
    Else
        Kill WindowsDir & "PCBak.bak"
        Name WindowsDir & Arq As WindowsDir & "PCBak.bak"
        'Kill WindowsDir & Arq
    End If
    
    On Error GoTo erro
    If Not frmMain.MaskEdBox1.Text = "__:__" Then
        Call WritePrivateProfileString("When", "AlwaysAt", frmMain.MaskEdBox1.Text, Arq)
        IniTime = TimeSerial(Hour(frmMain.MaskEdBox1.Text), Minute(frmMain.MaskEdBox1.Text), 0)
    Else
        Call WritePrivateProfileString("When", "AlwaysAt", "???", Arq)
        IniTime = vbEmpty
    End If
    
    If Not frmMain.MaskEdBox2.Text = "__:__" Then
        Call WritePrivateProfileString("When", "Each", frmMain.MaskEdBox2.Text, Arq)
        Interval = TimeSerial(Hour(frmMain.MaskEdBox2.Text), Minute(frmMain.MaskEdBox2.Text), 0)
    Else
        Call WritePrivateProfileString("When", "Each", "???", Arq)
        Interval = vbEmpty
    End If
    
    If frmMain.optAlways Then
        Call WritePrivateProfileString("When", "Default", False, Arq)
    Else
        Call WritePrivateProfileString("When", "Default", True, Arq)
    End If
    
    If frmMain.optDaily Then
        Call WritePrivateProfileString("When", "Days", "0", Arq)
    Else
        Days = 0
        Dim n As Byte
        For n = 0 To 6
            If frmMain.chkDays(n + 1) Then Days = Days + 2 ^ n
        Next n
        Call WritePrivateProfileString("When", "Days", Days, Arq)
    End If
            
    If frmMain.chkLog Then
        Call WritePrivateProfileString("Log", "Save", "True", Arq)
    Else
        Call WritePrivateProfileString("Log", "Save", "False", Arq)
    End If
            
    If frmMain.chkIncr Then
        Call WritePrivateProfileString("Backup", "Incremental", "True", Arq)
    Else
        Call WritePrivateProfileString("Backup", "Incremental", "False", Arq)
    End If
    If frmMain.chkCustom Then
        Call WritePrivateProfileString("Destination", "Custom", "True", Arq)
    Else
        Call WritePrivateProfileString("Destination", "Custom", "False", Arq)
    End If
    If frmMain.txtCusDir <> "" Then
        Call WritePrivateProfileString("Destination", "CustomName", frmMain.txtCusDir, Arq)
    End If
    If frmMain.optAnd Then
        Call WritePrivateProfileString("Destination", "And", "True", Arq)
    Else
        Call WritePrivateProfileString("Destination", "And", "False", Arq)
    End If
    If frmMain.optOr Then
        Call WritePrivateProfileString("Destination", "Or", "True", Arq)
    Else
        Call WritePrivateProfileString("Destination", "Or", "False", Arq)
    End If
    If frmMain.cmbSep <> "" Then
        Call WritePrivateProfileString("Destination", "DateSeperator", frmMain.cmbSep, Arq)
    End If
    If frmMain.chkDated Then
        Call WritePrivateProfileString("Destination", "DatedDirectories", "True", Arq)
    Else
        Call WritePrivateProfileString("Destination", "DatedDirectories", "False", Arq)
    End If

    Call WritePrivateProfileString("Destination", "BaseDir", BaseDir, Arq)
    
    If frmMain.chkServ Then
        Call WritePrivateProfileString("Options", "LoadWin", True, Arq)
    Else
        Call WritePrivateProfileString("Options", "LoadWin", False, Arq)
    End If
    
    For NLoops = 0 To frmMain.lstSource.ListCount - 1
        If WritePrivateProfileString("Entries", CStr(NLoops), frmMain.lstSource.List(NLoops), Arq) = 0 Then
            MsgBox "INI file full." & vbLf & "Last saved entry: " & frmMain.lstSource.List(NLoops - 1), vbCritical
            GoTo Saída
        End If
    Next NLoops

    Screen.MousePointer = vbDefault
    
Saída:
    Exit Function
    
erro:
    MsgBox Err.Number & vbLf & Err.Description, vbCritical
    Resume Saída

End Function

Function VerificaDup(Item As String) As Boolean

    For LoopDup = 0 To frmMain.lstSource.ListCount - 1
        If frmMain.lstSource.List(LoopDup) = Item Then
            VerificaDup = True
            Exit Function
        End If
    Next LoopDup
    
    VerificaDup = False

End Function


Function VerifyErrors() As Boolean

    If frmMain.lstSource.ListCount = 0 Then
        MsgBox "You must specify at least one file or directory for the backup!", vbCritical
        frmMain.SSTab1.Tab = 0
        GoTo erro
    End If
    
    If Len(frmMain.txtDest.Text) = 0 Then
        MsgBox "You must specify the destination directory.", vbCritical
        frmMain.SSTab1.Tab = 1
        frmMain.txtDest.SetFocus
        GoTo erro
    ElseIf frmMain.txtDest.Text = "c:\" Or frmMain.txtDest.Text = "C:\" Then
        If MsgBox("The destination dir was left as C:\." & vbLf & vbLf & "Confirm?", _
            vbYesNo + vbExclamation) = vbNo Then
            frmMain.SSTab1.Tab = 1
            frmMain.txtDest.SetFocus
            GoTo erro
        End If
    ElseIf frmMain.optAlways And frmMain.MaskEdBox1.Text = "__:__" Then
        MsgBox "You must specify a time for the backup!", vbCritical
        frmMain.SSTab1.Tab = 2
        frmMain.MaskEdBox1.SetFocus
        GoTo erro
    ElseIf frmMain.optEach And frmMain.MaskEdBox2.Text = "__:__" Then
        MsgBox "You must specify an interval for the backup!", vbCritical
        frmMain.SSTab1.Tab = 2
        frmMain.MaskEdBox2.SetFocus
        GoTo erro
    End If
    
    VerifyErrors = False

Saída:
    Exit Function
    
erro:
    VerifyErrors = True
    
End Function

Function GetParentDir(ByVal Caminho As String)
'On Error GoTo erro

    Dim b As String, n As Integer, ShortPath As String, bDirCreate As Boolean, i As Integer
    Dim pathOnly As String, tmpPath As String, newPath As String, dirFolder As String
    'ShortPath = Right(Caminho, Len(Caminho) - 3)
    ShortPath = Caminho
        For i = 0 To Len(ShortPath)
        pathOnly = Left(ShortPath, i)
        If Right(pathOnly, 1) = "\" Then
            tmpPath = Right(ShortPath, Len(ShortPath) - i)
            newPath = DestinDir & tmpPath
            dirFolder = Dir$(newPath, vbDirectory)
            If dirFolder = "" Then
                'MkDir (newPath)
                bDirCreate = gblnCreateDir(newPath)
                If bDirCreate Then
                    GetParentDir = newPath
                End If
                Exit Function
            Else
                GetParentDir = newPath
                Exit Function
          
            End If
        End If
    Next
GetParentDir = newPath
Exit Function

  
End Function

Function killClose()
Dim Ret As Long

hMenu = GetSystemMenu(frmMain.hwnd, 0)
MII.cbSize = Len(MII)
MII.dwTypeData = String(80, 0)
MII.cch = Len(MII.dwTypeData)
MII.fMask = MIIM_STATE
MII.wID = SC_CLOSE
Ret = GetMenuItemInfo(hMenu, MII.wID, False, MII)
Ret = SetId(SwapID)
If Ret <> 0 Then

    If MII.fState = (MII.fState Or MFS_GRAYED) Then
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If

    MII.fMask = MIIM_STATE
    Ret = SetMenuItemInfo(hMenu, MII.wID, False, MII)
    If Ret = 0 Then
        Ret = SetId(ResetID)
    End If

    Ret = SendMessage(frmMain.hwnd, WM_NCACTIVATE, True, 0)
End If
End Function

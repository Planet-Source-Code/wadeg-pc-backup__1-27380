Attribute VB_Name = "modMakeDirs"
Option Explicit

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_READONLY = &H1

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Function gblnCreateDir(ByVal PrmStrPath As String) As Boolean
'***********************************************************************
' Purpose : To create a directory with parent directories
' Inputs  : PrmStrPath - Directory to be create
'           PrmDatabase - If the session is in the begin trans mode you
'                         should send the Database. Because if any error
'                         occurs during the execution of this function,
'                         this function will rollback the entire
'                         transaction and return FALSE
' Returns : True if Successful otherwise return FALSE
'***********************************************************************

Dim strError As String
Dim blnError As Boolean
Dim strCreatedPath As String
Dim intCounter As Integer
Dim strCurPath As String

    gblnCreateDir = False

    On Error GoTo Err_gblnCreateDir

    PrmStrPath = Trim(PrmStrPath)

    If gblnIsDirExist(PrmStrPath) Or (Len(PrmStrPath) = 3 And Mid(PrmStrPath, 2) = ":\") Then
        gblnCreateDir = True
        GoTo Exit_gblnCreateDir
    End If

    While Right(PrmStrPath, 1) = "\"
        PrmStrPath = Left(PrmStrPath, Len(PrmStrPath) - 1)
    Wend
    
    strCreatedPath = ""
    
    intCounter = 0
    While intCounter < Len(PrmStrPath)
    
        strCurPath = ""
        While True And intCounter < Len(PrmStrPath)
            intCounter = intCounter + 1
            If Mid(PrmStrPath, intCounter, 1) <> "\" Then
                strCurPath = strCurPath & Mid(PrmStrPath, intCounter, 1)
            Else
                strCurPath = strCurPath & "\"
                If strCurPath <> "\\" And strCurPath <> "\" Then
                    GoTo EndLoop
                End If
            End If
        Wend
EndLoop:
        strCreatedPath = strCreatedPath & strCurPath

        strCreatedPath = Trim(strCreatedPath)
        While Right(strCreatedPath, 1) = "\"
            strCreatedPath = Left(strCreatedPath, Len(strCreatedPath) - 1)
        Wend
    
        On Error GoTo NetPath
        If Left(strCreatedPath, 2) = "\\" And (glngOccurs(strCreatedPath, "\") = 3 Or glngOccurs(strCreatedPath, "\") = 2) Or (Len(strCreatedPath) = 2 And Right(strCreatedPath, 1) = ":") Then GoTo NetPath
        If Not gblnIsDirExist(strCreatedPath) Then
            If Not gblnMKDir(strCreatedPath) Then GoTo Err_gblnCreateDir
        End If
        
NetPath:
        If blnError Then GoTo Exit_gblnCreateDir
        On Error GoTo Err_gblnCreateDir
        strCreatedPath = strCreatedPath & "\"
    
    Wend
        
    On Error GoTo Err_gblnCreateDir
    If Not gblnIsDirExist(strCreatedPath) Then GoTo Err_gblnCreateDir
    
    gblnCreateDir = True
    strError = ""
    strCreatedPath = ""
    strCurPath = ""

Exit_gblnCreateDir:
    Exit Function

Err_gblnCreateDir:
    
    blnError = True
    strError = Err.Description
    MsgBox "Cannot create directory '" & PrmStrPath & "'" & Chr(13) & strError, vbExclamation
    Resume Exit_gblnCreateDir
    
End Function

Function gblnIsDirExist(ByVal PrmStrPath As String) As Boolean
'***********************************************************************
' Purpose : To check the existence of passed Diretory
' Inputs  : PrmStrPath - Directory Name
' Returns : Return TRUE if Exists otherwise return FALSE
'***********************************************************************

    Dim ss As WIN32_FIND_DATA

    gblnIsDirExist = False

    PrmStrPath = Trim(PrmStrPath)
    If Right(PrmStrPath, 1) = "\" Then PrmStrPath = Left(PrmStrPath, Len(PrmStrPath) - 1)
    
    If FindFirstFile(PrmStrPath, ss) = INVALID_HANDLE_VALUE Then GoTo Next_Option
    gblnIsDirExist = (ss.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 And gstrVBString(ss.cFileName) <> ""
    If gblnIsDirExist Then Exit Function

Next_Option:
    If FindFirstFile(PrmStrPath & "\*.*", ss) = INVALID_HANDLE_VALUE Then Exit Function
    gblnIsDirExist = (ss.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 And gstrVBString(ss.cFileName) <> ""

End Function

Public Function gblnMKDir(ByVal PrmStrDirName As String) As Boolean
'***********************************************************************
' Purpose : To create a directory
' Inputs  : PrmStrDirName - Directory Name
' Returns : Return TRUE if successfull, FALSE if fail
'***********************************************************************

    Dim SecAttr As SECURITY_ATTRIBUTES

    PrmStrDirName = Trim(PrmStrDirName)
    If gblnIsDirExist(PrmStrDirName) Then
        gblnMKDir = True
        Exit Function
    End If

    If Right(PrmStrDirName, 1) = "\" Then PrmStrDirName = Left(PrmStrDirName, Len(PrmStrDirName) - 1)
    gblnMKDir = CreateDirectory(PrmStrDirName, SecAttr) > 0

End Function

Public Function glngOccurs(ByVal PrmString As String, ByVal PrmSearchstring As String) As Long
'***********************************************************************
' Purpose : To find the no. of occurrences of a string with in a string
' Inputs  : PrmString - Source string
'           PrmSearchstring - String to be find the occurrences
' Returns : If the string is not found then 0, otherwise No. of occurences
'***********************************************************************

    glngOccurs = 0

    While True
        If PrmSearchstring = "" Then Exit Function
        If InStr(PrmString, PrmSearchstring) = 0 Then Exit Function
        PrmString = Mid(PrmString, InStr(PrmString, PrmSearchstring) + Len(PrmSearchstring), Len(PrmString))
        glngOccurs = glngOccurs + 1
    Wend

End Function

Public Function gstrVBString(ByVal PrmStrNullTerminatedString As String) As String
'***********************************************************************
' Purpose : To convert null terminated string to VB String
' Inputs  : PrmStrNullTerminatedString - Null terminated string
' Returns : Returns VB string
'***********************************************************************

    gstrVBString = Trim(PrmStrNullTerminatedString)
    If InStr(gstrVBString, Chr(0)) = 0 Then Exit Function
    gstrVBString = Left(gstrVBString, InStr(gstrVBString, Chr(0)) - 1)

End Function




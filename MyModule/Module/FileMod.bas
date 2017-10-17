Attribute VB_Name = "FileMod"
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_ALL = &H10000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Private Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function llseek Lib "kernel32.dll" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long

Function GetFile(ByVal Filename As String)
    Dim b() As Byte
    a = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If a = -1 Then GoTo err
    n = GetFileSize(a, 0)
    ReDim b(n)
    ReadFile a, b(0), n, n, 0&
    lclose a
    GetFile = b
    Exit Function
err:
End Function

Sub PutFile(ByVal Filename As String, tBytes() As Byte, Optional ByVal Mode As Byte = 2)
    a = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, Mode, 0&, 0&)
    n = UBound(tBytes)
    WriteFile a, tBytes(0), n, n, 0&
    lclose a
End Sub

Sub CopyFileAPI(ByVal tSource As String, ByVal tTarget As String, Optional ByVal tM As Long = 0, Optional ByVal tStartFrom As Long = 1)
    Dim b() As Byte
    Dim n As Long
    Dim tmM As Long
    a = CreateFile(tSource, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    'If tStartFrom = 1 Then
        c = CreateFile(tTarget, GENERIC_ALL, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, 2, 0&, 0&)
    'Else
    '    c = CreateFile(tTarget, GENERIC_ALL, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    '    n = GetFileSize(c, 0)
    '    llseek c, n, 0
    'End If
    n = tStartFrom * tM
    llseek a, n, 0
    'n = GetFileSize(a, 0)
    'If n = 0 Then GoTo fin
    tmM = n
    'If tM = 0 Or tM > tmM Then tM = tmM
    ReDim b(tM - 1)
    'For i = 2 To tStartFrom
    '    ReadFile a, b(0), tM, tM, 0&
    '    tmM = tmM - tM
    'Next
    While tM <> 0
        ReadFile a, b(0), tM, tM, 0&
        WriteFile c, b(0), tM, tM, 0&
        'tmM = tmM - tM
    Wend
    'ReadFile a, b(0), tM, tM, 0&
    'WriteFile c, b(0), tM, tM, 0&
fin:
    lclose a
    lclose c
End Sub

Function GetFile2(ByVal Filename As String) As String
    Dim b() As Byte
    a = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If a = -1 Then GoTo err
    n = GetFileSize(a, 0)
    ReDim b(n)
    hread a, b(0), n
    lclose a
    GetFile2 = StrConv(b(), vbUnicode)
    Exit Function
err:
End Function

Sub PutFile2(ByVal Filename As String, tString As String, Optional ByVal Mode As Byte = 2)
    a = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, Mode, 0&, 0&)
    hwrite a, tString, Len(tString)
    lclose a
End Sub


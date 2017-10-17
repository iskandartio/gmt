Attribute VB_Name = "Start"
Global CN As New ADODB.Connection
Global RS As New ADODB.Recordset
Global ConnectionString As String
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global pProvider As String
Global pDataSource As String
Global pDBName As String
Global pDBPassword As String
Global pReportDir As String
Global pServerDir As String
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Function CompactDB(ByVal pServerDir As String, ByVal pDBName As String, ByVal pDBPassword As String) As Boolean
On Error GoTo err
Dim db As New DAO.DBEngine
    
    If Dir(pServerDir & "gmt.ldb") <> "" Then
        Kill pServerDir & "gmt.ldb"
    End If
    db.CompactDatabase pServerDir & pDBName, pServerDir & "a.b", ";pwd=" & pDBPassword, , ";pwd=" & pDBPassword
    If FileSystem.FileLen(pServerDir & "a.b") < 100000000 Then
        MsgBox "Ukuran a terlalu kecil"
        GoTo err
    End If
    Kill pServerDir & pDBName
    FileSystem.FileCopy pServerDir & "a.b", pServerDir & pDBName
    
    CompactDB = True
    Exit Function
err:
    MsgBox err.Description
    MsgBox "Tutup Dulu Aplikasi Akses"
    CompactDB = False
End Function

Sub Main()
'On Error Resume Next
Dim AllString As String
    If App.PrevInstance Then ExitProcess 0
    MousePointer = vbHourglass
    AllString = Decrypt(App.Path & "\" & "setting.txt", "Iskandar Tio")
    'Encrypt App.Path & "\LOCAL.txt", App.Path & "\" & "setting.txt", "Iskandar Tio", True
    Dim a() As String
    AllString = Replace(AllString, "app.path", App.Path)
    a = Split(AllString, Chr(13) & Chr(10))
    For i = 0 To UBound(a)
        If InStr(a(i), "ConnectionString") = 1 Then
            ConnectionString = Mid(a(i), 18)
        ElseIf InStr(a(i), "Provider") = 1 Then
            pProvider = Mid(a(i), 10)
        ElseIf InStr(a(i), "DataSource") = 1 Then
            pDataSource = Mid(a(i), 12)
        ElseIf InStr(a(i), "DatabaseName") = 1 Then
            pDBName = Mid(a(i), 14)
        ElseIf InStr(a(i), "DatabasePassword") = 1 Then
            pDBPassword = Mid(a(i), 18)
        ElseIf InStr(a(i), "ServerDir") = 1 Then
            pServerDir = Mid(a(i), 11)
        ElseIf InStr(a(i), "ReportDir") = 1 Then
            pReportDir = Mid(a(i), 11)
        End If
    Next
    'Dim p As New cPing
    'If p.GetIP("server") = "" Then
    '    MsgBox "Network server belum aktif"
    '    ExitProcess 0
    'End If
    
    ConnectionString = "Provider=" & pProvider & ";Data Source=" & pServerDir & "\" & pDBName & ";Persist Security Info=False;Jet OLEDB:Database Password=" & pDBPassword
    If Not CompactDB(pServerDir, pDBName, pDBPassword) Then Exit Sub
    FormServer.Show
    MousePointer = vbDefault

End Sub

Sub TestDir(ByVal tStr As String)
On Error GoTo err
    MkDir tStr
err:
End Sub

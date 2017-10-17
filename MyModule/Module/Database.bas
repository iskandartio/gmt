Attribute VB_Name = "Database"
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Sub Connect(ByVal a As String)
On Error GoTo err
    If CN.State = 1 Then CN.Close
    CN.Open a
    RS.ActiveConnection = CN
    RS.LockType = 1 'adLockReadOnly
    RS.CursorType = 1 'adOpenKeyset
    RS.CursorLocation = 3 'adUseClient
    
    Exit Sub
err:
    Connect2 Replace(a, "server", "192.168.0.1")
End Sub

Sub Connect2(ByVal a As String)
    If CN.State = 1 Then CN.Close
    CN.Open a
    pServerDir = "\\192.168.0.1\s3rv3rABC$"
    RS.ActiveConnection = CN
    RS.LockType = 1 'adLockReadOnly
    RS.CursorType = 1 'adOpenKeyset
    RS.CursorLocation = 3 'adUseClient
End Sub

Sub query(ByVal QueryString As String)
On Error GoTo err
    a = Replace(QueryString, "~", pTipe)
    If RS.State = 1 Then RS.Close
    RS.Open Trim(a)
    Exit Sub
err:
    MsgBox err.Description
End Sub

Function ExecMe(ByVal tSQLString As String, Optional ByVal tAutoChange As Boolean = True) As Long
'On Error Resume Next
    ExecMe = -1
    If tAutoChange Then
        a = Replace(tSQLString, "~", pTipe)
    Else
        a = tSQLString
    End If
    If tSQLString <> "" Then CN.Execute a, ExecMe
    If pOffLineMode And ExecMe > 0 Then
        If pTransState = 0 Then
            Open App.Path & "\SQLLog.txt" For Append As #1
            Print #1, a
            Close
        Else
            pSQLLog = pSQLLog & a & vbCrLf
        End If
    End If
End Function

Sub BeginTransaction()
    If pOffLineMode Then
        pTransState = 1
        pSQLLog = ""
    End If
    CN.BeginTrans
End Sub

Sub CommitTransaction()
    If pOffLineMode Then
        pTransState = 0
        If pSQLLog <> "" Then
            Open App.Path & "\SQLLog.txt" For Append As #1
            Print #1, Left(pSQLLog, Len(pSQLLog) - 2)
            Close
            pSQLLog = ""
        End If
    End If
    CN.CommitTrans
End Sub

Sub RollBackTransaction()
    If pOffLineMode Then
        pTransState = 0
        pSQLLog = ""
    End If
    CN.RollbackTrans
End Sub

Sub MappingNetWork(ByVal tNetWorkName As String, ByVal tNetWorkUser As String, ByVal tNetWorkPassword As String)
Dim a As NETRESOURCE
    a.lpRemoteName = tNetWorkName
    WNetAddConnection2 a, tNetWorkPassword, tNetWorkUser, 0
End Sub

Sub UnmapNetwork(ByVal tNetWorkName As String)
    WNetCancelConnection2 tNetWorkName, 0, -1
End Sub

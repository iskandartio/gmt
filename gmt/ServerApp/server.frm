VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormServer 
   Caption         =   "SERVER"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCompactDB 
      Caption         =   "Compact DB"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton fQuit 
      Caption         =   "&QUIT"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton FSET 
      Caption         =   "Set Offline Data"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3840
      Top             =   5880
   End
   Begin VB.TextBox lquery 
      Height          =   7455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   9855
   End
   Begin VB.ListBox lname 
      Height          =   7470
      ItemData        =   "server.frx":0000
      Left            =   9840
      List            =   "server.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Client 
      Index           =   0
      Left            =   960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FormServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1 As String
Dim kontinu As Byte
Dim ConnectTime As String
Dim NumClient As Byte
Dim backup As String
Dim q As String
Dim q1 As String
Dim p As Long
Dim p1 As Long
Dim lp As Long
Dim lq As String
Dim y As New XArrayDB
Dim mUnload As Boolean
Private WithEvents m_cZ As cZip
Attribute m_cZ.VB_VarHelpID = -1
Private WithEvents m_cUZ As cUnzip
Attribute m_cUZ.VB_VarHelpID = -1
Dim flag_ok As Boolean

Private Sub Client_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim s As String
    Client(Index).GetData s, vbString
    s1 = s1 & s
Dim j As Long
    j = InStr(1, s1, Chr(8), vbBinaryCompare)
    While j <> 0
        Akses Left(s1, j - 1), Index
        s1 = Mid(s1, j + 1)
        j = InStr(1, s1, Chr(8), vbBinaryCompare)
    Wend
End Sub

Sub Akses(ByVal s As String, ByVal Index As Integer)
    If Left(s, 1) = "0" Then 'Nama User
        Client(Index).Tag = Mid(s, 4)
        lname.List(NumClient - 1) = Client(Index).Tag
        lname.ItemData(NumClient - 1) = CInt(Mid(s, 2, 2))
    ElseIf Left(s, 1) = "1" Then 'Logging
        Open "d:\backup\log" & ConnectTime & ".txt" For Append As #1
        Print #1, Mid(s, 2)
        Close
        c1 lq, lp, Client(Index).Tag & " : " & Mid(s, 2) & vbCrLf
        lquery = Left(lq, lp - 1)
    ElseIf Left(s, 1) = "2" Then 'BatchQuery
        a = Batch_Q(Mid(s, 2))
        If a Then
            Client(Index).SendData "2SUKSES"
        Else
            Client(Index).SendData "2GAGAL"
        End If
    End If
End Sub

Private Sub Client_Close(Index As Integer)
    For i = 0 To lname.ListCount - 1
        If lname.ItemData(i) = Index Then
            lname.RemoveItem i
            Exit For
        End If
    Next
    Client(Index).Close
    If Index = Client.Count - 1 And Index <> 0 Then
        Unload Client(Client.Count - 1)
    End If
    NumClient = NumClient - 1
End Sub
Private Sub Client_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    For i = 0 To lname.ListCount - 1
        If lname.ItemData(i) = Index Then
            lname.RemoveItem i
            Exit For
        End If
    Next
    Client(Index).Close
    If Index = Client.Count - 1 And Index <> 0 Then
        Unload Client(Client.Count - 1)
    End If
    NumClient = NumClient - 1
End Sub

Private Function Batch_Q(ByVal s As String) As Boolean
On Error GoTo err
'    Dim AccApp As New Access.Application
'    Dim db As DAO.Database
'    Set db = AccApp.DBEngine.OpenDatabase(HOSTPATH & "\" & DBName, False, False, ";pwd=" & DBPassword)
'    AccApp.OpenCurrentDatabase HOSTPATH & "\" & DBName
'    AccApp.Run "Batch_Q", s
    Batch_Q = True
    CN.BeginTrans
    a = Split(s, vbCrLf)
    b = 0
    For i = 0 To UBound(a) - 1
        If a(i) <> "" Then
            CN.Execute a(i), b
            b = b + 1
        End If
    Next
    CN.CommitTrans
    Exit Function
err:
    CN.RollbackTrans
    Batch_Q = False
End Function


Private Sub cmdCompactDB_Click()
    CompactDB pServerDir, pDBName, pDBPassword
    d = Format(date, "ddmmyy")
    t = Format(Time, "hhmmss")
    ConnectTime = Left(d, 2) & Mid(d, 4, 2) & Right(d, 2) & Left(t, 2) & Mid(t, 4, 2) & Right(t, 2)
    ZipDB pServerDir & "a.b", "d:\backup\db" & ConnectTime & ".bbb"
End Sub

Private Sub Form_Load()
    kontinu = 0
    flag_ok = False
    d = Format(date, "ddmmyy")
    t = Format(Time, "hhmmss")
    ConnectTime = Left(d, 2) & Mid(d, 4, 2) & Right(d, 2) & Left(t, 2) & Mid(t, 4, 2) & Right(t, 2)
    ZipDB pServerDir & "a.b", "d:\backup\db" & ConnectTime & ".bbb"
    Kill pServerDir & "a.b"
    'Open "d:\backup\log" & ConnectTime & ".txt" For Output As #1
    'Close
    'Connect ConnectionString
    Server.LocalPort = 81
    Server.Listen
    Encrypt Server.LocalHostName, pServerDir & "\host.txt", "Server"
    lq = Space(10000)
    lp = 1
    c1 lq, lp, "Server " & Server.LocalHostName & " Started At " & date & ", " & Time & vbCrLf
    lquery = Left(lq, lp - 1)
    q = Space(10000)
    q1 = Space(10000)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lquery.Width = ScaleWidth - lname.Width
    lquery.Height = ScaleHeight
    lname.Height = ScaleHeight
    lname.Left = lquery.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err
    'If lname.ListCount > 0 Then
    '    Kill pServerDir & "a.zip"
    '    ZipDB pDataSource, pServerDir & "a.zip"
    '    For i = 0 To Client.Count - 1
    '        Client(i).SendData "XX"
    '        DoEvents
    '    Next
    '    mUnload = False
    '    Cancel = True
    '    Exit Sub
    'End If
    'If Format(Time, "hh:nn:ss") > "16:00:00" Or mUnload Then
    '    d = Format(date, "ddmmyy")
    '    t = Format(Time, "hhmmss")
    '    ConnectTime = Left(d, 2) & Mid(d, 4, 2) & Right(d, 2) & Left(t, 2) & Mid(t, 4, 2) & Right(t, 2)
    '    ZipDB pDataSource, "d:\backup\db" & ConnectTime & ".zip"
    '    Kill pServerDir & "host.txt"
    '    Exit Sub
    'End If
    'MsgBox "Jangan Ditutup Donk!!!"
    'Cancel = True
err:
End Sub

Sub Keluar()
On Error GoTo err
    mUnload = True
    Unload Me
err:
End Sub

Private Sub fQuit_Click()
    Keluar
End Sub

Private Sub FSET_Click()
On Error GoTo err
    For i = 0 To Client.Count - 1
        Client(i).SendData "XX"
        DoEvents
    Next
err:
End Sub

Private Sub lquery_DblClick()
    Keluar
End Sub

Private Sub lquery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub lquery_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then KeyAscii = 0
    KeyAscii = 0
End Sub
Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    Dim Found As Boolean
    For i = 0 To Client.Count - 1
        If Client(i).State = sckClosed Then
            Client(i).Accept requestID
            Found = True
            Exit For
        End If
    Next
    If Not Found Then
        Load Client(i)
        Client(i).Accept requestID
    End If
    NumClient = NumClient + 1
    Client(i).SendData "0" & i
End Sub

Private Sub Timer1_Timer()
    If Hour(Now) = 1 And Not flag_ok Then
        cmdCompactDB_Click
        flag_ok = True
        Shell "shutdown /s"
    End If
    If Hour(Now) <> 1 Then
        flag_ok = False
    End If
End Sub

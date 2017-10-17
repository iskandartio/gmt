VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Begin VB.Form FormLogin 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LOGIN"
   ClientHeight    =   3000
   ClientLeft      =   2505
   ClientTop       =   1440
   ClientWidth     =   3720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   3720
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin UsrText.IText fUsr 
      Height          =   270
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton fOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1260
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin UsrText.IText fPwd 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   476
      DataType        =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label fWarning 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID ATAU PASSWORD SALAH"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1860
      Width           =   2775
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim mCtrl As Boolean


Private Sub fOK_Click()
'On Error GoTo err
    MousePointer = vbHourglass
    a = "select bos, tipe, usr, bag1, bag2, bagSee1, bagSee2, bagAdd1, bagAdd2, bagEdit1, bagEdit2, bagDelete1, bagDelete2, bagPrint1, bagPrint2, bag3, bag4, UpdateHargaSC from m_users left join m_departemen on m_users.kddept=m_departemen.kddept where usr='" & esc(fUsr) & "' and pwd='" & esc(fPwd) & "'"
    query a
    If RS.RecordCount <> 0 Then
        pBosDepartemen = RS.Fields("bos").value
        pTipe = RS.Fields("tipe").value
        pUsr = RS.Fields("usr").value
        pBag1 = RS.Fields("bag1").value
        pBag2 = RS.Fields("bag2").value
        pBagSee1 = RS.Fields("bagSee1").value
        pBagSee2 = RS.Fields("bagSee2").value
        pBagAdd1 = RS.Fields("bagAdd1").value
        pBagAdd2 = RS.Fields("bagAdd2").value
        pBagEdit1 = RS.Fields("bagEdit1").value
        pBagEdit2 = RS.Fields("bagEdit2").value
        pBagDelete1 = RS.Fields("bagDelete1").value
        pBagDelete2 = RS.Fields("bagDelete2").value
        pBagPrint1 = RS.Fields("bagPrint1").value
        pBagPrint2 = RS.Fields("bagPrint2").value
        pBag3 = RS.Fields("bag3").value
        pBag4 = RS.Fields("bag4").value
        pUpdateHargaSC = RS.Fields("UpdateHargaSC").value
        FormLogin.Visible = False
        If pServerDir <> "" Then
            If Not CheckDir(pServerDir & "\host.txt") Then
                MsgBox "Server belum aktif"
                Unload Me
                Exit Sub
            End If
            Winsock1.RemoteHost = Decrypt(pServerDir & "\host.txt   ", "Server")
            Winsock1.RemotePort = 81
            Winsock1.Connect
        End If
        
        FormMenu.Show
        GoTo err
    End If
    fWarning = "USER ID ATAU PASSWORD SALAH"
    fUsr = ""
    fPwd = ""
err:
    MousePointer = vbDefault
End Sub
    
    Private Function CheckDir(ByVal tDir As String) As Boolean
        On Error GoTo err
        CheckDir = False
        If Dir(pServerDir & "\host.txt") <> "" Then CheckDir = True
err:
    End Function
    
    
Private Sub FOK_LostFocus()
    fWarning = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnmapNetwork pNetWorkPath
End Sub

Private Sub fPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fOK_Click
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    fWarning = ""
    pAddNo = "/" & Right(pServerDate, 2)
    pAddNoLong = Right(pServerDate, 2) * 10000
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim s As String, s1 As String
    Winsock1.GetData s, vbString
    If Left(s, 1) = "0" Then
        s1 = Mid(s, 2)
        Winsock1.Tag = s1
        Winsock1.SendData "0" & Format(FormLogin.Winsock1.Tag, "00") & pUsr & Chr(8)
    End If
End Sub

    

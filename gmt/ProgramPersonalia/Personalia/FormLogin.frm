VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Begin VB.Form FormLogin 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LOGIN"
   ClientHeight    =   3615
   ClientLeft      =   2505
   ClientTop       =   1440
   ClientWidth     =   3990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   3990
   Begin UsrText.IText fUsr 
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton fOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin UsrText.IText fPwd 
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      DataType        =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Left            =   1320
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label fWarning 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID ATAU PASSWORD SALAH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fGantiServer_Click()
    b = InputBox("Masukkan No IP Server", "IP Address Server")
    pServerName = "\\" & b
    SaveSetting App.EXEName, App.EXEName, "ServerName", pServerName
End Sub


Private Sub Form_DblClick()
    b = InputBox("Masukkan No IP Server", "IP Address Server")
    pServerName = "\\" & b
    SaveSetting App.EXEName, App.EXEName, "ServerName", pServerName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub fOK_Click()
On Error GoTo err
    a = "select * from m_users"
    query a
    MousePointer = vbHourglass
    a = "select pwd,tipe,usr,bag1,bag2,bagSee1,bagSee2,bagAdd1,bagAdd2,bagEdit1,bagEdit2,bagDelete1,bagDelete2,bagPrint1,bagPrint2,bag3,bag4,UpdateHargaSC from m_users where usr='" & esc(fUsr) & "'"
    'a = "select tipe,usr,bag1,bag2,bagSee1,bagSee2,bagAdd1,bagAdd2,bagEdit1,bagEdit2,bagDelete1,bagDelete2,bagPrint1,bagPrint2,bag3,bag4,UpdateHargaSC from m_users where usr='" & esc(fUsr) & "' and pwd='" & esc(fPwd) & "'"
    query a
    If rs.RecordCount <> 0 Then
        If rs.Fields("pwd").Value = fPwd Then
            pTipe = rs.Fields("tipe").Value
            pUsr = rs.Fields("usr").Value
            pBag1 = rs.Fields("bag1").Value
            pBag2 = rs.Fields("bag2").Value
            pBagSee1 = rs.Fields("bagSee1").Value
            pBagSee2 = rs.Fields("bagSee2").Value
            pBagAdd1 = rs.Fields("bagAdd1").Value
            pBagAdd2 = rs.Fields("bagAdd2").Value
            pBagEdit1 = rs.Fields("bagEdit1").Value
            pBagEdit2 = rs.Fields("bagEdit2").Value
            pBagDelete1 = rs.Fields("bagDelete1").Value
            pBagDelete2 = rs.Fields("bagDelete2").Value
            pBagPrint1 = rs.Fields("bagPrint1").Value
            pBagPrint2 = rs.Fields("bagPrint2").Value
            pBag3 = rs.Fields("bag3").Value
            pBag4 = rs.Fields("bag4").Value
            pUpdateHargaSC = rs.Fields("UpdateHargaSC").Value
            FormLogin.Visible = False
            FormMenu.Show
            GoTo err
        End If
    End If
    fWarning = "USER ID ATAU PASSWORD SALAH"
    fUsr = ""
    fPwd = ""
err:
    MousePointer = vbDefault
End Sub

Private Sub FOK_LostFocus()
    fWarning = ""
End Sub

Private Sub Masuk()
    Hide
    FormMenu.Show
End Sub

Private Sub fPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fOK_Click
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    fWarning = ""
End Sub


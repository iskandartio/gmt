VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Begin VB.Form FormDefault 
   BackColor       =   &H00FFC0C0&
   Caption         =   "DEFAULT"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNaikGaji 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1740
      Width           =   6975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   540
      Width           =   1095
   End
   Begin UsrText.IText txtPremiHadir 
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      DataType        =   1
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
   Begin UsrText.IText txtUangShift 
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      DataType        =   1
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
   Begin UsrText.IText txtPengaliLembur 
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      DataType        =   1
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
   Begin UsrText.IText txtAstek 
      Height          =   270
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      DataType        =   1
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula Kenaikan Gaji"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1500
      Width           =   2355
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "% Astek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pengali Lembur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uang Shift"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Premi Hadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FormDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
On Error GoTo err
Dim rs1 As New ADODB.Recordset
    BeginTransaction
    s = "select * from m_Default"
    Set rs1 = Nothing
    rs1.Open s, CN, adOpenKeyset, adLockOptimistic
    If rs1.EOF Then rs1.AddNew
    rs1!PremiHadir = txtPremiHadir.Text
    rs1!UangShift = txtUangShift.Text
    rs1!PengaliLembur = txtPengaliLembur.Text
    rs1!PersenAstek = txtAstek.Text
    rs1.Update
    s = "delete from t_NaikGaji"
    ExecMe s
    s1 = Split(txtNaikGaji.Text, vbCrLf)
    For i = 0 To UBound(s1)
        s = "insert into t_NaikGaji(Ket) values('" & s1(i) & "')"
        ExecMe s
    Next
    CommitTransaction
    MsgBox "Sukses"
    Exit Sub
err:
    RollBackTransaction
    MsgBox "Gagal"
End Sub

Private Sub Form_Load()
    s = "select * from m_Default"
    query s
    If rs.EOF Then Exit Sub
    txtPremiHadir.Text = rs!PremiHadir
    txtUangShift.Text = rs!UangShift
    txtPengaliLembur.Text = rs!PengaliLembur
    txtAstek.Text = rs!PersenAstek
    s = "select * from t_NaikGaji"
    query s
    s = ""
    Do Until rs.EOF
        s = s & vbCrLf & rs!Ket
        rs.MoveNext
    Loop
    If s <> "" Then txtNaikGaji.Text = Mid(s, 3) Else txtNaikGaji.Text = ""
End Sub

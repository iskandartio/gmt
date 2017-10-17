VERSION 5.00
Begin VB.Form FormNaikGaji 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Naik Gaji Global"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Tag             =   "15"
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtNaikGaji 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   540
      Width           =   6975
   End
   Begin VB.Label Label1 
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
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2355
   End
End
Attribute VB_Name = "FormNaikGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Load()
    s = "select * from t_NaikGaji"
    query s
    s = ""
    Do Until rs.EOF
        s = s & vbCrLf & rs!Ket
        rs.MoveNext
    Loop
    If s <> "" Then txtNaikGaji.Text = Mid(s, 3) Else txtNaikGaji.Text = ""
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    s = "delete from t_NaikGaji"
    ExecMe s
    s1 = Split(txtNaikGaji.Text, vbCrLf)
    For i = 0 To UBound(s1)
        s = "insert into t_NaikGaji(Ket) values('" & s1(i) & "')"
        ExecMe s
    Next
    CommitTransaction
    MsgBox "Sukses"
err:
    RollBackTransaction
    MsgBox "Gagal"
End Sub

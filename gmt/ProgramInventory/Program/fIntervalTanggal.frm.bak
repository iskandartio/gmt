VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Begin VB.Form fIntervalTanggal 
   BackColor       =   &H00FFC0C0&
   Caption         =   "INTERVAL TANGGAL"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton fOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      Text            =   "__/__/__"
      DataType        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
   End
   Begin UsrText.IText fAwal 
      Height          =   270
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      Text            =   "__/__/__"
      DataType        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Top             =   240
      Width           =   2475
   End
End
Attribute VB_Name = "fIntervalTanggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLaporan As String
Dim mCrystal As Boolean

Private Sub fAkhir_LostFocus()
    If cD(fAkhir) > cD(pServerDate) Or cD(fAkhir) = "A" Then fAkhir = pServerDate
End Sub


Sub LoadMe(ByVal tLaporan As String, Optional ByVal tCrystal As Boolean)
    mLaporan = tLaporan
    mCrystal = tCrystal
    Label1 = mLaporan
    Show
End Sub

Private Sub fOK_Click()
    akhir = cD(fAkhir)
    If akhir = 0 Then
        MsgBox "Tanggal Harus Diisi"
        Exit Sub
    End If
    If mCrystal Then
        'FormReport.LoadMe mLaporan & pTipe & ".rpt", cD(fAwal) & "@" & akhir
    Else
        FormPreview.LoadMe Me, mLaporan, fAwal & "@" & fAkhir
    End If
End Sub


Private Sub Form_Load()
    fAkhir = pServerDate
    fAwal = cTanggal((cD(fAkhir) \ 100) * 100 + 1)
End Sub



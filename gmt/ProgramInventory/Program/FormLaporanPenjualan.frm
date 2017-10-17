VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Begin VB.Form FormLaporanPenjualan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LAPORAN PENJUALAN"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin UsrText.IText fCustomer 
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   540
      Width           =   4035
      _ExtentX        =   7117
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
   Begin VB.OptionButton Options 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort No Surat Jalan"
      Height          =   255
      Index           =   0
      Left            =   2340
      TabIndex        =   7
      Top             =   1020
      Value           =   -1  'True
      Width           =   1875
   End
   Begin VB.OptionButton Options 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Top Customer"
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   6
      Top             =   1260
      Width           =   1875
   End
   Begin VB.CommandButton fSummary 
      Caption         =   "&SUMMARY"
      Height          =   375
      Left            =   1140
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin UsrText.IText fAwal 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   1140
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
   Begin VB.CommandButton fDetail 
      Caption         =   "&DETAIL"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   915
   End
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Top             =   1140
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "F1 untuk Pilihan"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   300
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   300
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   1035
   End
End
Attribute VB_Name = "FormLaporanPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LCustomer As Boolean
Dim mKode As String

Private Sub fAkhir_Validate(Cancel As Boolean)
    If cD(fAkhir) > cD(pServerDate) Then fAkhir = pServerDate
End Sub

Private Sub fCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        HelpMe "Nama Customer With Kode", Me
    End If
End Sub

Sub FormHelpKeyDown(ByVal tVal As String)
    a = Split(tVal, "@")
    ActiveControl.Text = a(1)
    mKode = a(0)
End Sub

Private Sub fDetail_Click()
    If esc(fCustomer) <> "" Then
        param = " Filtered: " & fCustomer
    End If
    If Options(0).Value = True Then
        FormPreview.LoadMe Me, "LaporanPenjualan", fAwal & "@" & fAkhir & "@" & fAwal & " - " & fAkhir & param & "@" & fCustomer.Text
    Else
        FormPreview.LoadMe Me, "LaporanPenjualanPerCustomer", fAwal & "@" & fAkhir & "@" & fAwal & " - " & fAkhir & param & "@" & fCustomer.Text & "@Detail"
    End If
End Sub

Private Sub Form_Activate()
    LCustomer = False
End Sub


Private Sub Form_Load()
    fAwal = "01/" & Mid(pServerDate, 4)
    fAkhir = pServerDate
    mKode = ""
End Sub

Private Sub fSummary_Click()
    If esc(fCustomer) <> "" Then
        param = " Filtered: " & fCustomer
    End If
    If Options(0).Value = True Then
        FormPreview.LoadMe Me, "SummaryPenjualan", fAwal & "@" & fAkhir & "@" & fAwal & " - " & fAkhir & param & "@" & fCustomer.Text
    Else
        FormPreview.LoadMe Me, "LaporanPenjualanPerCustomer", fAwal & "@" & fAkhir & "@" & fAwal & " - " & fAkhir & param & "@" & fCustomer.Text & "@Summary"
    End If
End Sub




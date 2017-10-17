VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Begin VB.Form fIntervalTanggal 
   BackColor       =   &H00FFC0C0&
   Caption         =   "INTERVAL TANGGAL"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKartuAbsensi 
      Caption         =   "KARTU ABSENSI"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbsensi 
      Caption         =   "REKAP ABSENSI"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRekapRinciUang 
      Caption         =   "REKAP RINCI UANG"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRekapKecil 
      Caption         =   "REKAP KECIL"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSummary 
      Caption         =   "REKAP GAJI SUMMARY"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdStruk 
      Caption         =   "STRUK GAJI"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cmbDepartemen 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton fOK 
      Caption         =   "REKAP GAJI"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin UsrText.IText fAwal 
      Height          =   270
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   476
      Text            =   "__/__/____"
      DataType        =   6
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
   Begin UsrText.IText txtNIK 
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NIK"
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
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Departemen"
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
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      TabIndex        =   2
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "fIntervalTanggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLaporan As String
Dim mCrystal As Boolean


Private Sub cmbDepartemen_DropDown()
    s = "select distinct Departemen from m_Karyawan where Status=0"
    query s
    Do Until rs.EOF
        cmbDepartemen.AddItem rs.Fields(0).Value
        rs.MoveNext
    Loop
End Sub

Private Sub cmdAbsensi_Click()
    LapAbsensi "RekapAbsensi.rpt"
End Sub

Private Sub cmdKartuAbsensi_Click()
    LapAbsensi "KartuAbsensi.rpt"
End Sub

Private Sub cmdRekapKecil_Click()
    LapGaji "RekapKecil.rpt"
End Sub

Private Sub cmdRekapRinciUang_Click()
    LapGaji "RekapRinciUang.rpt"
End Sub

Private Sub cmdStruk_Click()
    LapGaji "StrukGaji.rpt"
End Sub

Private Sub cmdSummary_Click()
    LapGaji "RekapGajiSummary.rpt"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Sub LoadMe(ByVal tLaporan As String, Optional ByVal tCrystal As Boolean)
    mLaporan = tLaporan
    mCrystal = tCrystal
    Label1 = mLaporan
    Show
End Sub

Private Sub fOK_Click()
    LapGaji "RekapGaji.rpt"
End Sub

Sub LapGaji(ByVal a As String)
Dim MyFilter As String
    MyFilter = "{Command.Tanggal}=" & cD(fAwal, True)
    If cmbDepartemen.Text <> "" Then MyFilter = MyFilter & " and {Command.gDepartemen}='" & esc(cmbDepartemen.Text) & "'"
    If txtNIK.Text <> "" Then MyFilter = MyFilter & " and {Command.gNIK}=" & txtNIK.Text
    frmReport.LoadMe a, "", MyFilter
End Sub

Sub LapAbsensi(ByVal a As String)
Dim MyFilter As String
    MyFilter = "{?Tanggal}=" & cD(fAwal, True)
    If cmbDepartemen.Text <> "" Then MyFilter = MyFilter & " and {Command.Departemen}='" & esc(cmbDepartemen.Text) & "'"
    If txtNIK.Text <> "" Then MyFilter = MyFilter & " and {Command.NIK}=" & txtNIK.Text
    frmReport.LoadMe a, cD(fAwal, True), MyFilter
End Sub


Private Sub Form_Load()
    s = "select max(Tanggal) from t_Gaji"
    query s
    fAwal = cTanggal(rs.Fields(0).Value, True)
End Sub



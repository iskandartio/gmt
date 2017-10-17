VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Object = "{F6D22ACD-8630-4FE1-97C4-D56AB4AD4DEA}#1.0#0"; "UsrTrueCombo.ocx"
Begin VB.Form FormBulanan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "TRANSAKSI KARYAWAN"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Tag             =   "54"
   WindowState     =   2  'Maximized
   Begin UsrText.IText txtPeriode 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Top             =   600
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
   Begin UsrTrueCombo.ITrueCombo cmbNIK 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   960
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
   End
   Begin UsrText.IText txtNama 
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
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
   Begin UsrText.IText txtS 
      Height          =   270
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
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
   Begin UsrText.IText txtI 
      Height          =   270
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
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
   Begin UsrText.IText txtM 
      Height          =   270
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
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
   Begin UsrText.IText txtHariShift 
      Height          =   270
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
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
   Begin UsrText.IText txtJamLembur 
      Height          =   270
      Left            =   1320
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      DataType        =   2
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
   Begin UsrText.IText txtTunjKhusus 
      Height          =   270
      Left            =   1320
      TabIndex        =   9
      Top             =   3840
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
   Begin UsrText.IText txtTunjLain 
      Height          =   270
      Left            =   1320
      TabIndex        =   10
      Top             =   4200
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
   Begin UsrText.IText txtTunjJabatan 
      Height          =   270
      Left            =   1320
      TabIndex        =   8
      Top             =   3480
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
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tunj Jabatan"
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
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Periode"
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
      TabIndex        =   22
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tunj Lain2"
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
      TabIndex        =   19
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tunj Khusus"
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
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Jam Lembur"
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
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Shift"
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
      TabIndex        =   16
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mangkir"
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
      TabIndex        =   15
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Izin"
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
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sakit"
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
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
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
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "FormBulanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDBObject.XArrayDB
Dim mTitle As String
Dim Loaded As Boolean

Private Sub cmbNIK_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs1() As Variant
    If Not Loaded Then
        s = "select NIK, Nama from m_Karyawan where StatusKerja='A' and Status=0 order by NIK"
        query s
        rs1 = rs.GetRows
        cmbNIK.SetHeader "NIK@Nama"
        cmbNIK.SetWidth "1000@2000"
        cmbNIK.SetType "Integer@String"
        cmbNIK.SetDB rs1
        Loaded = True
    End If
End Sub

Private Sub cmbNIK_Validate(Cancel As Boolean)
On Error GoTo err
    s = "select m_Karyawan.NIK as tNIK, * from m_Karyawan left join t_Gaji on t_Gaji.gNIK=m_Karyawan.NIK where m_Karyawan.NIK=" & cmbNIK.Text
    query s
    If Not rs.EOF Then
        cmbNIK.Text = Format(rs!tNIK, "000000")
        txtNama.Text = rs!Nama
        txtS.Text = rs!s & ""
        txtI.Text = rs!i & ""
        txtM.Text = rs!m & ""
        txtHariShift.Text = rs!HariShift & ""
        txtTunjKhusus.Text = rs!TKhusus & ""
        txtTunjLain.Text = rs!TLain & ""
        txtTunjJabatan.Text = rs!TJabatan & ""
        txtJamLembur.Text = rs!JamLembur & ""
    Else
        cmbNIK.Text = ""
        txtNama.Text = ""
        txtS.Text = ""
        txtI.Text = ""
        txtM.Text = ""
        txtHariShift.Text = ""
        txtTunjJabatan.Text = ""
        txtTunjKhusus.Text = ""
        txtTunjLain.Text = ""
        txtJamLembur.Text = ""
    End If
err:
End Sub

Private Sub cmdSave_Click()
Dim rs1 As New ADODB.Recordset
    s = "select * from m_Karyawan where NIK=" & cmbNIK.Text
    query s
    s = "select * from t_Gaji where Tanggal=" & cD(txtPeriode.Text, True) & " and gNIK=" & cmbNIK.Text
    Set rs1 = Nothing
    rs1.Open s, CN, adOpenKeyset, adLockOptimistic
    If rs1.EOF Then rs1.AddNew
    rs1!s = cNum(txtS.Text)
    rs1!i = cNum(txtI.Text)
    rs1!m = cNum(txtM.Text)
    rs1!HariShift = cNum(txtHariShift.Text)
    rs1!TJabatan = cNum(txtTunjJabatan.Text)
    rs1!TKhusus = cNum(txtTunjKhusus.Text)
    rs1!TLain = cNum(txtTunjLain.Text)
    rs1!JamLembur = cNum(txtJamLembur.Text)
    rs1!gGaji = rs!Gaji
    rs1!gDepartemen = rs!Departemen
    rs1!gUangMakan = rs!UangMakan
    rs1!gUangTransport = rs!UangTransport
    rs1!gJabatan = rs!Jabatan
    rs1!gShift = rs!Shift
    rs1!gTunjanganMasaKerja = rs!TunjanganMasaKerja
    rs1.Update
    MsgBox "Data tersimpan"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
End Sub

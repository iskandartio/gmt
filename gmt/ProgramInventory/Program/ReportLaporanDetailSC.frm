VERSION 5.00
Begin VB.Form ReportLaporanDetailSC 
   Caption         =   "Detail SC"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dSumJumlahKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@JUMLAH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   15
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label dPageNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "@HALAMAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label dSatKecil 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SATUAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label dJumlahKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@JUMLAH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label dTanggalSJ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@TANGGAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label dNoSJ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NOSJ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label dCustomer 
      Caption         =   "@CUSTOMER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label dParams 
      Caption         =   "@NoContract"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label dNamaBarang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   480
      X2              =   11471
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label dHead 
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "JUMLAH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9600
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label dHead 
      Caption         =   "NO SJ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label dHead 
      Caption         =   "NAMA BARANG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   480
      X2              =   11471
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label dHead 
      Caption         =   "DETAIL SALES CONTRACT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label dHead 
      Caption         =   "CUSTOMER:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "PT GEMILANG MAJU TEXINDOTAMA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "ReportLaporanDetailSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim SumJumlahKG As Double
Dim mTotalPage As Integer
Dim mObj As Object
Dim mParam() As String
Dim iNamaCustomer As Integer
Dim iNamaBarang As Integer
Dim iNoSJ As Integer
Dim iTanggal As Integer
Dim iJumlah As Integer
Dim iSatKecil As Integer

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    iNamaCustomer = 0
    iNamaBarang = 1
    iNoSJ = 2
    iTanggal = 3
    iJumlah = 4
    iSatKecil = 5
    mParam = Split(tParams, "@")
    a = "select Nama, Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as NamaBarang, t_SPPDetail~.NoSJ, TanggalDetail, t_SPPDetail~.JumlahKG, SatKecil from (t_SPPDetail~ left join m_Stock~ on m_Stock~.IdStock=t_SPPDetail~.IdStock) left join m_Customer on m_Customer.Kode=t_SPPDetail~.KodeCustomerDetail where NoSC='" & tParams & "' and statusDetail>1 order by t_SPPDetail~.NoSJ, t_SPPDetail~.IdSPP"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    res = RS.GetRows
    FormPreview.Picture1.Height = 15500
    mTotalPage = 0
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    m = IIf(TypeName(mObj) = "Printer" And Not tSign, tPlus, 0)
    PaintHeader mParam(0), mObj, dParams
    PaintDetail res(iNamaCustomer, i), mObj, dCustomer, dCustomer.Top, tMax
    t = dNamaBarang.Top
    SumJumlahKG = 0
    Do
        tMax = 0
        PaintDetail res(iNamaBarang, i), mObj, dNamaBarang, t, tMax, tSign
        PaintDetail res(iNoSJ, i), mObj, dNoSJ, t, tMax, tSign
        PaintDetail cTanggal(res(iTanggal, i)), mObj, dTanggalSJ, t, tMax, tSign
        PaintDetail cDecimal(res(iJumlah, i)), mObj, dJumlahKG, t, tMax, tSign
        PaintDetail res(iSatKecil, i), mObj, dSatKecil, t, tMax, tSign
        t = t + tMax * dNamaBarang.Height
        If t > 15000 + m And tSign Then Exit Do
        SumJumlahKG = SumJumlahKG + res(iJumlah, i)
        If i = UBound(res, 2) Then
            PrintFooter t, i, tSign
            i = i + 1
            mTotalPage = tPage
            FormPreview.SetTotalPage mTotalPage
            Exit Do
        ElseIf i + 1 > iLast And Not tSign Then
            Exit Do
        End If
        i = i + 1
    Loop
    PrintData = i
End Function

Private Sub PrintFooter(t As Single, ByVal i As Long, ByVal tSign As Boolean)
Dim tMax As Byte
    If Not tSign Then PaintLine mObj, Lines(0).x1, t, Lines(0).x2, t
    PaintDetail cDecimal(SumJumlahKG), mObj, dSumJumlahKG, t, tMax, tSign
    PaintDetail res(iSatKecil, i), mObj, dSatKecil, t, tMax, tSign
    t = t + tMax * dSumJumlahKG.Height
    t = t + 100
End Sub

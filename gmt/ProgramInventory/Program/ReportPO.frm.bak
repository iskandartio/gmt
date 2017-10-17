VERSION 5.00
Begin VB.Form ReportPO 
   Caption         =   "PURCHASE ORDER"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dAlamat 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@Alamat"
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
      Left            =   420
      TabIndex        =   33
      Top             =   1560
      Width           =   6855
   End
   Begin VB.Line dFooter 
      X1              =   720
      X2              =   1920
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label dFoot 
      Caption         =   "PURCHASING MANAGER"
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
      Left            =   420
      TabIndex        =   32
      Top             =   5460
      Width           =   2415
   End
   Begin VB.Label dGrandTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GRANDTOTAL"
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
      Left            =   9720
      TabIndex        =   31
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label dFoot 
      Alignment       =   1  'Right Justify
      Caption         =   "GRAND TOTAL"
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
      Left            =   8160
      TabIndex        =   30
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label dPPN 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PPN"
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
      Left            =   9720
      TabIndex        =   29
      Top             =   4620
      Width           =   1575
   End
   Begin VB.Label dFoot 
      Alignment       =   1  'Right Justify
      Caption         =   "PPN"
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
      Left            =   8160
      TabIndex        =   28
      Top             =   4620
      Width           =   1455
   End
   Begin VB.Label dHead 
      Caption         =   "JL. GATOT SUBROTO KM 6,5 JATAKE TANGERANG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   27
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label dHead 
      Caption         =   "TELP: (021)-5926688, FAX: (021)-5908840"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   480
      TabIndex        =   26
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label dFoot 
      Caption         =   "HORMAT KAMI,"
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
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label dSatuan 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SAT"
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
      Left            =   7200
      TabIndex        =   24
      Top             =   2505
      Width           =   1095
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "SATUAN"
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
      Index           =   11
      Left            =   7320
      TabIndex        =   23
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "HARGA"
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
      Index           =   10
      Left            =   8700
      TabIndex        =   22
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label dFoot 
      BackStyle       =   0  'Transparent
      Caption         =   "LAY KHEK CHAN"
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
      Index           =   4
      Left            =   420
      TabIndex        =   21
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label dFoot 
      Alignment       =   1  'Right Justify
      Caption         =   "SUB TOTAL"
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
      Index           =   1
      Left            =   8160
      TabIndex        =   20
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label dSubTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SUMTOTAL"
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
      Left            =   9720
      TabIndex        =   19
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11471
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label dTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@TOTAL"
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
      Left            =   9720
      TabIndex        =   18
      Top             =   2505
      Width           =   1575
   End
   Begin VB.Label dHarga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@HARGA"
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
      TabIndex        =   17
      Top             =   2505
      Width           =   1455
   End
   Begin VB.Label dJumlah 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@Jumlah"
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
      Left            =   6120
      TabIndex        =   16
      Top             =   2505
      Width           =   975
   End
   Begin VB.Label dNamaBarang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NAMA BARANG"
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
      Left            =   1500
      TabIndex        =   15
      Top             =   2505
      Width           =   4635
   End
   Begin VB.Label dNoUrut 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NO"
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
      TabIndex        =   14
      Top             =   2505
      Width           =   975
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   480
      X2              =   11471
      Y1              =   2460
      Y2              =   2460
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
      Index           =   8
      Left            =   6060
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL"
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
      Index           =   12
      Left            =   10500
      TabIndex        =   12
      Top             =   2160
      Width           =   735
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
      Index           =   9
      Left            =   1500
      TabIndex        =   11
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label dHead 
      Alignment       =   2  'Center
      Caption         =   "NO PO"
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
      Index           =   7
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   480
      X2              =   11471
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label dMataUang 
      Caption         =   "@MATA UANG"
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
      Left            =   9720
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label dJatuhTempo 
      Caption         =   "@JATUH TEMPO"
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
      Left            =   9720
      TabIndex        =   8
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label dTanggal 
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
      Left            =   9720
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label dSupplier 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SUPPLIER"
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
      Left            =   420
      TabIndex        =   6
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label dHead 
      Caption         =   "MATA UANG"
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
      Left            =   8160
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label dHead 
      Caption         =   "TEMPO BAYAR"
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
      Left            =   8160
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
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
      Left            =   8160
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "PURCHASE ORDER"
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
      Index           =   3
      Left            =   6240
      TabIndex        =   2
      Top             =   180
      Width           =   5055
   End
   Begin VB.Label dHead 
      Caption         =   "KEPADA YTH,"
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
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label dHead 
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   180
      Width           =   5055
   End
End
Attribute VB_Name = "ReportPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim res() As Variant
Dim iSupplier As Integer
Dim iAlamat As Integer
Dim iKota As Integer
Dim iTelepon As Integer
Dim iFax As Integer
Dim iTanggal As Integer
Dim iMataUang As Integer
Dim iTempo As Integer
Dim iNoPO As Integer
Dim iNamaBarang As Integer
Dim iJumlah As Integer
Dim iSatuan As Integer
Dim iHarga As Integer
Dim iTotal As Integer
Dim iPPN As Integer

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object, ByVal tSQL As String)
    a = tSQL
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    iSupplier = 0
    iAlamat = 1
    iKota = 2
    iTelepon = 3
    iFax = 4
    iTanggal = 5
    iMataUang = 6
    iTempo = 7
    iNoPO = 8
    iNamaBarang = 9
    iJumlah = 10
    iSatuan = 11
    iHarga = 12
    iTotal = 13
    iPPN = 14
    res = RS.GetRows
    Set mObj = obj
    FormPreview.SetTotalPage -1
    FormPreview.Picture1.Height = 7500
End Sub

Sub PrintData()
On Error Resume Next
    i = 0
    PaintHeader res(iSupplier, i), mObj, dSupplier
    Alamat = res(iAlamat, i) & " " & res(iKota, i) & IIf(Trim(res(iTelepon, i)) <> "", " Telp. " & res(iTelepon, i), "") & IIf(Trim(res(iFax, i)) <> "", " Fax. " & res(iFax, i), "")
    PaintHeader Alamat & " " & res(iKota, i), mObj, dAlamat
    PaintHeader cTanggal(res(iTanggal, i)), mObj, dTanggal
    PaintHeader res(iMataUang, i), mObj, dMataUang
    PaintHeader res(iTempo, i), mObj, dJatuhTempo
    Dim t As Single
    Dim tMax As Byte
    Dim SumTotal As Double
    SumTotal = 0
    t = dNoUrut.Top
    For i = 0 To RS.RecordCount - 1
        tMax = 1
        PaintDetail res(iNoPO, i), mObj, dNoUrut, t, tMax
        PaintDetail res(iNamaBarang, i), mObj, dNamaBarang, t, tMax
        PaintDetail cDecimal(res(iJumlah, i)), mObj, dJumlah, t, tMax
        PaintDetail res(iSatuan, i), mObj, dSatuan, t, tMax
        PaintDetail cDecimal(res(iHarga, i)), mObj, dHarga, t, tMax
        PaintDetail cDecimal(res(iTotal, i)), mObj, dTotal, t, tMax
        SumTotal = SumTotal + res(iTotal, i)
        t = t + tMax * dNoUrut.Height
        RS.MoveNext
    Next
    t = FormPreview.Picture1.Height - dFooter.y1
    PaintLine mObj, Line1.x1, t - 50 + dSubTotal.Top, Line1.x2, t - 50 + dSubTotal.Top
    PaintDetail cDecimal(SumTotal), mObj, dSubTotal, t + dSubTotal.Top, tMax
    PaintDetail cDecimal(res(iPPN, 0) / 10 * SumTotal), mObj, dPPN, t + dPPN.Top, tMax
    PaintDetail cDecimal((1 + res(iPPN, 0) / 10) * SumTotal), mObj, dGrandTotal, t + dGrandTotal.Top, tMax
    FormPreview.fFirst.Enabled = False
    FormPreview.fLast.Enabled = False
    FormPreview.fPrev.Enabled = False
    FormPreview.fNext.Enabled = False
    FormPreview.fPage = "1"
    FormPreview.fPage.Enabled = False
    FormPreview.fPagesPrint = "1"
    FormPreview.fPagesPrint.Enabled = False
End Sub


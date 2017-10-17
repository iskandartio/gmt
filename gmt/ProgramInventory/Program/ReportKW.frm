VERSION 5.00
Begin VB.Form ReportKW 
   Caption         =   "KWITANSI"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.Line FootLines 
      Index           =   1
      X1              =   8040
      X2              =   9960
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label dTanggalCetak 
      Caption         =   "CETAK      :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   41
      Top             =   6360
      Width           =   4995
   End
   Begin VB.Line dFooterLine 
      Tag             =   "3000"
      X1              =   0
      X2              =   9420
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Label dFoot 
      Caption         =   "HIJAU       : MARKETING"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   40
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label dFoot 
      Caption         =   "KUNING    : ACCOUNTING"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   39
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label dFoot 
      Caption         =   "MERAH     : FINANCE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   38
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label dFoot 
      Caption         =   "PUTIH       : CUSTOMER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   37
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label dFoot 
      Caption         =   "KETERANGAN:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   36
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Line FootLines 
      Index           =   0
      X1              =   480
      X2              =   11520
      Y1              =   4995
      Y2              =   4995
   End
   Begin VB.Label dTTD 
      Alignment       =   2  'Center
      Caption         =   "(BINGRIANTO)"
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
      Left            =   7800
      TabIndex        =   3
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label dCompany 
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
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label dLTangerang 
      Alignment       =   2  'Center
      Caption         =   "TANGERANG, 10/01/06"
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
      Left            =   7920
      TabIndex        =   35
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11475
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Label dMsg 
      Caption         =   "Pembayaran dengan Cek/Giro harap ditulis atas nama PT. GEMILANG MAJU TEXINDOTAMA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   34
      Top             =   4740
      Width           =   8175
   End
   Begin VB.Label dLTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL/SJ"
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
      Left            =   5160
      TabIndex        =   33
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   11475
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label dTerbilang 
      Caption         =   "@TERBILANG"
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
      TabIndex        =   32
      Top             =   4500
      Width           =   10815
   End
   Begin VB.Label dLGrandTotal 
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
      Left            =   6840
      TabIndex        =   31
      Top             =   3900
      Width           =   1695
   End
   Begin VB.Label dMataUang 
      Caption         =   "@MATAUANG"
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
      Left            =   9000
      TabIndex        =   30
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label dTanggalKW 
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
      Height          =   375
      Left            =   9000
      TabIndex        =   29
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label dNoKW 
      Caption         =   "@NOKW"
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
      Left            =   9000
      TabIndex        =   28
      Top             =   1200
      Width           =   2175
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
      Left            =   480
      TabIndex        =   27
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Label dAlamat 
      Caption         =   "@ALAMAT"
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
      TabIndex        =   26
      Top             =   1680
      Width           =   6615
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
      Index           =   1
      Left            =   7320
      TabIndex        =   25
      Top             =   1680
      Width           =   1575
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
      Index           =   6
      Left            =   7320
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label dHead 
      Caption         =   "NO KWITANSI"
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
      Left            =   7320
      TabIndex        =   23
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label dHead 
      Caption         =   "NAMA DAN ALAMAT CUSTOMER"
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
      TabIndex        =   22
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label dCompany 
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
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label dSat2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S.CONES"
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
      Left            =   7680
      TabIndex        =   21
      Top             =   3540
      Width           =   855
   End
   Begin VB.Label dSumTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9720
      TabIndex        =   20
      Top             =   3540
      Width           =   1575
   End
   Begin VB.Label dMataUang2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@MATA UANG"
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
      Left            =   8520
      TabIndex        =   19
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Label dSumJumlahKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@JUMLAH"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Label dSatKecil 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S.CONES"
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
      Left            =   7680
      TabIndex        =   17
      Top             =   3060
      Width           =   855
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
      Left            =   6480
      TabIndex        =   16
      Top             =   3060
      Width           =   1095
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
      Left            =   8520
      TabIndex        =   15
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label dTanggalSJ 
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
      Left            =   6000
      TabIndex        =   14
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label dNoSJ 
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
      Left            =   1440
      TabIndex        =   13
      Top             =   2700
      Width           =   1815
   End
   Begin VB.Label dLTanggalSJ 
      Caption         =   "TANGGAL:"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label dTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
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
      TabIndex        =   11
      Top             =   3060
      Width           =   1575
   End
   Begin VB.Label dJenis 
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
      Left            =   840
      TabIndex        =   10
      Top             =   3060
      Width           =   5535
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   480
      X2              =   11471
      Y1              =   2640
      Y2              =   2640
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
      Index           =   4
      Left            =   9120
      TabIndex        =   9
      Top             =   2340
      Width           =   735
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
      Index           =   5
      Left            =   10560
      TabIndex        =   8
      Top             =   2340
      Width           =   735
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
      Index           =   3
      Left            =   7440
      TabIndex        =   7
      Top             =   2340
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
      Left            =   840
      TabIndex        =   6
      Top             =   2340
      Width           =   2175
   End
   Begin VB.Label dLNoSJ 
      Caption         =   "NO SJ:"
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
      Left            =   600
      TabIndex        =   5
      Top             =   2700
      Width           =   735
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   480
      X2              =   11471
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "KWITANSI"
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
      Left            =   9600
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label dCompanyHead 
      Caption         =   "PT GEMILANG MAJU TEXINDOTAMA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "ReportKW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SumJumlahKG As Double
Dim SumTotal As Double
Dim GrandTotal As Double
Dim mObj As Object
Dim iTbl As New ClassProperties
Dim mParams As String
Dim res() As Variant

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object, Optional tSQL As String)
Dim param() As String
Dim tField As String
    param = Split(tParams, "@")
    mParams = tParams
    If InStr(tParams, "PrintNamaPenerima") > 0 Then
        tField = "NamaPenerimaDetail as Nama, AlamatPenerimaDetail as Alamat, "
    Else
        tField = "Nama, Alamat, "
    End If
    If tSQL = "" Then
        a = "select NoKWDetail, TanggalKWDetail, MataUangDetail, " & tField & "Kota, t_SPPDetail~.NoSJ, TanggalDetail, Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as JenisBarang, t_SPPDetail~.JumlahKG, SatKecil, Harga, t_SPPDetail~.JumlahKG*Harga as Total from (t_SPPDetail~ left join m_Stock~ on m_Stock~.IdStock=t_SPPDetail~.IdStock) left join m_Customer on m_Customer.Kode=t_SPPDetail~.KodeCustomerDetail where NoKWDetail='" & param(0) & "' order by t_SPPDetail~.NoSJ, t_SPPDetail~.IdSPP"
    Else
        a = tSQL
    End If
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    Set mObj = obj
    res = RS.GetRows
    For i = 0 To RS.Fields.Count - 1
        iTbl(RS.Fields(i).Name) = i
    Next
    FormPreview.SetPageHeight 15500
    FormPreview.SetTotalPage -1
        If pTipe = "PE" Then
        dFoot(2).Caption = "KUNING    : ACCOUNTING"
        dFoot(3).Caption = "MERAH     : FINANCE"
        dFoot(4).Caption = "HIJAU       : MARKETING"
    ElseIf pTipe = "DTY" Then
        dFoot(2).Caption = "KUNING    : MARKETING"
        dFoot(3).Caption = "MERAH     : ACCOUNTING"
        dFoot(4).Caption = "HIJAU       : FINANCE"
    End If
End Sub

Sub PrintData()
Dim tMax As Byte
Dim PF As Boolean
    If InStr(mParams, "HideCompany") > 0 Then
        PaintHeader "", mObj, dCompanyHead
    Else
        PaintHeader dCompanyHead, mObj, dCompanyHead
    End If
    PaintHeader res(iTbl("Nama"), 0), mObj, dCustomer
    If InStr(mParams, "PrintNamaPenerima") > 0 Then
        PaintHeader res(iTbl("Alamat"), 0), mObj, dAlamat
    Else
        PaintHeader res(iTbl("Alamat"), 0) & " " & res(iTbl("Kota"), 0), mObj, dAlamat
    End If
    PaintHeader res(iTbl("NoKWDetail"), 0), mObj, dNoKW
    PaintHeader cTanggal(res(iTbl("TanggalKWDetail"), 0)), mObj, dTanggalKW
    PaintHeader res(iTbl("MataUangDetail"), 0), mObj, dMataUang
    MataUang = res(iTbl("MataUangDetail"), 0)
    tanggalkw = res(iTbl("TanggalKWDetail"), 0)
    seb = ""
    t = dNoSJ.Top
    tMax = 0
    GrandTotal = 0
    For i = 0 To UBound(res, 2)
        If res(iTbl("NoSJ"), i) <> seb Then
            PaintDetail dLNoSJ, mObj, dLNoSJ, t, tMax
            PaintDetail res(iTbl("NoSJ"), i), mObj, dNoSJ, t, tMax
            PaintDetail dLTanggalSJ, mObj, dLTanggalSJ, t, tMax
            PaintDetail cTanggal(res(iTbl("TanggalDetail"), i)), mObj, dTanggalSJ, t, tMax
            t = t + dLNoSJ.Height
            seb = res(iTbl("NoSJ"), i)
        End If
        SatKecil = res(iTbl("SatKecil"), i)
        PaintDetail res(iTbl("JenisBarang"), i), mObj, dJenis, t, tMax
        PaintDetail cDecimal(res(iTbl("JumlahKG"), i)), mObj, dJumlahKG, t, tMax
        PaintDetail res(iTbl("SatKecil"), i), mObj, dSatKecil, t, tMax
        PaintDetail cDecimal(res(iTbl("Harga"), i)), mObj, dHarga, t, tMax
        PaintDetail cDecimal(res(iTbl("Total"), i)), mObj, dTotal, t, tMax
        t = t + dJenis.Height * tMax
        SumJumlahKG = SumJumlahKG + res(iTbl("JumlahKG"), i)
        SumTotal = SumTotal + res(iTbl("Total"), i)
        GrandTotal = GrandTotal + res(iTbl("Total"), i)
        PF = False
        If i = UBound(res, 2) Then
            PF = True
        ElseIf res(iTbl("NoSJ"), i + 1) <> seb Then
            PF = True
        End If
        If PF Then
            PaintLine mObj, Line2.x1, t, Line2.x2, t
            t = t + 25
            PaintDetail dLTotal, mObj, dLTotal, t, tMax
            PaintDetail cDecimal(SumJumlahKG), mObj, dSumJumlahKG, t, tMax
            PaintDetail res(iTbl("SatKecil"), i), mObj, dSat2, t, tMax
            PaintDetail cDecimal(SumTotal), mObj, dSumTotal, t, tMax
            PaintDetail MataUang, mObj, dMataUang2, t, tMax
            t = t + dSumJumlahKG.Height + 100
            SumJumlahKG = 0
            SumTotal = 0
        End If
    Next
    't = t + 100
    PaintDetail dLGrandTotal, mObj, dLGrandTotal, t, tMax
    PaintDetail cDecimal(GrandTotal), mObj, dSumTotal, t, tMax
    PaintDetail MataUang, mObj, dMataUang2, t, tMax
    t = t + dSumJumlahKG.Height
    PaintLine mObj, Line1.x1, t, Line1.x2, t
    t = t + 100
    a = "select Nama from m_matauang where Kode='" & MataUang & "'"
    query a
    PaintDetail Terbilang(GrandTotal, RS.Fields(0).Value), mObj, dTerbilang, t, tMax
    t = t + dTerbilang.Height
    
    If InStr(mParams, "HideCompany") = 0 Then
        PaintDetail dMsg, mObj, dMsg, t, tMax
    End If
    
    t = FormPreview.Picture1.Height - dFooterLine.Tag - dFooterLine.y1 + dLTangerang.Top
    PaintDetail "TANGERANG: " & date, mObj, dLTangerang, t, tMax
    t = FormPreview.Picture1.Height - dFooterLine.Tag - dFooterLine.y1 + dTanggalCetak.Top
    PaintDetail "CETAK      : " & date & " " & Time, mObj, dTanggalCetak, t, tMax
    If InStr(mParams, "HideCompany") = 0 Then
        t = FormPreview.Picture1.Height - dFooterLine.Tag - dFooterLine.y1 + dTTD.Top
        PaintDetail dTTD, mObj, dTTD, t, tMax
    End If
    
    FormPreview.fFirst.Enabled = False
    FormPreview.fLast.Enabled = False
    FormPreview.fPrev.Enabled = False
    FormPreview.fNext.Enabled = False
    FormPreview.fPage = "1"
    FormPreview.fPage.Enabled = False
    FormPreview.fPagesPrint = "1"
    FormPreview.fPagesPrint.Enabled = False
End Sub

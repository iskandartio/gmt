VERSION 5.00
Begin VB.Form ReportLaporanPenjualan 
   Caption         =   "LAPORAN PENJUALAN"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dKet 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@Ket"
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
      Left            =   3360
      TabIndex        =   25
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label dDP 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@DP"
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
      Left            =   2520
      TabIndex        =   24
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label dNoSC 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NOSC"
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
      TabIndex        =   23
      Top             =   2520
      Width           =   1815
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
      TabIndex        =   22
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label dSat2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SATUAN"
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
      Left            =   7800
      TabIndex        =   21
      Top             =   2520
      Width           =   735
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
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label dMataUang 
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
      Top             =   2520
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
      Left            =   6480
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
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
      Left            =   7800
      TabIndex        =   17
      Top             =   2040
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
      Left            =   6600
      TabIndex        =   16
      Top             =   2040
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
      Top             =   2040
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
      Left            =   10200
      TabIndex        =   14
      Top             =   1680
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
      Left            =   7200
      TabIndex        =   13
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
      Left            =   1920
      TabIndex        =   12
      Top             =   1680
      Width           =   4455
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
      Left            =   9120
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label dParams 
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
      Left            =   480
      TabIndex        =   10
      Top             =   720
      Width           =   7215
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11471
      Y1              =   2880
      Y2              =   2880
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
      TabIndex        =   9
      Top             =   2040
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
      TabIndex        =   8
      Top             =   2040
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
      TabIndex        =   7
      Top             =   1200
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
      TabIndex        =   6
      Top             =   1200
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1200
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
      Left            =   6480
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   480
      X2              =   11471
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label dHead 
      Caption         =   "LAPORAN PENJUALAN"
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
      Width           =   3375
   End
   Begin VB.Label dLCustomer 
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
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
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
Attribute VB_Name = "ReportLaporanPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim SumJumlahKG As Double
Dim SumTotal As Double
Dim mObj As Object
Dim mParam() As String
Dim mNoSC As String
Dim iKeterangan As Integer
Dim iNoSC As Integer
Dim iDP As Integer
Dim iMataUang As Integer
Dim iNamaCustomer As Integer
Dim iNoSJ As Integer
Dim iTanggal As Integer
Dim iNamaBarang As Integer
Dim iJumlah As Integer
Dim iSatKecil As Integer
Dim iHarga As Integer
Dim iTotal As Integer

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
Dim tKodeCustomer As Long
    iKeterangan = 0
    iNoSC = 1
    iDP = 2
    iMataUang = 3
    iNamaCustomer = 4
    iNoSJ = 5
    iTanggal = 6
    iNamaBarang = 7
    iJumlah = 8
    iSatKecil = 9
    iHarga = 10
    iTotal = 11
    mParam = Split(tParams, "@")
    If cD(mParam(0)) <> "A" Then MyFilter = " and TanggalDetail >= " & cD(mParam(0))
    If cD(mParam(1)) <> "A" Then MyFilter = MyFilter & " and TanggalDetail <= " & cD(mParam(1))
    If esc(mParam(3)) <> "" Then
        a = "select Kode from m_Customer where Nama='" & mParam(3) & "'"
        query a
        If RS.EOF Then Exit Sub
        MyFilter = MyFilter & " and KodeCustomerDetail=" & RS.Fields(0).value
    End If
    a = "select KeteranganSPPHeader, NoSC, DP, MataUangDetail, Nama, t_SPPDetail~.NoSJ, TanggalDetail, Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as JenisBarang, t_SPPDetail~.JumlahKG, SatKecil, Harga, t_SPPDetail~.JumlahKG*Harga as Total from (t_SPPDetail~ left join m_Stock~ on m_Stock~.IdStock=t_SPPDetail~.IdStock) left join m_Customer on m_Customer.Kode=t_SPPDetail~.KodeCustomerDetail where statusDetail>1 " & MyFilter & " order by t_SPPDetail~.NoSJ, t_SPPDetail~.IdSPP"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    res = RS.GetRows
    FormPreview.Picture1.Height = 17500
    Printer.PaperSize = 5
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    m = IIf(TypeName(mObj) = "Printer" And Not tSign, tPlus, 0)
    If Not tSign Then PaintHeader mParam(2), mObj, dParams, m
    t = dCustomer.Top + m
    seb = ""
    mNoSC = ""
    Do
        tMax = 0
        If seb <> res(iNoSJ, i) Then
            k = i
            PaintDetail dLCustomer, mObj, dLCustomer, t, tMax, tSign
            PaintDetail res(iNamaCustomer, i), mObj, dCustomer, t, tMax, tSign
            PaintDetail dLNoSJ, mObj, dLNoSJ, t, tMax, tSign
            PaintDetail res(iNoSJ, i), mObj, dNoSJ, t, tMax, tSign
            PaintDetail dLTanggalSJ, mObj, dLTanggalSJ, t, tMax, tSign
            PaintDetail cTanggal(res(iTanggal, i)), mObj, dTanggalSJ, t, tMax, tSign
            seb = res(iNoSJ, i)
            t = t + tMax * dCustomer.Height
            SumJumlahKG = 0
            SumTotal = 0
        End If
        tMax = 0
        If InStr(mNoSC, res(iNoSC, i)) = 0 Then
            mNoSC = mNoSC & "  " & res(iNoSC, i)
        End If
        PaintDetail res(iNamaBarang, i), mObj, dJenis, t, tMax, tSign
        PaintDetail cDecimal(res(iJumlah, i)), mObj, dJumlahKG, t, tMax, tSign
        PaintDetail res(iSatKecil, i), mObj, dSatKecil, t, tMax, tSign
        PaintDetail cDecimal(res(iHarga, i)), mObj, dHarga, t, tMax, tSign
        PaintDetail cDecimal(res(iTotal, i)), mObj, dTotal, t, tMax, tSign
        t = t + tMax * dJenis.Height
        If t > 15500 + m And tSign Then Exit Do
        SumJumlahKG = SumJumlahKG + res(iJumlah, i)
        SumTotal = SumTotal + res(iTotal, i)
        If i + 1 > iLast And Not tSign Then
            PrintFooter t, i, tSign
            Exit Do
        ElseIf i + 1 > UBound(res, 2) Then
            k = i + 1
            PrintFooter t, i, tSign
            FormPreview.SetTotalPage tPage
            Exit Do
        ElseIf seb <> res(iNoSJ, i + 1) Then
            PrintFooter t, i, tSign
        End If
        i = i + 1
    Loop
    PrintData = k
End Function

Private Sub PrintFooter(t As Single, ByVal i As Long, ByVal tSign As Boolean)
Dim tMax As Byte
    PaintDetail Mid(mNoSC, 3), mObj, dNoSC, t, tMax, tSign
    mNoSC = ""
    PaintDetail res(iKeterangan, i) & "", mObj, dKet, t, tMax, tSign
    If Not tSign Then
        If res(iDP, i) = 1 Then
            x1 = Round(ScaleX(dDP.Left, 1, 3))
            y1 = Round(ScaleX(t + 35, 1, 3))
            PaintBox mObj, x1 * 15, y1 * 15, (x1 + 10) * 15, (y1 + 10) * 15
            PaintLine mObj, (x1 + 2) * 15, (y1 + 5) * 15, (x1 + 5) * 15, (y1 + 8) * 15
            PaintLine mObj, (x1 + 5) * 15, (y1 + 8) * 15, (x1 + 9) * 15, (y1 + 2) * 15
            PaintDetail "DP", mObj, dDP, t, tMax
        End If
    End If
    PaintDetail cDecimal(SumJumlahKG), mObj, dSumJumlahKG, t, tMax, tSign
    PaintDetail res(iSatKecil, i), mObj, dSat2, t, tMax, tSign
    PaintDetail res(iMataUang, i), mObj, dMataUang, t, tMax, tSign
    PaintDetail cDecimal(SumTotal), mObj, dSumTotal, t, tMax, tSign
    t = t + tMax * dTotal.Height
    If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
    t = t + 300
End Sub


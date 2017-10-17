VERSION 5.00
Begin VB.Form ReportLaporanRetur 
   Caption         =   "RETUR PENJUALAN"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dNoSJ 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99999"
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
      Left            =   0
      TabIndex        =   27
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label dSat 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sat"
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
      Left            =   4620
      TabIndex        =   26
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label dSumReturKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@JUMLAHRETUR"
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
      Left            =   4200
      TabIndex        =   25
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label dSumDiscKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@JUMLAHPOT"
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
      TabIndex        =   24
      Top             =   2640
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
      Left            =   6480
      TabIndex        =   23
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label dReturKG 
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
      Left            =   5640
      TabIndex        =   22
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "RETUR"
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
      Left            =   4920
      TabIndex        =   21
      Top             =   1080
      Width           =   2295
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
      Left            =   1920
      TabIndex        =   19
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label dLNoKW 
      Caption         =   "NO KWITANSI:"
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
      Left            =   480
      TabIndex        =   20
      Top             =   1800
      Width           =   1455
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
      Left            =   9480
      TabIndex        =   18
      Top             =   120
      Width           =   1695
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
      Left            =   9600
      TabIndex        =   17
      Top             =   2640
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
      Left            =   8400
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label dPotonganKG 
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
      Left            =   7680
      TabIndex        =   15
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label dPotongan 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3.000.000,00"
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
      Left            =   8400
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label dTanggalNR 
      Caption         =   "@TGLNR"
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
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label dNoNR 
      Caption         =   "@NONR"
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
      Top             =   1560
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
      TabIndex        =   11
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label dLTanggalNR 
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
      Left            =   8400
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
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
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   11351
      Y1              =   3000
      Y2              =   3000
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
      Left            =   9600
      TabIndex        =   8
      Top             =   2160
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
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   360
      X2              =   11351
      Y1              =   1440
      Y2              =   1440
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
      Left            =   10440
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "POTONGAN"
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
      Left            =   7680
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
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
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label dLNoNR 
      Caption         =   "NO NR:"
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
      Left            =   8400
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   360
      X2              =   11351
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label dHead 
      Caption         =   "RETUR PENJUALAN"
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
      Top             =   240
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
      Left            =   480
      TabIndex        =   1
      Top             =   1560
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
      Left            =   6720
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "ReportLaporanRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim SumReturKG As Double
Dim SumDiscKG As Double
Dim SumTotal As Double
Dim mTotalPage As Integer
Dim mObj As Object
Dim iTbl As New ClassProperties
Dim mParam() As String

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    mParam = Split(tParams, "@")
    a = "select clng(left(NoSJ,5)), MataUangDetail, Nama, NoNR, TanggalNRDetail, NoKWDetail, Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as NamaBarang, ReturKG, SatKecil, Harga, DiscKG, Discount, ReturKG*Harga+DiscKG*Discount as Total, StatusNRDetail from (t_NRDetail~ left join m_Stock~ on m_Stock~.IdStock=t_NRDetail~.IdStock) left join m_Customer on m_Customer.Kode=t_NRDetail~.KodeCustomerDetail where TanggalNRDetail>=" & cD(mParam(0)) & " and TanggalNRDetail<=" & cD(mParam(1)) & " order by NoNR"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    For i = 0 To RS.Fields.Count - 1
        iTbl(RS.Fields(i).Name) = i
    Next
    res = RS.GetRows
    FormPreview.Picture1.Height = 17500
    Printer.PaperSize = 5
    mTotalPage = 0
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    m = IIf(TypeName(mObj) = "Printer" And Not tSign, tPlus, 0)
    PaintHeader mParam(0) & " - " & mParam(1), mObj, dParams, m
    t = dCustomer.Top + m
    seb = ""
    Do
        tMax = 0
        If seb <> res(iTbl("NoNR"), i) Then
            k = i
            PaintDetail dLCustomer, mObj, dLCustomer, t, tMax, tSign
            PaintDetail res(iTbl("Nama"), i), mObj, dCustomer, t, tMax, tSign
            PaintDetail dLNoNR, mObj, dLNoNR, t, tMax, tSign
            PaintDetail res(iTbl("NoNR"), i), mObj, dNoNR, t, tMax, tSign
            seb = res(iTbl("NoNR"), i)
            t = t + tMax * dCustomer.Height
            PaintDetail dLNoKW, mObj, dLNoKW, t, tMax, tSign
            PaintDetail res(iTbl("NoKWDetail"), i), mObj, dNoKW, t, tMax, tSign
            PaintDetail dLTanggalNR, mObj, dLTanggalNR, t, tMax, tSign
            PaintDetail cTanggal(res(iTbl("TanggalNRDetail"), i)), mObj, dTanggalNR, t, tMax, tSign
            t = t + tMax * dNoKW.Height
            SumReturKG = 0
            SumDiscKG = 0
            SumTotal = 0
        End If
        tMax = 0
        PaintDetail res(iTbl("NoSJ"), i), mObj, dNoSJ, t, tMax, tSign
        PaintDetail res(iTbl("NamaBarang"), i), mObj, dJenis, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("ReturKG"), i)), mObj, dReturKG, t, tMax, tSign
        PaintDetail res(iTbl("SatKecil"), i), mObj, dSat, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("Harga"), i)), mObj, dHarga, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("DiscKG"), i)), mObj, dPotonganKG, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("Discount"), i)), mObj, dPotongan, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("Total"), i)), mObj, dTotal, t, tMax, tSign
        t = t + tMax * dJenis.Height
        If t > 13500 + m And tSign Then Exit Do
        SumReturKG = SumReturKG + res(iTbl("ReturKG"), i)
        SumDiscKG = SumDiscKG + res(iTbl("DiscKG"), i)
        SumTotal = SumTotal + res(iTbl("Total"), i)
        If i + 1 > iLast And Not tSign Then
            PrintFooter t, i, tSign
            Exit Do
        ElseIf i + 1 > UBound(res, 2) Then
            k = i + 1
            mTotalPage = tPage
            FormPreview.SetTotalPage mTotalPage
            Exit Do
        ElseIf seb <> res(iTbl("NoNR"), i + 1) Then
            PrintFooter t, i, tSign
        End If
        i = i + 1
    Loop
    PrintData = k
End Function

Private Sub PrintFooter(t As Single, ByVal i As Long, ByVal tSign As Boolean)
Dim tMax As Byte
    PaintDetail cDecimal(SumReturKG), mObj, dSumReturKG, t, tMax, tSign
    PaintDetail cDecimal(SumDiscKG), mObj, dSumDiscKG, t, tMax, tSign
    PaintDetail res(iTbl("MataUangDetail"), i), mObj, dMataUang, t, tMax, tSign
    PaintDetail cDecimal(SumTotal), mObj, dSumTotal, t, tMax, tSign
    t = t + tMax * dTotal.Height
    If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
    t = t + 100
End Sub


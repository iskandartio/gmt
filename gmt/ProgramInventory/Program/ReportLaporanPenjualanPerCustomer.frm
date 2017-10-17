VERSION 5.00
Begin VB.Form ReportLaporanPenjualanPerCustomer 
   Caption         =   "LAPORAN PENJUALAN PER CUSTOMER"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dJumlah3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5820
      TabIndex        =   22
      Top             =   2340
      Width           =   2355
   End
   Begin VB.Label dTotal3 
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
      Left            =   9000
      TabIndex        =   21
      Top             =   2340
      Width           =   2355
   End
   Begin VB.Label dMataUang3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USD"
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
      Left            =   8280
      TabIndex        =   20
      Top             =   2340
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11471
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label dMataUang2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USD"
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
      TabIndex        =   19
      Top             =   1740
      Width           =   675
   End
   Begin VB.Label dTotal2 
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
      Left            =   9000
      TabIndex        =   18
      Top             =   1740
      Width           =   2355
   End
   Begin VB.Label dJumlah2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH"
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
      Left            =   5820
      TabIndex        =   17
      Top             =   1740
      Width           =   2355
   End
   Begin VB.Label dMataUang 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USD"
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
      TabIndex        =   16
      Top             =   2040
      Width           =   675
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   480
      X2              =   11471
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "JUMLAH KG"
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
      Left            =   6780
      TabIndex        =   15
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label dJumlah 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH KG"
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
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label dNamaId 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@PE RW"
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
      TabIndex        =   13
      Top             =   1440
      Width           =   3915
   End
   Begin VB.Label dHeadNoContract 
      Caption         =   "NO CONTRACT"
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
      Left            =   4620
      TabIndex        =   12
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label dHeadTanggal 
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
      Left            =   3180
      TabIndex        =   11
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label dNoContract 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NOCONTRACT"
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
      TabIndex        =   10
      Top             =   2040
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
      TabIndex        =   9
      Top             =   120
      Width           =   1695
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
      Left            =   3180
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
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
      Left            =   840
      TabIndex        =   7
      Top             =   2040
      Width           =   2235
   End
   Begin VB.Label dCustomer 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   6
      Top             =   1740
      Width           =   4455
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
      TabIndex        =   5
      Top             =   720
      Width           =   9255
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
      Left            =   9000
      TabIndex        =   4
      Top             =   2040
      Width           =   2355
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
      Index           =   3
      Left            =   10560
      TabIndex        =   3
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label dHeadNoSJ 
      Caption         =   "NO SURAT JALAN"
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
      Left            =   780
      TabIndex        =   2
      Top             =   1140
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
      Caption         =   "LAPORAN PENJUALAN PER CUSTOMER"
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
      TabIndex        =   1
      Top             =   360
      Width           =   5955
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
Attribute VB_Name = "ReportLaporanPenjualanPerCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim mParam() As String
Dim mNoSC As String

Dim iNamaId As Integer
Dim iNamaCustomer As Integer
Dim iNoSJ As Integer
Dim iTanggal As Integer
Dim iNoContract As Integer
Dim iJumlah As Integer
Dim iMataUang As Integer
Dim iTotal As Integer
Dim iSumTotal As Integer
Dim iSumJumlah As Integer
Dim res() As Variant
Dim SumTotal As Double
Dim SumJumlah As Double

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    iNamaId = 0
    iNamaCustomer = 1
    iNoSJ = 2
    iTanggal = 3
    iNoContract = 4
    iJumlah = 5
    iMataUang = 6
    iTotal = 7
    iSumTotal = 8
    iSumJumlah = 9
    mParam = Split(tParams, "@")
    If cD(mParam(0)) <> "A" Then MyFilter = " and TanggalDetail >= " & cD(mParam(0))
    If cD(mParam(1)) <> "A" Then MyFilter = MyFilter & " and TanggalDetail <= " & cD(mParam(1))
    If esc(mParam(3)) <> "" Then
        a = "select Kode from m_Customer where Nama='" & esc(mParam(3)) & "'"
        query a
        MyFilter = MyFilter & " and KodeCustomerDetail=" & RS.Fields(0).Value
    End If
    a = "SELECT QDetail.NamaId&'    '&QDetail.MataUangDetail, QDetail.Nama, QDetail.NoSJ, QDetail.TanggalDetail, QDetail.NoSC, QDetail.SumJumlah, QDetail.MataUangDetail, QDetail.SumTotal, QSumFromTop.SumTotal, QSumFromTop.SumJumlah " & _
    "from (SELECT NamaId, m_Customer.Nama, NoSJ, TanggalDetail, NoSC, sum(t_SPPDetail~.JumlahKG*KonversiKG) AS SumJumlah, t_SPPDetail~.MataUangDetail, sum(t_SPPDetail~.JumlahKG*t_SPPDetail~.Harga) AS SumTotal, t_SPPDetail~.KodeCustomerDetail, IdKodeBarang FROM (t_SPPDetail~ LEFT JOIN m_Stock~ ON m_Stock~.IdStock=t_SPPDetail~.IdStock) LEFT JOIN m_Customer ON m_Customer.Kode=t_SPPDetail~.KodeCustomerDetail where StatusDetail>1" & MyFilter & " group by MataUangDetail, m_Customer.Nama, NamaId, NoSJ,TanggalDetail,NoSC, KodeCustomerDetail, IdKodeBarang) as QDetail " & _
    " inner join (select MataUangDetail, KodeCustomerDetail, IdKodeBarang, sum(t_SPPDetail~.JumlahKG*KonversiKG) as SumJumlah, sum(t_SPPDetail~.JumlahKG*t_SPPDetail~.Harga) AS SumTotal FROM t_SPPDetail~ LEFT JOIN m_Stock~ ON t_SPPDetail~.IdStock=m_Stock~.IdStock where StatusDetail>1 and t_SPPDetail~.Harga>0 " & MyFilter & " group by MataUangDetail, IdKodeBarang, KodeCustomerDetail) as QSumFromTop" & _
    " ON (QDetail.IdKodeBarang=QSumFromTop.IdKodeBarang) AND (clng(QDetail.KodeCustomerDetail)=clng(QSumFromTop.KodeCustomerDetail)) and (QDetail.MataUangDetail=QSumFromTop.MataUangDetail)" & _
    " where QDetail.SumTotal > 0 order by QDetail.MataUangDetail, QSumFromTop.IdKodeBarang, QSumFromTop.SumTotal Desc, QDetail.NoSJ"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    res = RS.GetRows
    FormPreview.Picture1.Height = 17500
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    m = IIf(TypeName(mObj) = "Printer" And Not tSign, tPlus, 0)
    If Not tSign Then PaintHeader mParam(2), mObj, dParams, m
    If mParam(4) = "Detail" Then
        PaintHeader dHeadNoSJ, mObj, dHeadNoSJ, tPlus
        PaintHeader dHeadTanggal, mObj, dHeadTanggal, tPlus
        PaintHeader dHeadNoContract, mObj, dHeadNoContract, tPlus
    Else
        PaintHeader "NAMA CUSTOMER", mObj, dHeadNoSJ, tPlus
    End If
    t = dNamaId.Top + m
    Dim seb As String
    Dim sebCus As String
    seb = ""
    sebCus = ""
    SumJumlah = 0
    SumTotal = 0
    If i > 0 Then
        seb = res(iNamaId, i - 1)
        sebCus = res(iNamaCustomer, i - 1)
        j = i - 1
        Do While res(iNamaId, j) = seb
            SumJumlah = SumJumlah + res(iJumlah, j)
            SumTotal = SumTotal + res(iTotal, j)
            If j = 0 Then Exit Do
            j = j - 1
        Loop
    End If
    Do
        tMax = 0
        If seb <> res(iNamaId, i) Then
            SumJumlah = 0
            SumTotal = 0
            If i > 0 Then t = t + 200
            PaintDetail res(iNamaId, i), mObj, dNamaId, t, tMax, tSign
            seb = res(iNamaId, i)
            t = t + tMax * dNamaId.Height
            sebCus = ""
            If mParam(4) = "Detail" Then t = t - 200
        End If
        If sebCus <> res(iNamaCustomer, i) Then
            If mParam(4) <> "Summary" Then t = t + 200
            PaintDetail res(iNamaCustomer, i), mObj, dCustomer, t, tMax, tSign
            If mParam(4) = "Summary" Then
                PaintDetail cDecimal(res(iSumJumlah, i)), mObj, dJumlah2, t, tMax, tSign
                PaintDetail res(iMataUang, i), mObj, dMataUang2, t, tMax, tSign
                PaintDetail cDecimal(res(iSumTotal, i)), mObj, dTotal2, t, tMax, tSign
            Else
                PaintDetail cDecimal(res(iSumJumlah, i)), mObj, dJumlah2, t, tMax, tSign
                PaintDetail res(iMataUang, i), mObj, dMataUang2, t, tMax, tSign
                PaintDetail cDecimal(res(iSumTotal, i)), mObj, dTotal2, t, tMax, tSign
            End If
            sebCus = res(iNamaCustomer, i)
            'If Not tSign Then PaintLine mObj, Line1.x1, t - 50, Line1.x2, t - 50
            t = t + tMax * dCustomer.Height
            If Not tSign And mParam(4) = "Detail" Then PaintLine mObj, Line1.x1, t, Line1.x2, t
        End If
        tMax = 0
        If mParam(4) = "Detail" Then
            PaintDetail res(iNoSJ, i), mObj, dNoSJ, t, tMax, tSign
            PaintDetail cTanggal(res(iTanggal, i)), mObj, dTanggalSJ, t, tMax, tSign
            PaintDetail res(iNoContract, i), mObj, dNoContract, t, tMax, tSign
            PaintDetail cDecimal(res(iJumlah, i)), mObj, dJumlah, t, tMax, tSign
            PaintDetail res(iMataUang, i), mObj, dMataUang, t, tMax, tSign
            PaintDetail cDecimal(res(iTotal, i)), mObj, dTotal, t, tMax, tSign
            t = t + tMax * dNoSJ.Height
        End If
        SumJumlah = SumJumlah + res(iJumlah, i)
        SumTotal = SumTotal + res(iTotal, i)
        If t > 15500 + m And tSign Then Exit Do
        If i + 1 > iLast And Not tSign Then
            If i = UBound(res, 2) Then
                PrintFooter i, t, tSign
            ElseIf seb <> res(iNamaId, i + 1) Then
                PrintFooter i, t, tSign
            End If
            Exit Do
        ElseIf i + 1 > UBound(res, 2) Then
            PrintFooter i, t, tSign
            i = i + 1
            FormPreview.SetTotalPage tPage
            Exit Do
        ElseIf seb <> res(iNamaId, i + 1) Then
            PrintFooter i, t, tSign
        End If
        i = i + 1
    Loop
    PrintData = i
End Function

Private Sub PrintFooter(ByVal i As Long, tt As Single, ByVal tSign As Boolean)
Dim t As Single
Dim tMax As Byte
    t = tt
    PaintDetail cDecimal(SumJumlah), mObj, dJumlah3, t, tMax, tSign
    PaintDetail res(iMataUang, i), mObj, dMataUang3, t, tMax, tSign
    PaintDetail cDecimal(SumTotal), mObj, dTotal3, t, tMax, tSign
    SumJumlah = 0
    SumTotal = 0
End Sub



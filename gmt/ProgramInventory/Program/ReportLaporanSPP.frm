VERSION 5.00
Begin VB.Form ReportLaporanSPP 
   Caption         =   "LAPORAN SPP"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
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
      Left            =   9840
      TabIndex        =   26
      Top             =   1920
      Width           =   1455
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
      Left            =   6600
      TabIndex        =   25
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label dSatKecil2 
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
      Left            =   7560
      TabIndex        =   24
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label dCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@CURR"
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
      Left            =   8160
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label dJumlahBox 
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
      Left            =   4800
      TabIndex        =   22
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label dSatBesar 
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
      Left            =   5640
      TabIndex        =   21
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label dTanggalSJ 
      Caption         =   "@TANGGALSJ"
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
      Left            =   10080
      TabIndex        =   20
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label dLTanggalSJ 
      Caption         =   "SJ:"
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
      TabIndex        =   19
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label dSatBesar2 
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
      Left            =   5640
      TabIndex        =   18
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label dSumTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@TOTAL"
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
      Left            =   9840
      TabIndex        =   17
      Top             =   2280
      Width           =   1455
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
      Left            =   9240
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label dSumJumlahBox 
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
      Left            =   4680
      TabIndex        =   15
      Top             =   2280
      Width           =   855
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
      Left            =   7560
      TabIndex        =   14
      Top             =   1920
      Width           =   615
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
      TabIndex        =   13
      Top             =   1920
      Width           =   855
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
      Left            =   8760
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label dTanggalKirim 
      Caption         =   "@TANGGAL KIRIM"
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
      Left            =   8640
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label dNoSPP 
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
      Left            =   6240
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
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
      TabIndex        =   9
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label dLTanggalSPP 
      Caption         =   "SPP:"
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
      Left            =   8040
      TabIndex        =   8
      Top             =   1560
      Width           =   495
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
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   11351
      Y1              =   2760
      Y2              =   2760
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
      TabIndex        =   6
      Top             =   1920
      Width           =   4095
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
      TabIndex        =   5
      Top             =   1080
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
      TabIndex        =   4
      Top             =   1080
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
      Left            =   6960
      TabIndex        =   3
      Top             =   1080
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
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   360
      X2              =   11351
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label dHead 
      Caption         =   "LAPORAN SPP"
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
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5055
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
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "ReportLaporanSPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim SumJumlahBox As Double
Dim SumJumlahKG As Double
Dim SumTotal As Double
Dim mTotalPage As Integer
Dim mObj As Object
Dim iTbl As New ClassProperties
Dim params() As String

Sub LoadHeader(ByVal tParams As String, obj As Object)
    params = Split(tParams, "@")
    a = "select TanggalKirimDetail, MataUangDetail, Nama, t_SPPDetail~.NoSPP, TanggalDetail, Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as JenisBarang, t_SPPDetail~.JumlahBox, SatBesar, t_SPPDetail~.JumlahKG, SatKecil, Harga, t_SPPDetail~.JumlahKG*Harga as Total from (t_SPPDetail~ left join m_Stock~ on m_Stock~.IdStock=t_SPPDetail~.IdStock) left join m_Customer on m_Customer.Kode=t_SPPDetail~.KodeCustomerDetail where TanggalKirimDetail>=" & cD(params(0)) & " and TanggalKirimDetail<=" & cD(params(1)) & " and statusDetail>0 order by t_SPPDetail~.TanggalKirimDetail, t_SPPDetail~.NoSPP"
    a = Replace(a, "~", pTipe)
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    For i = 0 To RS.Fields.Count - 1
        iTbl(RS.Fields(i).Name) = i
    Next
    res = RS.GetRows
    Set mObj = obj
End Sub


Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
Dim tMax As Byte
Dim t As Single
    m = 0
    If Not tSign And TypeName(mObj) = "Printer" Then m = tPlus
    t = dCustomer.Top + m
    PaintHeader params(0) & " - " & params(1), mObj, dParams, m
    Do
        tMax = 0
        If seb <> res(iTbl("NoSPP"), i) Then
            k = i
            PaintDetail res(iTbl("Nama"), i), mObj, dCustomer, t, tMax, tSign
            PaintDetail res(iTbl("NoSPP"), i), mObj, dNoSPP, t, tMax, tSign
            PaintDetail dLTanggalSPP, mObj, dLTanggalSPP, t, tMax, tSign
            PaintDetail cTanggal(res(iTbl("TanggalKirimDetail"), i)), mObj, dTanggalKirim, t, tMax, tSign
            PaintDetail dLTanggalSJ, mObj, dLTanggalSJ, t, tMax, tSign
            PaintDetail cTanggal(res(iTbl("TanggalDetail"), i)), mObj, dTanggalSJ, t, tMax, tSign
            seb = res(iTbl("NoSPP"), i)
            t = t + tMax * dCustomer.Height
            SumJumlahBox = 0
            SumJumlahKG = 0
            SumTotal = 0
        End If
        tMax = 0
        PaintDetail res(iTbl("JenisBarang"), i), mObj, dJenis, t, tMax, tSign
        PaintDetail res(iTbl("JumlahBox"), i), mObj, dJumlahBox, t, tMax, tSign
        PaintDetail res(iTbl("SatBesar"), i), mObj, dSatBesar, t, tMax, tSign
        If res(iTbl("SatBesar"), i) <> res(iTbl("SatKecil"), i) Then
            PaintDetail cDecimal(res(iTbl("JumlahKG"), i)), mObj, dJumlahKG, t, tMax, tSign
            PaintDetail res(iTbl("SatKecil"), i), mObj, dSatKecil, t, tMax, tSign
        End If
        PaintDetail res(iTbl("MataUangDetail"), i), mObj, dCurr, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("Harga"), i)), mObj, dHarga, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("Total"), i)), mObj, dTotal, t, tMax, tSign
        t = t + tMax * dJenis.Height
        If t > 14500 + m And tSign Then Exit Do
        SumJumlahBox = SumJumlahBox + res(iTbl("JumlahBox"), i)
        SumJumlahKG = SumJumlahKG + res(iTbl("JumlahKG"), i)
        SumTotal = SumTotal + res(iTbl("Total"), i)
        If i + 1 > iLast And Not tSign Then
            PrintFooter t, i, tSign
            Exit Function
        ElseIf i + 1 > UBound(res, 2) Then
            k = i + 1
            PrintFooter t, i, tSign
            mTotalPage = tPage
            FormPreview.SetTotalPage mTotalPage
            Exit Do
        ElseIf seb <> res(iTbl("NoSPP"), i + 1) Then
            PrintFooter t, i, tSign
        End If
        i = i + 1
    Loop
    PrintData = k
End Function

Private Sub PrintFooter(t As Single, ByVal i As Long, ByVal tSign As Boolean)
Dim tMax As Byte
    PaintDetail SumJumlahBox, mObj, dSumJumlahBox, t, tMax, tSign
    PaintDetail res(iTbl("SatBesar"), i), mObj, dSatBesar2, t, tMax, tSign
    If res(iTbl("SatBesar"), i) <> res(iTbl("SatKecil"), i) Then
        PaintDetail cDecimal(SumJumlahKG), mObj, dSumJumlahKG, t, tMax, tSign
        PaintDetail res(iTbl("SatKecil"), i), mObj, dSatKecil2, t, tMax, tSign
    End If
    PaintDetail res(iTbl("MataUangDetail"), i), mObj, dMataUang, t, tMax, tSign
    PaintDetail cDecimal(SumTotal), mObj, dSumTotal, t, tMax, tSign
    t = t + tMax * dSumTotal.Height
    If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
    t = t + 100
End Sub



VERSION 5.00
Begin VB.Form ReportSummaryPenjualan 
   Caption         =   "SUMMARY PENJUALAN"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dGrade 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grade"
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
      Left            =   5040
      TabIndex        =   22
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label dWarna 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WARNA"
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
      Left            =   4200
      TabIndex        =   21
      Top             =   2400
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
      Left            =   9480
      TabIndex        =   20
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label dSatuan 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SATUAN"
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
      Left            =   6960
      TabIndex        =   19
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label dGrandJumlahKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7920
      TabIndex        =   18
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label dLGrandTotal 
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
      Left            =   6000
      TabIndex        =   17
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label dLSubTotal 
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
      Left            =   6000
      TabIndex        =   16
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label dJenis 
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
      Left            =   1200
      TabIndex        =   10
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label dSumJumlahKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7920
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
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
      Left            =   9480
      TabIndex        =   14
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label dJumlahKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAHKG"
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
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label dJumlah 
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
      Left            =   5760
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label dLJenis 
      Caption         =   "JENIS:"
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
      TabIndex        =   11
      Top             =   1680
      Width           =   735
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2040
      Width           =   4455
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
      Index           =   4
      Left            =   8040
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
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
      TabIndex        =   7
      Top             =   600
      Width           =   6435
   End
   Begin VB.Line dLineTotal 
      X1              =   1080
      X2              =   11475
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label dKodeBarang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE BARANG"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   3015
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
      TabIndex        =   5
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
      Left            =   5880
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label dHead 
      Caption         =   "KODE BARANG"
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
      Left            =   1200
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
      Caption         =   "SUMMARY PENJUALAN"
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
      Width           =   5055
   End
   Begin VB.Label dLMataUang 
      Caption         =   "MATA UANG:"
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
      Left            =   720
      TabIndex        =   1
      Top             =   2040
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
      Left            =   6360
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "ReportSummaryPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim SumJumlahKGJenis As Double
Dim SumJumlahKGMataUang As Double
Dim SumGrandKG As Double
Dim SumTotal As Double
Dim mTotalPage As Integer
Dim mObj As Object
Dim iTbl As New ClassProperties
Dim params() As String
Dim iCurrent As Long

Sub LoadHeader(ByVal tParams As String, obj As Object)
    params = Split(tParams, "@")
    If cD(params(0)) <> "A" Then MyFilter = " and TanggalDetail >= " & cD(params(0))
    If cD(params(1)) <> "A" Then MyFilter = MyFilter & " and TanggalDetail <= " & cD(params(1))
    If esc(params(3)) <> "" Then
        a = "select Kode from m_Customer where Nama='" & esc(params(3)) & "'"
        query a
        MyFilter = MyFilter & " and KodeCustomerDetail=" & RS.Fields(0).value
    End If
    a = "select MataUangDetail, Jenis, KodeBarang, Warna, SatBesar, Grade, sum(t_SPPDetail~.JumlahBox) as SumBox, sum(t_SPPDetail~.JumlahKG*KonversiKG) as SumKG, sum(t_SPPDetail~.JumlahKG*Harga) as Total from t_SPPDetail~ left join m_Stock~ on t_SPPDetail~.IdStock=m_Stock~.IdStock where statusDetail>1 " & MyFilter & " group by Jenis, MataUangDetail, Warna, KodeBarang, SatBesar, Grade"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    Set mObj = obj
    For i = 0 To RS.Fields.count - 1
        iTbl(RS.Fields(i).Name) = i
    Next
    res = RS.GetRows
    ReDim mPageSign(255)
    mTotalPage = 0
    SumGrandKG = 0
    For i = 0 To UBound(res, 2)
        SumGrandKG = SumGrandKG + res(iTbl("SumKG"), i)
    Next
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Byte, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
Dim tMax As Byte
Dim t As Single
    If Not tSign And TypeName(mObj) = "Printer" Then m = tPlus
    t = dMataUang.Top + m
    PaintHeader params(2), mObj, dParams, m
    seb = ""
    seb1 = ""
    Do
        If seb1 <> res(iTbl("Jenis"), i) Then
            tMax = 0
            PaintDetail dLJenis, mObj, dLJenis, t, tMax, tSign
            PaintDetail res(iTbl("Jenis"), i), mObj, dJenis, t, tMax, tSign
            seb1 = res(iTbl("Jenis"), i)
            t = t + tMax * dLJenis.Height
            SumTotal = 0
            SumJumlahKGMataUang = 0
            SumJumlahKGJenis = 0
            seb = ""
        End If
        If seb <> res(iTbl("MataUangDetail"), i) Then
            tMax = 0
            PaintDetail dLMataUang, mObj, dLMataUang, t, tMax, tSign
            PaintDetail res(iTbl("MataUangDetail"), i), mObj, dMataUang, t, tMax, tSign
            seb = res(iTbl("MataUangDetail"), i)
            t = t + tMax * dLMataUang.Height
            SumJumlahKGMataUang = 0
            SumTotal = 0
        End If
        tMax = 0
'        If res(iTbl("KodeBarang"), i) = "300D/96F" Then MsgBox "dsadsa"
        PaintDetail res(iTbl("KodeBarang"), i), mObj, dKodeBarang, t, tMax, tSign
        PaintDetail res(iTbl("Warna"), i), mObj, dWarna, t, tMax, tSign
        PaintDetail res(iTbl("Grade"), i), mObj, dGrade, t, tMax, tSign
        PaintDetail cNoCents(res(iTbl("SumBox"), i)), mObj, dJumlah, t, tMax, tSign
        PaintDetail res(iTbl("SatBesar"), i), mObj, dSatuan, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("SumKG"), i)), mObj, dJumlahKG, t, tMax, tSign
        PaintDetail cDecimal(res(iTbl("Total"), i)), mObj, dTotal, t, tMax, tSign
        t = t + tMax * dKodeBarang.Height
        If t > FormPreview.Picture1.Height - 1000 + m Then Exit Do
        SumJumlahKGJenis = SumJumlahKGJenis + res(iTbl("SumKG"), i)
        SumTotal = SumTotal + res(iTbl("Total"), i)
        SumJumlahKGMataUang = SumJumlahKGMataUang + res(iTbl("SumKG"), i)
        iCurrent = i
        k = i
        If i + 1 <= UBound(res, 2) Then
            If seb <> res(iTbl("MataUangDetail"), i + 1) Or seb1 <> res(iTbl("Jenis"), i + 1) Then
                PrintFooterMataUang t, tSign
                k = i
            End If
            If seb1 <> res(iTbl("Jenis"), i + 1) Then
                PrintFooterJenis t, tSign
                k = i
            End If
        End If
        If i + 1 > UBound(res, 2) Then
            PrintFooterMataUang t, tSign
            PrintFooterJenis t, tSign
            PrintFooterGrand t, tSign
            mTotalPage = tPage
            FormPreview.SetTotalPage mTotalPage
            k = i + 1
            Exit Do
        ElseIf Not tSign And i + 1 > iLast Then
            Exit Do
        End If
        
        i = i + 1
    Loop
    PrintData = k
End Function

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Private Sub PrintFooterJenis(t As Single, Optional ByVal tSign As Boolean = False)
Dim tMax As Byte
    PaintDetail "SUB TOTAL (" & res(iTbl("Jenis"), iCurrent) & ")", mObj, dLSubTotal, t, tMax, tSign
    PaintDetail cDecimal(SumJumlahKGJenis), mObj, dSumJumlahKG, t, tMax, tSign
    t = t + tMax * dLSubTotal.Height
End Sub

Private Sub PrintFooterMataUang(t As Single, Optional ByVal tSign As Boolean = False)
Dim tMax As Byte
    PaintDetail "SUB TOTAL (" & res(iTbl("Mata Uang"), iCurrent) & ")", mObj, dLSubTotal, t, tMax, tSign
    PaintDetail cDecimal(SumJumlahKGMataUang), mObj, dSumJumlahKG, t, tMax, tSign
    PaintDetail cDecimal(SumTotal), mObj, dSumTotal, t, tMax, tSign
    t = t + tMax * dLSubTotal.Height
End Sub

Private Sub PrintFooterGrand(t As Single, Optional ByVal tSign As Boolean = False)
Dim tMax As Byte
    PaintDetail "GRAND TOTAL", mObj, dLGrandTotal, t, tMax, tSign
    PaintDetail cDecimal(SumGrandKG), mObj, dGrandJumlahKG, t, tMax, tSign
End Sub



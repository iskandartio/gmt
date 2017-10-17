VERSION 5.00
Begin VB.Form ReportLaporanWaste 
   Caption         =   "LAPORAN WASTE"
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
      Caption         =   "BELUM LUNAS"
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
      Left            =   1560
      TabIndex        =   23
      Top             =   2100
      Width           =   1455
   End
   Begin VB.Label dLKet 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Keterangan"
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
      Left            =   1560
      TabIndex        =   22
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Label dTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,000,000"
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
      TabIndex        =   21
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label dLTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Left            =   420
      TabIndex        =   20
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Customer"
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
      Left            =   2220
      TabIndex        =   19
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Harga"
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
      Left            =   10500
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label dParams 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Param"
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
      Left            =   300
      TabIndex        =   18
      Top             =   960
      Width           =   7455
   End
   Begin VB.Label dCurr 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RP"
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
      Left            =   9900
      TabIndex        =   17
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label dHarga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,000,000"
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
      Left            =   10500
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label dNamaBarang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PE 5000 YARD 40 S/2 CLR 1902 CONES A"
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
      Left            =   4920
      TabIndex        =   15
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Curr"
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
      Left            =   9900
      TabIndex        =   14
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Waste"
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
      Left            =   420
      TabIndex        =   13
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label dSumTotal 
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
      Left            =   -1320
      TabIndex        =   12
      Top             =   2400
      Width           =   2835
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
      Left            =   7200
      TabIndex        =   11
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label dSat 
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
      Left            =   9120
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label dQTY 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10,999"
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
      Left            =   8460
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label dNamaCustomer 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NUGRAH PRATAMA LABELINDO, CV"
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
      Left            =   2220
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label dNoWaste 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00001/06"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label dTanggal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10/02/06"
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
      Left            =   1380
      TabIndex        =   6
      Top             =   1800
      Width           =   795
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   360
      X2              =   14380
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal"
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
      Left            =   1380
      TabIndex        =   5
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY"
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
      Left            =   8400
      TabIndex        =   3
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Barang"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   360
      X2              =   14380
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WASTE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Width           =   4095
   End
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Index           =   1
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "ReportLaporanWaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim SumTotal As Double
Dim mTotalPage As Integer
Dim mObj As Object
Dim mParam As String
Dim iNoWaste As Integer
Dim iTanggal As Integer
Dim iNamaCustomer As Integer
Dim iNamaBarang As Integer
Dim iQTY As Integer
Dim iSatuan As Integer
Dim iMataUang As Integer
Dim iHarga As Integer
Dim iTotal As Integer
Dim iCaraBayar As Integer
Dim iPelunasan As Integer
Dim iTanggalLunas As Integer
Dim iKet As Integer

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub SetData(y As XArrayDB, tParams As String)
    Set x = y
    iNoWaste = 0
    iTanggal = 1
    iNamaCustomer = 2
    iIdStock = 3
    iNamaBarang = 4
    iQTY = 5
    iSatuan = 6
    iMataUang = 7
    iHarga = 8
    iTotal = 9
    iCaraBayar = 10
    iPelunasan = 11
    iTanggalLunas = 12
    iKet = 13
    mParam = tParams
    FormPreview.Picture1.Height = 11000
    FormPreview.Picture1.Width = 15000
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    m = IIf(TypeName(mObj) = "Printer" And Not tSign, tPlus, 0)
    t = dNamaBarang.Top
    SumTotal = 0
    adder = dHarga.Left + 500
    PaintHeader mParam, mObj, dParams, tPlus
    PaintDetail dLTotal, mObj, dLTotal, dHead(6).Top, tMax, tSign, adder
    PaintDetail dLKet, mObj, dLKet, dHead(6).Top, tMax, tSign, adder
    Do
        tMax = 0
        PaintDetail x(i, iNoWaste), mObj, dNoWaste, t, tMax, tSign
        PaintDetail cTanggal3(x(i, iTanggal)), mObj, dTanggal, t, tMax, tSign
        PaintDetail x(i, iNamaCustomer), mObj, dNamaCustomer, t, tMax, tSign
        PaintDetail x(i, iNamaBarang), mObj, dNamaBarang, t, tMax, tSign
        PaintDetail cDecimal(x(i, iQTY)), mObj, dQTY, t, tMax, tSign
        PaintDetail x(i, iSatuan), mObj, dSat, t, tMax, tSign
        PaintDetail cDecimal(x(i, iHarga)), mObj, dHarga, t, tMax, tSign
        PaintDetail x(i, iMataUang), mObj, dCurr, t, tMax, tSign
        PaintDetail cDecimal(x(i, iTotal)), mObj, dTotal, t, tMax, tSign, adder
        PaintDetail x(i, iKet) & " " & x(i, iCaraBayar) & " " & cTanggal3(x(i, iTanggalLunas)), mObj, dKet, t, tMax, tSign, adder
        t = t + tMax * dNamaBarang.Height
        If t > 11000 + m And tSign Then Exit Do
        SumTotal = SumTotal + x(i, iTotal)
        If i = x.UpperBound(1) Then
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
    adder = dHarga.Left + 500
    PaintDetail cDecimal(SumTotal), mObj, dSumTotal, t, tMax, tSign, adder
    PaintDetail x(i, iMataUang), mObj, dCurr, t, tMax, tSign
    t = t + tMax * dSumTotal.Height
    t = t + 100
End Sub


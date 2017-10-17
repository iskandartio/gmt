VERSION 5.00
Begin VB.Form ReportLaporanOutstandingSC 
   Caption         =   "KONTRAK"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      X1              =   540
      X2              =   11460
      Y1              =   1800
      Y2              =   1800
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
      Index           =   5
      Left            =   9780
      TabIndex        =   16
      Top             =   1500
      Width           =   975
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "SISA"
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
      Left            =   8400
      TabIndex        =   15
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   7260
      TabIndex        =   14
      Top             =   1500
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
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   1500
      Width           =   5055
   End
   Begin VB.Label dSatKecil 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SATK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   2100
      Width           =   1035
   End
   Begin VB.Label dHarga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999.999.999.99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   2100
      Width           =   1515
   End
   Begin VB.Label dTanggal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "05/05/05"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label dDP 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label dSisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SISA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label dQTY 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label dNamaBarang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2100
      Width           =   5475
   End
   Begin VB.Label dNamaCustomer 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4380
      TabIndex        =   3
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label dNoSC 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99999/SC/PE/06/07"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1140
      TabIndex        =   4
      Top             =   1800
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
      Left            =   9360
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label dMataUang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Curr"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   5
      Top             =   2100
      Width           =   615
   End
   Begin VB.Label dHead 
      Caption         =   "SALES CONTRACT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3015
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
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "ReportLaporanOutstandingSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim mTotalPage As Integer
Dim mObj As Object
Dim iNoSC As Integer
Dim iTanggal As Integer
Dim iCustomer As Integer
Dim iCurr As Integer
Dim iHarga As Integer
Dim iDP As Integer
Dim iNamaBarang As Integer
Dim iQTY As Integer
Dim iSisa As Integer
Dim iSatKecil As Integer

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    a = "select NoSC, TanggalSCDetail, Nama, MataUangDetail, DPDetail, t_SCDetail~.JenisBarang+' '+t_SCDetail~.KodeBarang+' '+t_SCDetail~.Warna+' '+t_SCDetail~.NoWarna+' '+t_SCDetail~.Tube+' GRADE '+t_SCDetail~.Grade as NamaBarang, Jumlah, Jumlah-Terpakai as Sisa,Satuan, Harga, StatusDetail from t_SCDetail~ left join m_Customer on m_Customer.Kode=t_SCDetail~.KodeCustomerDetail where StatusDetail=0 and Jumlah<>0 and Terpakai<Jumlah order by NoSC"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    iNoSC = 0
    iTanggal = 1
    iCustomer = 2
    iCurr = 3
    iDP = 4
    iNamaBarang = 5
    iQTY = 6
    iSisa = 7
    iSatKecil = 8
    iHarga = 9
    res = RS.GetRows
    mTotalPage = 0
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    m = IIf(TypeName(mObj) = "Printer" And Not tSign, tPlus, 0)
    t = dNamaCustomer.Top + m
    seb = ""
    Do
        tMax = 0
        If seb <> res(iNoSC, i) Then
            k = i
            t = t + 150
            PaintDetail res(iNoSC, i), mObj, dNoSC, t, tMax, tSign
            PaintDetail cTanggal(res(iTanggal, i)), mObj, dTanggal, t, tMax, tSign
            PaintDetail res(iCustomer, i), mObj, dNamaCustomer, t, tMax, tSign
            PaintDetail IIf(res(iDP, i) = 0, "", "DP"), mObj, dDP, t, tMax, tSign
            seb = res(iNoSC, i)
            t = t + tMax * dNamaCustomer.Height
            If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
            t = t + 100
        End If
        tMax = 0
        PaintDetail res(iNamaBarang, i), mObj, dNamaBarang, t, tMax, tSign
        PaintDetail cDecimal(res(iQTY, i)), mObj, dQTY, t, tMax, tSign
        If res(iQTY, i) > 0 Then
            PaintDetail cDecimal(res(iSisa, i)), mObj, dSisa, t, tMax, tSign
        End If
        PaintDetail res(iSatKecil, i), mObj, dSatKecil, t, tMax, tSign
        PaintDetail cDecimal(res(iHarga, i)), mObj, dHarga, t, tMax, tSign
        PaintDetail res(iCurr, i), mObj, dMataUang, t, tMax, tSign
        t = t + tMax * dNamaBarang.Height
        If t > 15000 + m And tSign Then Exit Do
        If i + 1 > iLast And Not tSign Then
            Exit Do
        ElseIf i + 1 > UBound(res, 2) Then
            k = i + 1
            mTotalPage = tPage
            FormPreview.SetTotalPage mTotalPage
            Exit Do
        End If
        i = i + 1
    Loop
    PrintData = k
End Function


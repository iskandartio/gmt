VERSION 5.00
Begin VB.Form ReportLaporanPembelianHarian 
   Caption         =   "BTB HARIAN"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dDept 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dept"
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
      Left            =   3360
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   360
      X2              =   11880
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Label dHead 
      Caption         =   "Dept"
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
      Index           =   6
      Left            =   3420
      TabIndex        =   14
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label dHead 
      Caption         =   "Satuan"
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
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "QTY"
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
      Index           =   4
      Left            =   1440
      TabIndex        =   12
      Top             =   960
      Width           =   675
   End
   Begin VB.Label dSatBesar 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SatBesar"
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
      Left            =   2280
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label dMasuk 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
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
      Left            =   1380
      TabIndex        =   5
      Top             =   1920
      Width           =   735
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
      Left            =   10200
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label dNoUrut 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   675
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
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Label dSupplier 
      Caption         =   "Supplier"
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
      Left            =   900
      TabIndex        =   10
      Top             =   1620
      Width           =   7275
   End
   Begin VB.Label dLTanggal 
      Caption         =   "Tanggal:"
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
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label dHead 
      Caption         =   "Nama Barang"
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
      Index           =   3
      Left            =   5400
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label dHead 
      Caption         =   "No PO"
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
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label dTanggal 
      Caption         =   "TANGGAL"
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
      Left            =   1500
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label dHead 
      Caption         =   "BUKTI TERIMA BARANG"
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
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   4095
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "ReportLaporanPembelianHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim res() As Variant
Dim iNoPO As Integer
Dim iNamaBarang As Integer
Dim iJumlah As Integer
Dim iSat As Integer
Dim iSupplier As Integer
Dim iTanggal As Integer
Dim iDept As Integer
Dim mParams As String

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    mParam = Split(tParams, "@")
    a = "select clng(left(NoPODetail,5)), NamaBarang, QTY, Satuan, NamaSupplierDetail, TanggalBTBDetail, t_BTBDetail.Dept from t_BTBDetail left join m_StockBeli on m_StockBeli.IdStock=t_BTBDetail.IdStock where TanggalBTBDetail>=" & cD(mParam(0)) & " and TanggalBTBDetail<=" & cD(mParam(1)) & " order by TanggalBTBDetail, NamaSupplierDetail, NoPODetail"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    iNoPO = 0
    iNamaBarang = 1
    iJumlah = 2
    iSat = 3
    iSupplier = 4
    iTanggal = 5
    iDept = 6
    res = RS.GetRows
    FormPreview.Picture1.Height = 17500
    
    mTotalPage = 0
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
On Error Resume Next
Dim tMax As Byte
    PaintHeader mParams, mObj, dTanggal
    seb = ""
    seb2 = ""
    t = dTanggal.Top
    If i > 0 Then
        If res(iSupplier, i) = res(iSupplier, i - 1) Then lanjutan = "--- " Else lanjutan = ""
    End If
    Do
        If seb <> res(iTanggal, i) Then
            PaintDetail dLTanggal, mObj, dLTanggal, t, tMax, tSign
            PaintDetail cTanggal(res(iTanggal, i)), mObj, dTanggal, t, tMax, tSign
            seb = res(iTanggal, i)
            t = t + dTanggal.Height
            seb2 = ""
        End If
        If seb2 <> res(iSupplier, i) Then
            PaintDetail lanjutan & res(iSupplier, i), mObj, dSupplier, t, tMax, tSign
            lanjutan = ""
            seb2 = res(iSupplier, i)
            t = t + dSupplier.Height
        End If
        tMax = 0
        PaintDetail res(iNoPO, i), mObj, dNoUrut, t, tMax, tSign
        PaintDetail res(iNamaBarang, i), mObj, dNamaBarang, t, tMax, tSign
        PaintDetail res(iJumlah, i), mObj, dMasuk, t, tMax, tSign
        PaintDetail res(iSat, i), mObj, dSatBesar, t, tMax, tSign
        PaintDetail res(iDept, i), mObj, dDept, t, tMax, tSign
        t = t + tMax * dNamaBarang.Height
        If t > 13000 Then Exit Do
        If i = iLast And Not tSign Then
            Exit Do
        End If
        If i + 1 > UBound(res, 2) Then
            FormPreview.SetTotalPage tPage
            Exit Do
        End If
        If seb <> res(iTanggal, i + 1) Then
            t = t + 100
            If Not tSign Then
                PaintLine mObj, Lines(0).x1, t, Lines(0).x2, t
            End If
        End If
        If seb2 <> res(iSupplier, i + 1) Then
            t = t + 100
        End If
        i = i + 1
    Loop
    PrintData = i + 1
End Function


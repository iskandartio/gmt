VERSION 5.00
Begin VB.Form ReportSC 
   Caption         =   "SALES CONTRACT"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dHead 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARYATI M.SE"
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
      Index           =   14
      Left            =   9360
      TabIndex        =   35
      Top             =   6120
      Width           =   1455
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
      Left            =   7080
      TabIndex        =   34
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label dHead 
      Caption         =   "HARI"
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
      Index           =   18
      Left            =   10320
      TabIndex        =   33
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label dHead 
      Caption         =   "HARI"
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
      Index           =   17
      Left            =   10320
      TabIndex        =   32
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label dKeterangan 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@KET"
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
      TabIndex        =   31
      Top             =   5280
      Width           =   8475
   End
   Begin VB.Label dHead 
      Caption         =   "KETERANGAN"
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
      Index           =   16
      Left            =   480
      TabIndex        =   30
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label dSatuan 
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
      Left            =   6960
      TabIndex        =   29
      Top             =   2565
      Width           =   1095
   End
   Begin VB.Label dHead 
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
      Left            =   6960
      TabIndex        =   28
      Top             =   2160
      Width           =   975
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
      Left            =   8520
      TabIndex        =   27
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label dLamaKontrak 
      Caption         =   "@LAMAKONTRAK"
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
      TabIndex        =   26
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label dHead 
      Caption         =   "LAMA KONTRAK"
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
      Index           =   15
      Left            =   7200
      TabIndex        =   25
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Line Lines 
      Index           =   2
      X1              =   9300
      X2              =   10860
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label dHead 
      Alignment       =   2  'Center
      Caption         =   "DISETUJUI"
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
      Index           =   13
      Left            =   9420
      TabIndex        =   24
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label dGrandTotal 
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
      Left            =   7920
      TabIndex        =   23
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label dSumGrand 
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
      Left            =   9480
      TabIndex        =   22
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11160
      Y1              =   2880
      Y2              =   2880
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
      Left            =   9480
      TabIndex        =   21
      Top             =   2565
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
      Left            =   8040
      TabIndex        =   20
      Top             =   2565
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
      Left            =   5880
      TabIndex        =   19
      Top             =   2565
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
      Left            =   1080
      TabIndex        =   18
      Top             =   2565
      Width           =   4815
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
      TabIndex        =   17
      Top             =   2565
      Width           =   495
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   480
      X2              =   11160
      Y1              =   2520
      Y2              =   2520
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
      Left            =   5880
      TabIndex        =   16
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
      Left            =   10320
      TabIndex        =   15
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
      Left            =   1080
      TabIndex        =   14
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label dHead 
      Alignment       =   2  'Center
      Caption         =   "NO"
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
      TabIndex        =   13
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   480
      X2              =   11160
      Y1              =   2040
      Y2              =   2040
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
      Left            =   9480
      TabIndex        =   12
      Top             =   1200
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
      Left            =   9480
      TabIndex        =   11
      Top             =   1440
      Width           =   735
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
      Left            =   9480
      TabIndex        =   10
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label dNo 
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
      Left            =   9480
      TabIndex        =   9
      Top             =   720
      Width           =   1815
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
      TabIndex        =   8
      Top             =   1200
      Width           =   6855
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
      TabIndex        =   7
      Top             =   960
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
      Left            =   7200
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label dHead 
      Caption         =   "JATUH TEMPO"
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
      Left            =   7200
      TabIndex        =   5
      Top             =   1440
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
      Left            =   7200
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "SALES CONTRACT"
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
      Left            =   8520
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label dHead 
      Caption         =   "NO SALES CONTRACT"
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
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Width           =   2295
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
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3855
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
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "ReportSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim res() As Variant
Dim iTbl As New ClassProperties

Sub SetObj(obj As Object)
    Set mObj = obj
    
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    a = "select Nama, NoSC, Alamat, KetDetail, TanggalSCDetail, DPDetail as DP, MataUangDetail, WaktuPembayaranDetail, LamaKontrakDetail, JenisBarang+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as NamaBarang, Jumlah, Satuan, Harga, Jumlah*Harga as Total, NamaCustomerSCDetail, NamaMarketingDetail from t_SCDetail~ left join m_Customer on m_Customer.Kode=t_SCDetail~.KodeCustomerDetail where t_SCDetail~.NoSC='" & tParams & "'"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    For i = 0 To RS.Fields.Count - 1
        iTbl(RS.Fields(i).Name) = i
    Next
    res = RS.GetRows
    Set mObj = obj
    FormPreview.SetTotalPage -1
    FormPreview.Picture1.Height = 7000
End Sub

Sub PrintData()
On Error Resume Next
    i = 0
    PaintHeader res(iTbl("Nama"), i), mObj, dCustomer
    'PaintHeader res(iTbl("NamaMarketingDetail"), i), mObj, dNamaMarketing
    PaintHeader res(iTbl("Alamat"), i), mObj, dAlamat
    PaintHeader res(iTbl("NoSC"), i), mObj, dNo
    PaintHeader cTanggal(res(iTbl("TanggalSCDetail"), i)), mObj, dTanggal
    Dim tMataUang As String
    tMataUang = res(iTbl("MataUangDetail"), i)
    PaintHeader res(iTbl("MataUangDetail"), i), mObj, dMataUang
    PaintHeader res(iTbl("WaktuPembayaranDetail"), i), mObj, dJatuhTempo
    Dim tLamaKontrak As String
    tLamaKontrak = res(iTbl("LamaKontrakDetail"), i)
    tLamaKontrak = IIf(tLamaKontrak = 0, "", tLamaKontrak)
    PaintHeader tLamaKontrak, mObj, dLamaKontrak
    PaintHeader res(iTbl("KetDetail"), i), mObj, dKeterangan
    Dim t As Single
    Dim tMax As Byte
    Dim SumTotal As Double
    SumTotal = 0
    t = dNoUrut.Top
    For i = 0 To RS.RecordCount - 1
        tMax = 1
        PaintDetail i + 1, mObj, dNoUrut, t, tMax
        PaintDetail res(iTbl("NamaBarang"), i), mObj, dNamaBarang, t, tMax
        PaintDetail cDecimal(res(iTbl("Jumlah"), i)), mObj, dJumlah, t, tMax
        PaintDetail res(iTbl("Satuan"), i), mObj, dSatuan, t, tMax
        PaintDetail cDecimal(res(iTbl("Harga"), i)), mObj, dHarga, t, tMax
        PaintDetail cDecimal(res(iTbl("Total"), i)), mObj, dTotal, t, tMax
        SumTotal = SumTotal + res(iTbl("Total"), i)
        t = t + tMax * dNoUrut.Height
        RS.MoveNext
    Next
    PaintLine mObj, Line1.x1, t, Line1.x2, t
    t = t + 50
    If res(iTbl("DP"), 0) = 1 Then
        PaintBox mObj, dDP.Left, t + 50, dDP.Left + 150, t + 200
        PaintLine mObj, dDP.Left, t + 125, dDP.Left + 60, t + 190
        PaintLine mObj, dDP.Left + 60, t + 190, dDP.Left + 150, t + 80
        mObj.FontBold = True
        mObj.Print " DP"
    End If
    PaintDetail dGrandTotal, mObj, dGrandTotal, t, tMax
    PaintDetail cDecimal(SumTotal), mObj, dSumGrand, t, tMax
    FormPreview.fFirst.Enabled = False
    FormPreview.fLast.Enabled = False
    FormPreview.fPrev.Enabled = False
    FormPreview.fNext.Enabled = False
    FormPreview.fPage = "1"
    FormPreview.fPage.Enabled = False
    FormPreview.fPagesPrint = "1"
    FormPreview.fPagesPrint.Enabled = False
End Sub

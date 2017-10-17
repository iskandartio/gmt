VERSION 5.00
Begin VB.Form ReportNR 
   Caption         =   "NOTA RETUR"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.Line FootLines 
      Index           =   1
      X1              =   6960
      X2              =   8880
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label dHead 
      Caption         =   "KETERANGAN : "
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
      Index           =   0
      Left            =   360
      TabIndex        =   45
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label dKeterangan 
      Caption         =   "Keterangan"
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
      TabIndex        =   44
      Top             =   1800
      Width           =   9615
   End
   Begin VB.Label dKet 
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
      Left            =   360
      TabIndex        =   43
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label dKet 
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
      Left            =   360
      TabIndex        =   42
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label dKet 
      Caption         =   "MERAH     : ACCOUNTING"
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
      Left            =   360
      TabIndex        =   41
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label dKet 
      Caption         =   "KUNING    : MARKETING"
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
      Left            =   360
      TabIndex        =   40
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label dKet 
      Caption         =   "HIJAU       : FINANCE"
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
      Left            =   360
      TabIndex        =   39
      Top             =   4920
      Width           =   2055
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
      Left            =   360
      TabIndex        =   38
      Top             =   5160
      Width           =   4995
   End
   Begin VB.Label dTTD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(BINGRIANTO)"
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
      Left            =   6960
      TabIndex        =   37
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label dTangerang 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TANGERANG, 01/06/06"
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
      TabIndex        =   36
      Top             =   3840
      Width           =   2895
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
      Left            =   3840
      TabIndex        =   35
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label dTerbilang 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   360
      TabIndex        =   34
      Top             =   3480
      Width           =   10935
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
      Left            =   10080
      TabIndex        =   33
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label dSumDiscKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SUMKGDISC"
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
      TabIndex        =   32
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label dSumReturKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SUMKGRETUR"
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
      TabIndex        =   31
      Top             =   3120
      Width           =   1335
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
      Left            =   10080
      TabIndex        =   30
      Top             =   2685
      Width           =   1215
   End
   Begin VB.Label dDisc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@KGRETUR"
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
      Left            =   9180
      TabIndex        =   29
      Top             =   2685
      Width           =   915
   End
   Begin VB.Label dKgDisc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@KGDISC"
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
      TabIndex        =   28
      Top             =   2685
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
      Left            =   6960
      TabIndex        =   27
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Label dKGRetur 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@KGRETUR"
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
      Left            =   6300
      TabIndex        =   26
      Top             =   2685
      Width           =   675
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
      Left            =   2820
      TabIndex        =   25
      Top             =   2685
      Width           =   3495
   End
   Begin VB.Label dNoSJ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NO SJ"
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
      Left            =   960
      TabIndex        =   24
      Top             =   2685
      Width           =   1815
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
      Left            =   360
      TabIndex        =   23
      Top             =   2685
      Width           =   495
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   360
      X2              =   11475
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "DISC"
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
      Index           =   14
      Left            =   9360
      TabIndex        =   22
      Top             =   2400
      Width           =   735
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
      Index           =   13
      Left            =   7080
      TabIndex        =   21
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label dSatKecil2 
      Alignment       =   1  'Right Justify
      Caption         =   "@SKECIL"
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
      Left            =   8160
      TabIndex        =   20
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label dSatKecil1 
      Alignment       =   1  'Right Justify
      Caption         =   "@SAT KECIL"
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
      Left            =   5520
      TabIndex        =   19
      Top             =   2400
      Width           =   1455
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
      Left            =   10200
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label dHead 
      Caption         =   "DISC"
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
      Left            =   9240
      TabIndex        =   17
      Top             =   2190
      Width           =   735
   End
   Begin VB.Label dHead 
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
      Index           =   10
      Left            =   6720
      TabIndex        =   16
      Top             =   2190
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
      Index           =   9
      Left            =   2880
      TabIndex        =   15
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label dHead 
      Caption         =   "NO SJ"
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
      Left            =   960
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
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
      Left            =   360
      TabIndex        =   13
      Top             =   2280
      Width           =   495
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   360
      X2              =   11475
      Y1              =   2160
      Y2              =   2160
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
      Left            =   9300
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label dNoKW 
      Caption         =   "@NO KWITANSI"
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
      Left            =   9300
      TabIndex        =   11
      Top             =   1320
      Width           =   2115
   End
   Begin VB.Label dTanggalNR 
      Caption         =   "@TANGGALNR"
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
      Left            =   9300
      TabIndex        =   10
      Top             =   1080
      Width           =   2115
   End
   Begin VB.Label dNoNR 
      Caption         =   "@NO NR"
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
      Left            =   9300
      TabIndex        =   9
      Top             =   840
      Width           =   2115
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
      Left            =   360
      TabIndex        =   8
      Top             =   1320
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
      Left            =   360
      TabIndex        =   7
      Top             =   1080
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
      Left            =   7740
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
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
      Index           =   5
      Left            =   7740
      TabIndex        =   5
      Top             =   1320
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
      Left            =   7740
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "NOTA RETUR"
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
      Left            =   6240
      TabIndex        =   3
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label dHead 
      Caption         =   "NO RETUR"
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
      Left            =   7740
      TabIndex        =   2
      Top             =   840
      Width           =   1215
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label dCompanyHead 
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "ReportNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim res() As Variant
Dim mParams As String
Dim iTbl As New ClassProperties

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    Dim tFilter As String
    Dim param() As String
    param = Split(tParams, "@")
    mParams = tParams
    If InStr(tParams, "PrintNamaPenerima") > 0 Then
        tFilter = "NamaPenerima as Nama, AlamatPenerima as Alamat,"
    Else
        tFilter = "Nama, Alamat,"
    End If
    
    a = "select " & tFilter & " t_NR~.NoNR,TanggalNR, t_NR~.NoKW, t_NR~.MataUang, t_NRDetail~.NoSJ, Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as NamaBarang, SatKecil, ReturKg, t_NRDetail~.Harga, DiscKG, Discount, ReturKg*t_NRDetail~.Harga+DiscKG*Discount as Total, Ket from (((t_NRDetail~ left join t_NR~ on t_NR~.NoNR=t_NRDetail~.NoNR) left join m_Customer on m_Customer.Kode=t_NR~.KodeCustomer) left join m_Stock~ on t_NRDetail~.IdStock=m_Stock~.IdStock) left join t_SPP~ on t_SPP~.NoSJ=t_NRDetail~.NoSJ where t_NRDetail~.NoNR='" & param(0) & "'"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    For i = 0 To RS.Fields.Count - 1
        iTbl(RS.Fields(i).Name) = i
    Next
    res = RS.GetRows
    Set mObj = obj
    FormPreview.SetTotalPage -1
End Sub

Sub PrintData()
    i = 0
    If InStr(mParams, "HideCompany") > 0 Then
        PaintHeader "", mObj, dCompanyHead
    Else
        PaintHeader dCompanyHead, mObj, dCompanyHead
    End If
    PaintHeader res(iTbl("Nama"), i) & "", mObj, dCustomer
    PaintHeader res(iTbl("Alamat"), i) & "", mObj, dAlamat
    PaintHeader res(iTbl("Ket"), i), mObj, dKeterangan
    PaintHeader res(iTbl("NoNR"), i), mObj, dNoNR
    PaintHeader cTanggal(res(iTbl("TanggalNR"), i)), mObj, dTanggalNR
    PaintHeader res(iTbl("NoKW"), i), mObj, dNoKW
    Dim tMataUang As String
    tMataUang = res(iTbl("MataUang"), i)
    PaintHeader res(iTbl("MataUang"), i), mObj, dMataUang
    PaintHeader res(iTbl("SatKecil"), i), mObj, dSatKecil1
    PaintHeader res(iTbl("SatKecil"), i), mObj, dSatKecil2
    Dim t As Single
    Dim tMax As Byte
    Dim SumKGRetur As Double
    Dim SumKGDisc As Double
    Dim SumGrand As Double
    SumKGRetur = 0
    SumKGDisc = 0
    SumGrand = 0
    t = dNoUrut.Top
    For i = 0 To UBound(res, 2)
        tMax = 1
        PaintDetail i + 1, mObj, dNoUrut, t, tMax
        PaintDetail res(iTbl("NoSJ"), i), mObj, dNoSJ, t, tMax
        PaintDetail res(iTbl("NamaBarang"), i), mObj, dNamaBarang, t, tMax
        PaintDetail cDecimal(res(iTbl("ReturKg"), i)), mObj, dKGRetur, t, tMax
        PaintDetail cDecimal(res(iTbl("Harga"), i)), mObj, dHarga, t, tMax
        PaintDetail cDecimal(res(iTbl("DiscKG"), i)), mObj, dKgDisc, t, tMax
        PaintDetail cDecimal(res(iTbl("Discount"), i)), mObj, dDisc, t, tMax
        PaintDetail cDecimal(res(iTbl("Total"), i)), mObj, dTotal, t, tMax
        SumKGRetur = SumKGRetur + res(iTbl("ReturKg"), i)
        SumKGDisc = SumKGDisc + res(iTbl("DiscKG"), i)
        SumGrand = SumGrand + res(iTbl("Total"), i)
        t = t + tMax * dNoUrut.Height
    Next
    PaintLine mObj, Line1.x1, t, Line1.x2, t
    t = t + 50
    PaintDetail dGrandTotal, mObj, dGrandTotal, t, tMax
    PaintDetail cDecimal(SumKGRetur), mObj, dSumReturKG, t, tMax
    PaintDetail cDecimal(SumKGDisc), mObj, dSumDiscKG, t, tMax
    PaintDetail cDecimal(SumGrand), mObj, dSumGrand, t, tMax
    t = t + dGrandTotal.Height
    a = "select Nama from m_matauang where Kode='" & tMataUang & "'"
    query a
    PaintDetail Terbilang(Round(SumGrand, 2), RS.Fields(0).Value), mObj, dTerbilang, t, tMax
    t = t + dTerbilang.Height * 3
    t2 = t
    PaintDetail "TANGERANG, " & pServerDate, mObj, dTangerang, t, tMax
    t = t + dTTD.Top - dTangerang.Top
    
    For k = 0 To 4
        PaintDetail dKet(k), mObj, dKet(k), t2 + dKet(k).Top - dKet(0).Top, tMax
    Next
    PaintDetail "CETAK      : " & pServerDate, mObj, dTanggalCetak, t2 + dTanggalCetak.Top - dKet(0).Top, tMax
    t = t + dTTD.Top - dTangerang.Top - 1100
        
    If InStr(mParams, "HideCompany") = 0 Then
        PaintDetail dTTD, mObj, dTTD, t, tMax
        
    End If
    PaintLine mObj, FootLines(1).x1, t + 200, FootLines(1).x2, t + 200
    FormPreview.fFirst.Enabled = False
    FormPreview.fLast.Enabled = False
    FormPreview.fPrev.Enabled = False
    FormPreview.fNext.Enabled = False
    FormPreview.fPage = "1"
    FormPreview.fPage.Enabled = False
    FormPreview.fPagesPrint = "1"
    FormPreview.fPagesPrint.Enabled = False
End Sub


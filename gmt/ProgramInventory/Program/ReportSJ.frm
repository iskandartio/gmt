VERSION 5.00
Begin VB.Form ReportSJ 
   Caption         =   "SURAT JALAN"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dAngkutan 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ANGKUTAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3540
      TabIndex        =   20
      Top             =   5565
      Width           =   2895
   End
   Begin VB.Label dSopir 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SOPIR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   300
      TabIndex        =   19
      Top             =   6285
      Width           =   2775
   End
   Begin VB.Label dNoKendaraan 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KENDARAAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   300
      TabIndex        =   18
      Top             =   5565
      Width           =   2775
   End
   Begin VB.Label dSatk2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SATK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10200
      TabIndex        =   17
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label dSumKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@KG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label dSatb2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SATB"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7980
      TabIndex        =   15
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label dSumBox 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@BOX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7140
      TabIndex        =   14
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label dKet 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@KET"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6180
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label dTotal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5520
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label dBox 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label dSatb 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SATB"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7980
      TabIndex        =   10
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label dKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9.123,99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label dSatk 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@SATK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label dNoSPP 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NOSPP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   7
      Top             =   1800
      Width           =   2835
   End
   Begin VB.Label dTanggalSPP 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9060
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label dNamaBarang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA BARANG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   6915
   End
   Begin VB.Label dNoUrut 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label dTanggalSJ 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGALSJ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9060
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label dNoSJ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@NOSJ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
      Width           =   2835
   End
   Begin VB.Label dAlamat 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@ALAMAT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   1380
      Width           =   5535
   End
   Begin VB.Label dCustomer 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@CUSTOMER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "ReportSJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim SumBox As Long
Dim SumKG As Double
Dim iTbl As New ClassProperties
Dim res() As Variant

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object, Optional ByVal tSQL As String)
    If tSQL = "" Then
        a = "select NamaPenerimaDetail, AlamatPenerimaDetail, NoSJ, TanggalDetail, NoSPP, TanggalSPPDetail ,Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRD '+Grade as NamaBarang, SatBesar, t_SPPDetail~.JumlahBox, SatKecil, t_SPPDetail~.JumlahKG, KeteranganSPPDetail, NoKendaraanDetail, NamaAngkutanDetail, NamaSopirDetail from t_SPPDetail~ left join m_Stock~ on m_Stock~.IdStock=t_SPPDetail~.IdStock where t_SPPDetail~.NoSJ='" & tParams & "'"
    Else
        a = tSQL
    End If
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    res = RS.GetRows
    For i = 0 To RS.Fields.Count - 1
        iTbl(RS.Fields(i).Name) = i
    Next
    Set mObj = obj
    FormPreview.SetTotalPage -1
    FormPreview.Picture1.Height = 7000
End Sub

Sub PrintData()
On Error Resume Next
    i = 0
    PaintHeader res(iTbl("NamaPenerimaDetail"), i), mObj, dCustomer
    PaintHeader res(iTbl("AlamatPenerimaDetail"), i), mObj, dAlamat
    PaintHeader res(iTbl("NoSJ"), i), mObj, dNoSJ
    PaintHeader cTanggal(res(iTbl("TanggalDetail"), i)), mObj, dTanggalSJ
    PaintHeader res(iTbl("NoSPP"), i), mObj, dNoSPP
    PaintHeader cTanggal(res(iTbl("TanggalSPPDetail"), i)), mObj, dTanggalSPP
    PaintHeader res(iTbl("NoKendaraanDetail"), i), mObj, dNoKendaraan
    PaintHeader res(iTbl("NamaAngkutanDetail"), i), mObj, dAngkutan
    PaintHeader res(iTbl("NamaSopirDetail"), i), mObj, dSopir
    Dim t As Single
    Dim tMax As Byte
    Dim SatB As String
    Dim SatK As String
    SatB = res(iTbl("SatBesar"), i)
    SatK = res(iTbl("SatKecil"), i)
    t = dNoUrut.Top
    SumBox = 0
    SumKG = 0
    For i = 0 To RS.RecordCount - 1
        tMax = 1
        PaintDetail i + 1, mObj, dNoUrut, t, tMax
        PaintDetail res(iTbl("NamaBarang"), i), mObj, dNamaBarang, t, tMax
        PaintDetail res(iTbl("JumlahBox"), i), mObj, dBox, t, tMax
        PaintDetail res(iTbl("SatBesar"), i), mObj, dSatb, t, tMax
        If SatB <> SatK Then
            PaintDetail cDecimal(res(iTbl("JumlahKG"), i)), mObj, dKG, t, tMax
            PaintDetail res(iTbl("SatKecil"), i), mObj, dSatk, t, tMax
        End If
        PaintDetail res(iTbl("KeteranganSPPDetail"), i), mObj, dKet, t, tMax
        SumBox = SumBox + res(iTbl("JumlahBox"), i)
        SumKG = SumKG + Round(res(iTbl("JumlahKG"), i), 2)
        t = t + tMax * dNoUrut.Height
    Next
    t = t + 100
    PaintDetail dTotal, mObj, dTotal, t, tMax
    PaintDetail SumBox, mObj, dSumBox, t, tMax
    PaintDetail SatB, mObj, dSatb, t, tMax
    If SatB <> SatK Then
        PaintDetail cDecimal(SumKG), mObj, dSumKG, t, tMax
        PaintDetail SatK, mObj, dSatk, t, tMax
        PaintBox mObj, dSumBox.Left - 50, t - 50, dSatk.Left + dSatk.Width + 50, t - 50 + dSumBox.Height
    Else
        PaintBox mObj, dSumBox.Left - 50, t - 50, dSatb.Left + dSatb.Width + 50, t - 50 + dSumBox.Height
    End If
    FormPreview.fFirst.Enabled = False
    FormPreview.fLast.Enabled = False
    FormPreview.fPrev.Enabled = False
    FormPreview.fNext.Enabled = False
    FormPreview.fPage = "1"
    FormPreview.fPage.Enabled = False
    FormPreview.fPagesPrint = "1"
    FormPreview.fPagesPrint.Enabled = False
End Sub


VERSION 5.00
Begin VB.Form ReportMutasi 
   Caption         =   "MUTASI"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dLSatB 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   9600
      TabIndex        =   26
      Top             =   1710
      Width           =   765
   End
   Begin VB.Label dLSatB 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7680
      TabIndex        =   25
      Top             =   1710
      Width           =   765
   End
   Begin VB.Label dLSatB 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5760
      TabIndex        =   24
      Top             =   1710
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11460
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label dAkhir2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9570
      TabIndex        =   23
      Top             =   2640
      Width           =   1845
   End
   Begin VB.Label dKeluar2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   22
      Top             =   2640
      Width           =   1845
   End
   Begin VB.Label dMasuk2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   21
      Top             =   2640
      Width           =   1845
   End
   Begin VB.Label dLTotal 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   3360
      TabIndex        =   20
      Top             =   2640
      Width           =   2265
   End
   Begin VB.Label dHead2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   7680
      TabIndex        =   19
      Top             =   1410
      Width           =   1845
   End
   Begin VB.Label dKeluar 
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
      Height          =   330
      Left            =   7680
      TabIndex        =   18
      Top             =   2310
      Width           =   675
   End
   Begin VB.Label dNoUrut 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   17
      Top             =   2310
      Width           =   615
   End
   Begin VB.Label dNamaBarang 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   2310
      Width           =   4455
   End
   Begin VB.Label dLSatK 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   8400
      TabIndex        =   15
      Top             =   1710
      Width           =   1125
   End
   Begin VB.Label dLSatK 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   10320
      TabIndex        =   14
      Top             =   1710
      Width           =   1125
   End
   Begin VB.Label dLSatK 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   13
      Top             =   1710
      Width           =   1125
   End
   Begin VB.Label dMasuk 
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
      Height          =   330
      Left            =   5760
      TabIndex        =   12
      Top             =   2310
      Width           =   675
   End
   Begin VB.Label dMasukKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      TabIndex        =   11
      Top             =   2310
      Width           =   1155
   End
   Begin VB.Label dKeluarKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8400
      TabIndex        =   10
      Top             =   2310
      Width           =   1155
   End
   Begin VB.Label dAkhir 
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
      Height          =   330
      Left            =   9600
      TabIndex        =   9
      Top             =   2310
      Width           =   675
   End
   Begin VB.Label dAkhirKG 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10320
      TabIndex        =   8
      Top             =   2310
      Width           =   1095
   End
   Begin VB.Label dJenisBarang 
      Caption         =   "Jenis Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1950
      Width           =   4935
   End
   Begin VB.Label dHead2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Akhir"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   9600
      TabIndex        =   6
      Top             =   1410
      Width           =   1845
   End
   Begin VB.Label dHead2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   5
      Top             =   1410
      Width           =   1845
   End
   Begin VB.Label dHead2 
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
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label dHead2 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label dTanggalMutasi 
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
      Left            =   480
      TabIndex        =   2
      Top             =   990
      Width           =   4095
   End
   Begin VB.Label dHead 
      Caption         =   "MUTASI STOCK"
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
      Top             =   630
      Width           =   2895
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
      Top             =   270
      Width           =   4575
   End
End
Attribute VB_Name = "ReportMutasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim sumMasuk As Integer
Dim sumKeluar As Integer
Dim sumAkhir As Integer
Dim sumMasukKG As Double
Dim sumKeluarKG As Double
Dim sumAkhirKG As Double
Dim iCurrent As Long

Dim iIdStock As Integer
Dim iNamaBarang As Integer
Dim iTanggal As Integer
Dim iMasuk As Integer
Dim iMasukKG As Integer
Dim iKeluar As Integer
Dim iKeluarKG As Integer
Dim iAkhir As Integer
Dim iAkhirKG As Integer
Dim iNoBukti As Integer
Dim iKet As Integer
Dim iSatB As Integer
Dim iSatK As Integer
Dim iJenis As Integer
Dim iKode As Integer

Dim TopLine As Single
Dim mParams As String
Dim x As New XArrayDB

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub SetData(y As XArrayDB, tParams As String)
    iIdStock = 0
    iNamaBarang = 1
    iTanggal = 2
    iMasuk = 3
    iMasukKG = 4
    iKeluar = 5
    iKeluarKG = 6
    iAkhir = 7
    iAkhirKG = 8
    iNoBukti = 9
    iKet = 10
    iSatB = 11
    iSatK = 12
    iJenis = 13
    iKode = 14
    mParams = tParams
    Set x = y
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
On Error Resume Next
Dim tMax As Byte
Dim sebSatuan As String
Dim sebJenis As String
Dim sebKodeBarang As String
Dim NoUrut As Long
Dim v As Boolean
Dim t As Single
    sumMasuk = 0
    sumMasukKG = 0
    sumKeluar = 0
    sumKeluarKG = 0
    sumAkhir = 0
    sumAkhirKG = 0
    PaintHeader mParams, mObj, dTanggalMutasi
    t = dHead2(2).Top
    
    TopLine = dHead2(2).Top - 40
    TopLine = t - 40
    If Not tSign Then PaintLine mObj, Line1.x1, TopLine, Line1.x2, TopLine
    For j = 0 To dHead2.Count - 1
        PaintDetail dHead2(j), mObj, dHead2(j), t, tMax, tSign
    Next
    t = t + dHead2(0).Height
    For j = 0 To dLSatB.Count - 1
        PaintDetail x(i, iSatB), mObj, dLSatB(j), t, tMax, tSign
    Next
    For j = 0 To dLSatK.Count - 1
        PaintDetail x(i, iSatK), mObj, dLSatK(j), t, tMax, tSign
    Next
    PaintLine mObj, dLSatB(0).Left, t, dLSatK(2).Left + dLSatK(2).Width, t
    t = t + dHead2(0).Height
    If tPage > 1 Then
        j = i - 1
        sebSatuan = x(j, iSatB) & "@" & x(j, iSatK)
        sebJenis = x(j, iJenis)
        sebKodeBarang = x(j, iKode)
        Do While j > -1
            If sebSatuan <> x(j, iSatB) & "@" & x(j, iSatK) Then Exit Do
            If sebJenis <> x(j, iJenis) Then Exit Do
            If sebKodeBarang <> x(j, iKode) Then Exit Do
            sumMasuk = sumMasuk + x(j, iMasuk)
            sumMasukKG = sumMasukKG + x(j, iMasukKG)
            sumKeluar = sumKeluar + x(j, iKeluar)
            sumKeluarKG = sumKeluarKG + x(j, iKeluarKG)
            sumAkhir = sumAkhir + x(j, iAkhir)
            sumAkhirKG = sumAkhirKG + x(j, iAkhirKG)
            j = j - 1
        Loop
        NoUrut = i - j - 1
    End If
    Do
        If sebSatuan <> x(i, iSatB) & "@" & x(i, iSatK) Then
            sebSatuan = x(i, iSatB) & "@" & x(i, iSatK)
            If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
            sebJenis = ""
            sebKodeBarang = ""
        End If
        If sebJenis <> x(i, iJenis) Then
            t = t + 40
            PaintDetail Replace(x(i, iJenis), "@", "  "), mObj, dJenisBarang, t, tMax, tSign
            t = t + dJenisBarang.Height
            If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
            sebJenis = x(i, iJenis)
            sebKodeBarang = ""
        End If
        If sebKodeBarang <> x(i, iKode) Then
            NoUrut = 0
            sebKodeBarang = x(i, iKode)
        End If
        tMax = 0
        iCurrent = i
        NoUrut = NoUrut + 1
        PaintDetail NoUrut, mObj, dNoUrut, t, tMax, tSign
        PaintDetail x(i, iNamaBarang), mObj, dNamaBarang, t, tMax, tSign
        If x(i, iSatB) = x(i, iSatK) Then
            If x(i, iMasuk) = 0 Then a = "" Else a = x(i, iMasuk)
            PaintDetail a, mObj, dMasuk2, t, tMax, tSign
            If x(i, iKeluar) = 0 Then a = "" Else a = x(i, iKeluar)
            PaintDetail a, mObj, dKeluar2, t, tMax, tSign
            a = x(i, iAkhir)
            PaintDetail a, mObj, dAkhir2, t, tMax, tSign
        Else
            If x(i, iMasuk) = 0 Then
                PaintDetail "", mObj, dMasuk2, t, tMax, tSign
            Else
                PaintDetail x(i, iMasuk), mObj, dMasuk, t, tMax, tSign
                PaintDetail cDecimal(x(i, iMasukKG)), mObj, dMasukKG, t, tMax, tSign
            End If
            If x(i, iKeluar) = 0 Then
                PaintDetail "", mObj, dKeluar2, t, tMax, tSign
            Else
                PaintDetail x(i, iKeluar), mObj, dKeluar, t, tMax, tSign
                PaintDetail cDecimal(x(i, iKeluarKG)), mObj, dKeluarKG, t, tMax, tSign
            End If
            PaintDetail x(i, iAkhir), mObj, dAkhir, t, tMax, tSign
            PaintDetail cDecimal(x(i, iAkhirKG)), mObj, dAkhirKG, t, tMax, tSign
        End If
        If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
        sumMasuk = sumMasuk + x(i, iMasuk)
        sumMasukKG = sumMasukKG + x(i, iMasukKG)
        sumKeluar = sumKeluar + x(i, iKeluar)
        sumKeluarKG = sumKeluarKG + x(i, iKeluarKG)
        sumAkhir = sumAkhir + x(i, iAkhir)
        sumAkhirKG = sumAkhirKG + x(i, iAkhirKG)
        t = t + tMax * 250
        If i + 1 > x.UpperBound(1) Then
            FormPreview.SetTotalPage tPage
            PrintTotal t, tSign
            PrintFooter t, tSign
            Exit Do
        End If
        If t > 14000 And tSign Then Exit Do
        If sebSatuan <> x(i + 1, iSatB) & "@" & x(i + 1, iSatK) Then
            NoUrut = 0
            PrintTotal t, tSign
            PrintFooter t, tSign
            sebKodeBarang = ""
            sebJenis = ""
        ElseIf sebKodeBarang <> x(i + 1, iKode) Then
            NoUrut = 0
            PrintTotal t, tSign
        End If
        If i = iLast And Not tSign Then
            v = True
            Exit Do
        End If
        i = i + 1
    Loop
    If i < x.UpperBound(1) Then
        If x(i, iKode) = x(i + 1, iKode) Then
            v = True
        End If
    End If
    If v Then
        If Not tSign Then PaintLine mObj, Line1.x1, TopLine, Line1.x1, t
        If Not tSign Then PaintLine mObj, dNamaBarang.Left - 40, TopLine, dNamaBarang.Left - 40, t
        If Not tSign Then PaintLine mObj, dMasuk.Left, TopLine, dMasuk.Left, t
        If Not tSign Then PaintLine mObj, dMasukKG.Left, TopLine + 330, dMasukKG.Left, t
        If Not tSign Then PaintLine mObj, dKeluar.Left, TopLine, dKeluar.Left, t
        If Not tSign Then PaintLine mObj, dKeluarKG.Left, TopLine + 330, dKeluarKG.Left, t
        If Not tSign Then PaintLine mObj, dAkhir.Left, TopLine, dAkhir.Left, t
        If Not tSign Then PaintLine mObj, dAkhirKG.Left, TopLine + 330, dAkhirKG.Left, t
        If Not tSign Then PaintLine mObj, Line1.x2, TopLine, Line1.x2, t
    End If
    If i = x.UpperBound(1) Then
        If Not tSign Then PaintLine mObj, Line1.x1, TopLine, Line1.x1, t - 350
        If Not tSign Then PaintLine mObj, dNamaBarang.Left - 40, TopLine, dNamaBarang.Left - 40, t - 350
        If Not tSign Then PaintLine mObj, dMasuk.Left, TopLine, dMasuk.Left, t - 350
        If Not tSign Then PaintLine mObj, dMasukKG.Left, TopLine + 330, dMasukKG.Left, t - 350
        If Not tSign Then PaintLine mObj, dKeluar.Left, TopLine, dKeluar.Left, t - 350
        If Not tSign Then PaintLine mObj, dKeluarKG.Left, TopLine + 330, dKeluarKG.Left, t - 350
        If Not tSign Then PaintLine mObj, dAkhir.Left, TopLine, dAkhir.Left, t - 350
        If Not tSign Then PaintLine mObj, dAkhirKG.Left, TopLine + 330, dAkhirKG.Left, t - 350
        If Not tSign Then PaintLine mObj, Line1.x2, TopLine, Line1.x2, t - 350
    End If
    PrintData = i + 1
End Function

Private Sub PrintFooter(t As Single, Optional ByVal tSign As Boolean)
Dim tMax As Byte
    If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
    If Not tSign Then PaintLine mObj, Line1.x1, TopLine, Line1.x1, t
    If Not tSign Then PaintLine mObj, dNamaBarang.Left - 40, TopLine, dNamaBarang.Left - 40, t
    If Not tSign Then PaintLine mObj, dMasuk.Left, TopLine, dMasuk.Left, t
    If Not tSign Then PaintLine mObj, dKeluar.Left, TopLine, dKeluar.Left, t
    If Not tSign Then PaintLine mObj, dAkhir.Left, TopLine, dAkhir.Left, t
    If Not tSign Then PaintLine mObj, Line1.x2, TopLine, Line1.x2, t
    t = t + dAkhirKG.Height
    
End Sub

Sub PrintTotal(t As Single, tSign As Boolean)
Dim tMax As Byte
    If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
    PaintDetail dLTotal, mObj, dLTotal, t, tMax, tSign
    If x(iCurrent, iSatB) = x(iCurrent, iSatK) Then
        PaintDetail sumMasuk, mObj, dMasuk2, t, tMax, tSign
        PaintDetail sumKeluar, mObj, dKeluar2, t, tMax, tSign
        PaintDetail sumAkhir, mObj, dAkhir2, t, tMax, tSign
    Else
        PaintDetail sumMasuk, mObj, dMasuk, t, tMax, tSign
        PaintDetail cDecimal(sumMasukKG), mObj, dMasukKG, t, tMax, tSign
        PaintDetail sumKeluar, mObj, dKeluar, t, tMax, tSign
        PaintDetail cDecimal(sumKeluarKG), mObj, dKeluarKG, t, tMax, tSign
        PaintDetail sumAkhir, mObj, dAkhir, t, tMax, tSign
        PaintDetail cDecimal(sumAkhirKG), mObj, dAkhirKG, t, tMax, tSign
    End If
    t = t + dAkhirKG.Height
    sumMasuk = 0
    sumMasukKG = 0
    sumKeluar = 0
    sumKeluarKG = 0
    sumAkhir = 0
    sumAkhirKG = 0
End Sub


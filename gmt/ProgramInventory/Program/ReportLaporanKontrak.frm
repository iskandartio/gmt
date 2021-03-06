VERSION 5.00
Begin VB.Form ReportLaporanKontrak 
   Caption         =   "KONTRAK"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dLTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL KG"
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
      TabIndex        =   27
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label dSumKurang 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SISA"
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
      TabIndex        =   26
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label dSumLebih 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SISA"
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
      TabIndex        =   25
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label dSumTerkirim 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.500,00"
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
      TabIndex        =   24
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label dSumQTY 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.500,00"
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
      TabIndex        =   23
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   360
      X2              =   11700
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "KURANG"
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
      Left            =   9180
      TabIndex        =   22
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label dKurang 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SISA"
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
      TabIndex        =   21
      Top             =   1980
      Width           =   1035
   End
   Begin VB.Line Lines 
      Index           =   7
      X1              =   11700
      X2              =   11700
      Y1              =   1620
      Y2              =   1320
   End
   Begin VB.Line Lines 
      Index           =   6
      X1              =   10200
      X2              =   10200
      Y1              =   1620
      Y2              =   1320
   End
   Begin VB.Line Lines 
      Index           =   5
      X1              =   9120
      X2              =   9120
      Y1              =   1620
      Y2              =   1320
   End
   Begin VB.Line Lines 
      Index           =   4
      X1              =   8040
      X2              =   8040
      Y1              =   1620
      Y2              =   1320
   End
   Begin VB.Line Lines 
      Index           =   3
      X1              =   6960
      X2              =   6960
      Y1              =   1620
      Y2              =   1320
   End
   Begin VB.Line Lines 
      Index           =   2
      X1              =   360
      X2              =   360
      Y1              =   1620
      Y2              =   1320
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   360
      X2              =   11700
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label dTerkirim 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.500,00"
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
      TabIndex        =   20
      Top             =   1980
      Width           =   1035
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "TERKIRIM"
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
      Left            =   7020
      TabIndex        =   19
      Top             =   1380
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   11700
      Y1              =   2460
      Y2              =   2460
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
      Left            =   10500
      TabIndex        =   18
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "LEBIH"
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
      Left            =   8100
      TabIndex        =   17
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "PESANAN"
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
      Left            =   5940
      TabIndex        =   16
      Top             =   1380
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
      Left            =   780
      TabIndex        =   15
      Top             =   1380
      Width           =   4275
   End
   Begin VB.Label dSatKecil 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SATK"
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
      TabIndex        =   14
      Top             =   1980
      Width           =   795
   End
   Begin VB.Label dHarga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999.999.999.99"
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
      TabIndex        =   13
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label dTanggal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "05/05/05"
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
      Left            =   2580
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label dParams 
      Caption         =   "TANGGAL"
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
      TabIndex        =   12
      Top             =   960
      Width           =   5235
   End
   Begin VB.Label dDP 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DP"
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
      Left            =   9720
      TabIndex        =   11
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label dClosed 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*"
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
      Left            =   360
      TabIndex        =   10
      Top             =   1980
      Width           =   375
   End
   Begin VB.Label dLebih 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SISA"
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
      TabIndex        =   9
      Top             =   1980
      Width           =   1035
   End
   Begin VB.Label dQTY 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.500,00"
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
      TabIndex        =   8
      Top             =   1980
      Width           =   1035
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
      Height          =   255
      Left            =   780
      TabIndex        =   7
      Top             =   1980
      Width           =   4215
   End
   Begin VB.Label dNamaCustomer 
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
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label dNoSC 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99999/SC/PE/06/07"
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
      TabIndex        =   4
      Top             =   1680
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
      Left            =   9780
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label dMataUang 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Curr"
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
      TabIndex        =   5
      Top             =   1680
      Width           =   915
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
      Left            =   420
      TabIndex        =   1
      Top             =   600
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
      Left            =   420
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "ReportLaporanKontrak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim mTotalPage As Integer
Dim mObj As Object
Dim mParam() As String
Dim iNoSC As Integer
Dim iTanggal As Integer
Dim iCustomer As Integer
Dim iCurr As Integer
Dim iHarga As Integer
Dim iDP As Integer
Dim iNamaBarang As Integer
Dim iQTY As Integer
Dim iTerkirim As Integer
Dim iSisa As Integer
Dim iSatKecil As Integer
Dim iClosed As Integer
Dim iKonvKG As Integer
Dim sumPesan As Double
Dim sumKirim As Double
Dim sumLebih As Double
Dim sumKurang As Double

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)

    mParam = Split(tParams, "@")
    a = "select NoSC, TanggalSCDetail, Nama, MataUangDetail, DPDetail, t_SCDetail~.JenisBarang+' '+t_SCDetail~.KodeBarang+' '+t_SCDetail~.Warna+' '+t_SCDetail~.NoWarna+' '+t_SCDetail~.Tube+' GRADE '+t_SCDetail~.Grade as NamaBarang, Jumlah, Terpakai, Jumlah-Terpakai as Sisa,Satuan, Harga, StatusDetail, KonversiKG from (t_SCDetail~ left join m_Customer on m_Customer.Kode=t_SCDetail~.KodeCustomerDetail) left join (select distinct KodeBarang, SatKecil, KonversiKG from m_Stock~) as Stock on (Stock.KodeBarang=t_SCDetail~.KodeBarang and Stock.SatKecil=t_SCDetail~.Satuan) where TanggalSCDetail>=" & cD(mParam(0)) & " and TanggalSCDetail<=" & cD(mParam(1)) & " order by NoSC"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    iNoSC = 0
    iTanggal = 1
    iCustomer = 2
    iCurr = 3
    iDP = 4
    iNamaBarang = 5
    iQTY = 6
    iTerkirim = 7
    iSisa = 8
    iSatKecil = 9
    iHarga = 10
    iClosed = 11
    iKonvKG = 12
    res = RS.GetRows
    For i = 0 To UBound(res, 2)
        If IsNull(res(iKonvKG, i)) Then
            MsgBox "Warning!!!, Konversi tidak ditemukan (" & res(iNoSC, i) & ")"
            res(iKonvKG, i) = 1
        End If
        sumPesan = sumPesan + res(iQTY, i) * res(iKonvKG, i)
        sumKirim = sumKirim + res(iTerkirim, i) * res(iKonvKG, i)
        If res(iSisa, i) > 0 Then
            sumKurang = sumKurang + res(iSisa, i) * res(iKonvKG, i)
        ElseIf res(iSisa, i) < 0 Then
            sumLebih = sumLebih - res(iSisa, i) * res(iKonvKG, i)
        End If
    Next
    mTotalPage = 0
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    m = IIf(TypeName(mObj) = "Printer" And Not tSign, tPlus, 0)
    PaintHeader mParam(0) & " - " & mParam(1), mObj, dParams, m
    t = dNamaCustomer.Top + m
    seb = ""
    Do
        tMax = 0
        If seb <> res(iNoSC, i) Then
            k = i
            t = t + 150
            If Not tSign Then PaintLine mObj, Line1.x1, t, Line1.x2, t
            t = t + 50
            PaintDetail res(iNoSC, i), mObj, dNoSC, t, tMax, tSign
            PaintDetail cTanggal(res(iTanggal, i)), mObj, dTanggal, t, tMax, tSign
            PaintDetail res(iCustomer, i), mObj, dNamaCustomer, t, tMax, tSign
            PaintDetail IIf(res(iDP, i) = 0, "", "DP"), mObj, dDP, t, tMax, tSign
            PaintDetail res(iCurr, i), mObj, dMataUang, t, tMax, tSign
            seb = res(iNoSC, i)
            t = t + tMax * dNamaCustomer.Height
        End If
        tMax = 0
        PaintDetail res(iNamaBarang, i), mObj, dNamaBarang, t, tMax, tSign
        PaintDetail cDecimal(res(iQTY, i)), mObj, dQTY, t, tMax, tSign
        PaintDetail cDecimal(res(iTerkirim, i)), mObj, dTerkirim, t, tMax, tSign
        If res(iQTY, i) > 0 Then
            If Abs(res(iSisa, i) / res(iQTY, i)) > 0.25 And res(iClosed, i) = 1 Then
                If res(iSisa, i) > 0 Then
                    PaintDetail cDecimal(Abs(res(iSisa, i))) & "*", mObj, dKurang, t, tMax, tSign
                ElseIf res(iSisa, i) < 0 Then
                    PaintDetail cDecimal(Abs(res(iSisa, i))) & "*", mObj, dLebih, t, tMax, tSign
                End If
            Else
                If res(iSisa, i) > 0 Then
                    PaintDetail cDecimal(Abs(res(iSisa, i))), mObj, dKurang, t, tMax, tSign
                ElseIf res(iSisa, i) < 0 Then
                    PaintDetail cDecimal(Abs(res(iSisa, i))), mObj, dLebih, t, tMax, tSign
                End If
            End If
        End If
        PaintDetail res(iSatKecil, i), mObj, dSatKecil, t, tMax, tSign
        PaintDetail cDecimal(res(iHarga, i)), mObj, dHarga, t, tMax, tSign
        PaintDetail IIf(res(iClosed, i) = 0, "", "*"), mObj, dClosed, t, tMax, tSign
        t = t + tMax * dNamaBarang.Height
        If t > 15000 + m And tSign Then Exit Do
        If i + 1 > UBound(res, 2) Then
            k = i + 1
            mTotalPage = tPage
            FormPreview.SetTotalPage mTotalPage
            PrintFooter t
            Exit Do
        ElseIf i + 1 > iLast And Not tSign Then
            Exit Do
        End If
        i = i + 1
    Loop
    PrintData = k
End Function

Private Sub PrintFooter(ByVal t As Single)
Dim tMax As Byte
    t = t + 200
    PaintDetail dLTotal.Caption, mObj, dLTotal, t, tMax
    PaintDetail cDecimal(sumPesan), mObj, dSumQTY, t, tMax
    PaintDetail cDecimal(sumKirim), mObj, dSumTerkirim, t, tMax
    PaintDetail cDecimal(sumLebih), mObj, dSumLebih, t, tMax
    PaintDetail cDecimal(sumKurang), mObj, dSumKurang, t, tMax
    PaintBox mObj, dLTotal.Left, t - 50, dSumKurang.Left + dSumKurang.Width + 50, t + dSumKurang.Height + 50
End Sub

VERSION 5.00
Begin VB.Form ReportGiro 
   Caption         =   "GIRO"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
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
      Left            =   9240
      TabIndex        =   20
      Top             =   180
      Width           =   1695
   End
   Begin VB.Label dSumNR 
      Alignment       =   1  'Right Justify
      Caption         =   "9,999,999,999.00"
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
      TabIndex        =   19
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label dSumSJ 
      Alignment       =   1  'Right Justify
      Caption         =   "9,999,999,999.00"
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
      Left            =   6660
      TabIndex        =   18
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label dSumTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "9,999,999,999.00"
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
      Left            =   4140
      TabIndex        =   17
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label dNoSJ 
      Caption         =   "@NoSJ"
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
      Left            =   5940
      TabIndex        =   16
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label dNoNR 
      Caption         =   "@NoNR"
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
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label dNoGiro 
      Caption         =   "@NoGiro"
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
      Left            =   2940
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label dCustomer 
      Caption         =   "@Customer"
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
      TabIndex        =   13
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label dTanggal 
      Caption         =   "@Tanggal"
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
      Left            =   1500
      TabIndex        =   12
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label dNilaiNR 
      Alignment       =   1  'Right Justify
      Caption         =   "9,999,999,999.00"
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
      Left            =   9600
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "NILAI NR"
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
      Left            =   9540
      TabIndex        =   10
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label dNilaiSJ 
      Alignment       =   1  'Right Justify
      Caption         =   "9,999,999,999.00"
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
      Left            =   7020
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "NILAI SJ"
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
      Left            =   7020
      TabIndex        =   8
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label dNilai 
      Alignment       =   1  'Right Justify
      Caption         =   "9,999,999,999.00"
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
      Left            =   4500
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label dHead 
      Caption         =   "NO NR"
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
      Left            =   8460
      TabIndex        =   6
      Top             =   1140
      Width           =   1035
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
      Index           =   5
      Left            =   6000
      TabIndex        =   5
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label dHead 
      Caption         =   "TANGGAL: "
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
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1095
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
      Index           =   3
      Left            =   4500
      TabIndex        =   3
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label dHead 
      Caption         =   "NO GIRO"
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
      Left            =   2940
      TabIndex        =   2
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label dHead 
      Caption         =   "NAMA CUSTOMER"
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
      Top             =   1140
      Width           =   2595
   End
   Begin VB.Line dFooter 
      X1              =   720
      X2              =   1920
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label dHead 
      Caption         =   "TANDA TERIMA GIRO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   180
      Width           =   5055
   End
End
Attribute VB_Name = "ReportGiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim mParams() As String


Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
    dHead(0) = "TANDA TERIMA GIRO " & pTipe
    mParams = Split(tParams, "@")
    Set mObj = obj
End Sub

Function PrintData(ByVal i As String, ByVal iLast As String, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As String
Dim t1 As Single
Dim t2 As Single
Dim t3 As Single
Dim tMax1 As Byte
Dim tMax2 As Byte
Dim tMax3 As Byte
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim Last1 As String
Dim Last2 As String
Dim Last3 As String
Dim tSTT As String
Dim MyFilter As String
Dim MyFilter2 As String
Dim MyFilter3 As String
Dim tLastSTT As String
Dim SumGiro As Double
Dim SumFaktur As Double
Dim SumNR As Double
    MyFilter = ""
    If i <> "0" Then
        MyFilter = " and NoSTT > '" & i & "'"
        MyFilter2 = " and t_STT~.NoSTT>'" & i & "'"
        MyFilter3 = " and t_STTPotongan~.NoSTT>'" & i & "'"
    End If
    If iLast <> "0" Then
        MyFilter = MyFilter & " and NoSTT <= '" & iLast & "'"
        MyFilter2 = MyFilter2 & " and t_STT~.NoSTT<='" & iLast & "'"
        MyFilter3 = MyFilter3 & " and t_STTPotongan~.NoSTT<='" & iLast & "'"
    End If
    s = "select * from t_STTPembayaran" & pTipe & " where Tanggal between " & cD(mParams(0)) & " and " & cD(mParams(1)) & MyFilter & " and CaraBayar='GIRO' order by NoSTT, TanggalGiro"
    Set rs1 = Nothing
    rs1.Open s, CN, adOpenKeyset, adLockOptimistic
    s = "select t_STT~.NoSTT, NoFaktur, t_STTPelunasan~.Nilai from t_STTPelunasan~ left join t_STT~ on t_STTPelunasan~.NoSTT=t_STT~.NoSTT where t_STT~.Tanggal between " & cD(mParams(0)) & " and " & cD(mParams(1)) & MyFilter2 & " and GiroNum<>0 and NoGiro<>'' order by t_STT~.NoSTT, NoFaktur"
    Set rs2 = Nothing
    s = Replace(s, "~", pTipe)
    rs2.Open s, CN, adOpenKeyset, adLockOptimistic
    s = "select distinct t_STTPotongan~.NoSTT, left(t_STTPotongan~.NoSTT,4), t_STTPotongan~.NoBukti, t_STTPotongan~.Nilai from t_STTPotongan~ left join t_STTPembayaran~ on t_STTPembayaran~.NoSTT=t_STTPotongan~.NoSTT where t_STTPotongan~.Tanggal between " & cD(mParams(0)) & " and " & cD(mParams(1)) & MyFilter3 & " and t_STTPotongan~.Nilai>0 and CaraBayar='GIRO' order by t_STTPotongan~.NoSTT, t_STTPotongan~.NoBukti"
    Set rs3 = Nothing
    s = Replace(s, "~", pTipe)
    rs3.Open s, CN, adOpenKeyset, adLockOptimistic
    PaintHeader mParams(0) & " - " & mParams(1), mObj, dTanggal
    t1 = dCustomer.Top
    t2 = t1
    t3 = t1
    If rs1.EOF Then
        MsgBox "No Data"
        Exit Function
    End If
    Last1 = rs1!NoSTT
    Last2 = rs1!NoSTT
    Last3 = rs1!NoSTT
    tLastSTT = rs1!NoSTT
    Do Until rs1.EOF
        SumGiro = 0
        tMax1 = 0
        tMax2 = 0
        tMax3 = 0
        Do While Last1 = rs1!NoSTT
            tSTT = Last1
            If tMax1 = 0 Then PaintDetail rs1!KetCustomer & " " & rs1!NoSTT, mObj, dCustomer, t1, tMax1, tSign
            PaintDetail rs1!NamaBank & " " & rs1!NoGiro, mObj, dNoGiro, t1, tMax1, tSign
            PaintDetail Format(rs1!Nilai, "#,##0.00"), mObj, dNilai, t1, tMax1, tSign
            SumGiro = SumGiro + rs1!Nilai
            t1 = t1 + tMax1 * dCustomer.Height
            rs1.MoveNext
            If rs1.EOF Then Exit Do
            If Last1 <> rs1!NoSTT Then
                Last1 = rs1!NoSTT
                Exit Do
            End If
        Loop
        SumFaktur = 0
        Do While Last2 = rs2!NoSTT
            PaintDetail Left(rs2!NoFaktur, 5), mObj, dNoSJ, t2, tMax2, tSign
            PaintDetail Format(rs2!Nilai, "#,##0.00"), mObj, dNilaiSJ, t2, tMax2, tSign
            SumFaktur = SumFaktur + rs2!Nilai
            t2 = t2 + tMax2 * dNoSJ.Height
            rs2.MoveNext
            If rs2.EOF Then Exit Do
            If Last2 <> rs2!NoSTT Then
                Last2 = rs2!NoSTT
                Exit Do
            End If
        Loop
        Last2 = Last1
        SumNR = 0
        If Not rs3.EOF Then
            Do While Last3 = rs3!NoSTT
                PaintDetail Left(rs3!NoBukti, 5), mObj, dNoNR, t3, tMax3, tSign
                PaintDetail Format(rs3!Nilai, "#,##0.00"), mObj, dNilaiNR, t3, tMax3, tSign
                SumNR = SumNR + rs3!Nilai
                t3 = t3 + tMax3 * dNoNR.Height
                rs3.MoveNext
                If rs3.EOF Then Exit Do
                If Last3 <> rs3!NoSTT Then
                    Last3 = rs3!NoSTT
                    Exit Do
                End If
            Loop
            Last3 = Last1
        End If
        m = t1
        If t2 > m Then m = t2
        If t3 > m Then m = t3
        tMax1 = 0
        PaintDetail Format(SumGiro, "#,##0.00"), mObj, dSumTotal, m, tMax1, tSign
        PaintDetail Format(SumFaktur, "#,##0.00"), mObj, dSumSJ, m, tMax1, tSign
        If SumNR <> 0 Then PaintDetail Format(SumNR, "#,##0.00"), mObj, dSumNR, m, tMax1, tSign
        m = m + (tMax1 + 1) * dSumNR.Height
        t1 = m
        t2 = m
        t3 = m
        If m > 14000 Then Exit Do
        tLastSTT = tSTT
    Loop
    PrintData = tLastSTT
    If rs1.EOF And m <= 14000 And iLast = "0" Then FormPreview.SetTotalPage tPage
End Function


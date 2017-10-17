VERSION 5.00
Begin VB.Form ReportDP 
   Caption         =   "CHART ACCOUNT"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dNoSC 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4080
      TabIndex        =   10
      Top             =   1500
      Width           =   1395
   End
   Begin VB.Label dTanggal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "21/08/08"
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
      TabIndex        =   9
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label dCustomer 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1260
      TabIndex        =   8
      Top             =   1500
      Width           =   2775
   End
   Begin VB.Label dHead 
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
      Index           =   8
      Left            =   8280
      TabIndex        =   7
      Top             =   1080
      Width           =   3075
   End
   Begin VB.Label dHead 
      Caption         =   "Nilai"
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
      Left            =   6900
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label dHead 
      Caption         =   "Mata Uang"
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
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label dHead 
      Caption         =   "No SC"
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
      Left            =   4140
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label dHead 
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
      Index           =   4
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label dHead 
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
      Index           =   3
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   915
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   360
      X2              =   11351
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   360
      X2              =   11351
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label dHead 
      Caption         =   "DP"
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
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5055
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
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "ReportDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim mTotalPage As Integer
Dim mObj As Object
Dim iNormalDK As Integer
Dim iNoAcc As Integer
Dim iDeskripsi As Integer
Dim iChild As Integer
Dim iLevelDiatas As Integer
Dim iLevel1 As Integer
Dim iLevel2 As Integer
Dim iLevel3 As Integer
Dim iLevel4 As Integer
Dim iLevel5 As Integer
Dim iLevel6 As Integer
Dim mPageHeight As Single

Sub LoadHeader(ByVal tParams As String, obj As Object)
    iNormalDK = 0
    iNoAcc = 1
    iDeskripsi = 2
    iChild = 3
    iLevelDiatas = 4
    iLevel1 = 5
    iLevel2 = 6
    iLevel3 = 7
    iLevel4 = 8
    iLevel5 = 9
    iLevel6 = 10
    a = "select NormalDK, NoAccount, Deskripsi, Child, LevelDiatas, Level1, Level2, Level3, Level4, Level5, Level6 from m_ChartAccount order by Level1, Level2, Level3, Level4, Level5, Level6"
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    res = RS.GetRows
    Set mObj = obj
    mPageHeight = 15000
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
Dim tMax As Byte
Dim t As Single
Dim tAdder As Single
    m = 0
    If Not tSign And TypeName(mObj) = "Printer" Then m = tPlus
    t = dNoAcc.Top + m
    Do
        PaintDetail res(iNormalDK, i), mObj, dNormalDK, t, tMax, tSign, GetAdder(i)
        PaintDetail res(iNoAcc, i), mObj, dNoAcc, t, tMax, tSign, GetAdder(i)
        PaintDetail res(iDeskripsi, i), mObj, dDeskripsi, t, tMax, tSign, GetAdder(i)
        t = t + tMax * dNoAcc.Height
        If t > mPageHeight Then Exit Do
        i = i + 1
        If Not tSign And i > iLast Then Exit Do
        If tSign And i > UBound(res, 2) Then
            FormPreview.SetTotalPage tPage
            Exit Do
        End If
    Loop
    PrintData = i
End Function

Function GetAdder(i As Long) As Single
    If res(iLevel2, i) = 0 Then
        GetAdder = 0
    ElseIf res(iLevel3, i) = 0 Then
        GetAdder = 300
    ElseIf res(iLevel4, i) = 0 Then
        GetAdder = 600
    ElseIf res(iLevel5, i) = 0 Then
        GetAdder = 900
    ElseIf res(iLevel6, i) = 0 Then
        GetAdder = 1200
    Else
        GetAdder = 1500
    End If
End Function

Private Sub Label1_Click()

End Sub

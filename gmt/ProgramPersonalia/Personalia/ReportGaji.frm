VERSION 5.00
Begin VB.Form ReportRekapGaji 
   Caption         =   "LAPORAN PENJUALAN"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dDept 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dept"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label dHead 
      Caption         =   "Departemen"
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
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label dGaji 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gaji"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label dNama 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
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
      Left            =   9600
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label dParams 
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
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   7215
   End
   Begin VB.Label dNIK 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NIK"
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
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   480
      X2              =   11471
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label dHead 
      Caption         =   "NIK"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   480
      X2              =   11471
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label dHead 
      Caption         =   "LAPORAN REKAP GAJI"
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
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3375
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
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "ReportRekapGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res() As Variant
Dim mObj As Object
Dim mParam() As String
Dim iNIK As Integer
Dim iNama As Integer
Dim iGaji As Integer
Dim iDept As Integer

Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub LoadHeader(ByVal tParams As String, obj As Object)
Dim tKodeCustomer As Long
    iNIK = 0
    iNama = 1
    iGaji = 2
    iDept = 3
    mParam = Split(tParams, "@")
    If cD(mParam(0)) > 0 Then MyFilter = " and Tanggal >= " & cD(mParam(0))
    If cD(mParam(1)) > 0 Then MyFilter = MyFilter & " and Tanggal <= " & cD(mParam(1))
    a = "select t_Gaji.NIK, m_Karyawan.Nama, t_Gaji.Gaji, m_Karyawan.Departemen from t_Gaji left join m_Karyawan on t_Gaji.NIK=m_Karyawan.NIK where 1=1" & MyFilter
    query a
    If RS.RecordCount <= 0 Then Exit Sub
    res = RS.GetRows
    FormPreview.Picture1.Height = 17500
    Printer.PaperSize = 5
    Set mObj = obj
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
    Dim t As Single
    Dim tMax As Byte
    Dim k As Integer
    Dim seb As String
    t = dDept.Top
    seb = res(iDept, i)
    PaintDetail Format(res(iDept, i), "000000"), mObj, dDept, t, tMax, tSign
    t = t + dDept.Height
    Do
        k = i
        PaintDetail Format(res(iNIK, i), "000000"), mObj, dNIK, t, tMax, tSign
        PaintDetail res(iNama, i), mObj, dNama, t, tMax, tSign
        PaintDetail Format(res(iGaji, i), "#,##0"), mObj, dGaji, t, tMax, tSign
        t = t + tMax * dNIK.Height
        If t > 5500 And tSign Then Exit Do
        If i = iLast Then Exit Do
        If i + 1 > UBound(res, 2) Then
            k = i + 1
            FormPreview.SetTotalPage tPage
            Exit Do
        End If
        If res(iDept, i + 1) <> seb Then Exit Do
        i = i + 1
    Loop
    PrintData = k
End Function

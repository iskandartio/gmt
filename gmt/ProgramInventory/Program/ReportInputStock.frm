VERSION 5.00
Begin VB.Form ReportInputStock 
   Caption         =   "INPUT STOCK"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.Label dDetail 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,99"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   1920
      Width           =   9975
   End
   Begin VB.Line Lines 
      Index           =   0
      X1              =   240
      X2              =   11460
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Label dHead 
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
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label dPrintedCode 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PrintedCode"
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
      Left            =   10260
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label dHead 
      Alignment       =   1  'Right Justify
      Caption         =   "Code"
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
      Index           =   5
      Left            =   10560
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label dSatKecil 
      Caption         =   "SatKecil"
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
      Left            =   9420
      TabIndex        =   10
      Top             =   1680
      Width           =   855
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
      Left            =   7080
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label dMasukKG 
      Alignment       =   1  'Right Justify
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
      Left            =   8280
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
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
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   615
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   5055
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
      Left            =   6240
      TabIndex        =   6
      Top             =   1680
      Width           =   735
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
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label dHead 
      Caption         =   "No"
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
      Top             =   1320
      Width           =   495
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
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label dHead 
      Caption         =   "NOTA OUTPUT PACKING"
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
      Left            =   360
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "ReportInputStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObj As Object
Dim x As XArrayDB
Dim iNoNota As Integer
Dim iNamaBarang As Integer
Dim in1 As Integer
Dim iSatB As Integer
Dim in2 As Integer
Dim iSatK As Integer
Dim iPrintedCode As Integer
Dim iDetail As Integer
Dim mParams As String
Sub SetObj(obj As Object)
    Set mObj = obj
End Sub

Sub SetData(y As XArrayDB, tParams As String)
    iNoNota = 0
    iNamaBarang = 1
    in1 = 2
    iSatB = 3
    in2 = 4
    iSatK = 5
    iPrintedCode = 6
    iDetail = 7
    mParams = tParams
    Set x = y
End Sub

Function PrintData(ByVal i As Long, ByVal iLast As Long, ByVal tPage As Long, Optional ByVal tPlus As Single = 0, Optional ByVal tSign As Boolean = False) As Integer
On Error Resume Next
Dim tMax As Byte
    PaintHeader mParams, mObj, dTanggal
    t = dNoUrut.Top
    Do
        tMax = 0
        PaintDetail x(i, iNoUrut), mObj, dNoUrut, t, tMax
        PaintDetail x(i, iNamaBarang), mObj, dNamaBarang, t, tMax, tSign
        PaintDetail x(i, in1), mObj, dMasuk, t, tMax, tSign
        If x(i, iSatK) <> x(i, iSatB) Then
            PaintDetail cDecimal(x(i, in2)), mObj, dMasukKG, t, tMax, tSign
            PaintDetail x(i, iSatK), mObj, dSatKecil, t, tMax, tSign
        End If
        PaintDetail x(i, iSatB), mObj, dSatBesar, t, tMax, tSign
        PaintDetail x(i, iPrintedCode), mObj, dPrintedCode, t, tMax, tSign
        If x(i, iDetail) <> "" Then
            t = t + tMax * dNamaBarang.Height
            PaintDetail x(i, iDetail), mObj, dDetail, t, tMax, tSign
        End If
        t = t + tMax * dNamaBarang.Height
        If t > 13000 Then Exit Do
        If i = iLast And Not tSign Then
            Exit Do
        End If
        If i + 1 > x.UpperBound(1) Then
            FormPreview.SetTotalPage tPage
            Exit Do
        End If
        i = i + 1
    Loop
    PrintData = i + 1
End Function


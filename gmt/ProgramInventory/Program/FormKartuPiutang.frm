VERSION 5.00
Object = "{8AAEAB20-E970-42F3-9E69-BC54C54CC273}#4.0#0"; "usrcombo.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{BD09B73E-A5EF-4CAB-A002-921A8335B40E}#1.0#0"; "usrtruecombo.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormKartuPiutang 
   BackColor       =   &H00FFC0C0&
   Caption         =   "KARTU PIUTANG"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Tag             =   "34"
   WindowState     =   2  'Maximized
   Begin UsrTrueCombo.ITrueCombo fCustomer 
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrCombo.ICombo cbBank 
      Height          =   315
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton fCopy 
      Caption         =   "&COPY"
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   8760
      TabIndex        =   6
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      Text            =   "__/__/__"
      DataType        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
   End
   Begin VB.CheckBox fDetail 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detail"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton fPiutang 
      Caption         =   "PIUTANG"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton fGiroBelumCair 
      Caption         =   "GIRO BELUM CAIR"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7858
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama Customer"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "No"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tanggal"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "0-15 Hari"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "16-30 Hari"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "31-45 Hari"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "46-60 Hari"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "61-75 Hari"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "76-~ Hari"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Total"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4207"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4128"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1217"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1138"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1429"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1349"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2619"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2699"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2619"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2196"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2117"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2275"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2196"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2196"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2117"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=61,.parent=2,.namedParent=63"
      _StyleDefs(17)  =   "FilterBarStyle:id=64,.parent=1,.namedParent=66"
      _StyleDefs(18)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=62,.parent=61"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=65,.parent=64"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=74,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=78,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=40,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=37,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=38,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=39,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=44,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=41,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=42,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=43,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=48,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=45,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=46,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=47,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=52,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=49,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=50,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=51,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=56,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=53,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=54,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=55,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=60,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=57,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=58,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=59,.parent=15"
      _StyleDefs(70)  =   "Named:id=29:Normal"
      _StyleDefs(71)  =   ":id=29,.parent=0"
      _StyleDefs(72)  =   "Named:id=30:Heading"
      _StyleDefs(73)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=30,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=31:Footing"
      _StyleDefs(76)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=32:Selected"
      _StyleDefs(78)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=33:Caption"
      _StyleDefs(80)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(81)  =   "Named:id=34:HighlightRow"
      _StyleDefs(82)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(83)  =   "Named:id=35:EvenRow"
      _StyleDefs(84)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(85)  =   "Named:id=36:OddRow"
      _StyleDefs(86)  =   ":id=36,.parent=29"
      _StyleDefs(87)  =   "Named:id=63:RecordSelector"
      _StyleDefs(88)  =   ":id=63,.parent=30"
      _StyleDefs(89)  =   "Named:id=66:FilterBar"
      _StyleDefs(90)  =   ":id=66,.parent=29"
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Setor"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FormKartuPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim w15 As Long
Dim w30 As Long
Dim w45 As Long
Dim w60 As Long
Dim w75 As Long
Dim x As New XArrayDB
Dim mHeader As String
Dim col1 As TrueOleDBGrid80.Columns
Dim LCustomer As Boolean

Private Sub fCopy_Click()
    CopyGrid TDBGrid1
End Sub

Private Sub fCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As String, rs1() As Variant
    If Not LCustomer Then
        s = "select Nama, Kode from m_customer order by Nama"
        query s
        rs1 = RS.GetRows
        fCustomer.SetDB rs1
        fCustomer.SetType "String"
        LCustomer = True
    End If
    If KeyCode = 13 Then
        fDetail_Click
    End If
End Sub

Private Sub fDetail_Click()
    If mHeader = "PIUTANG" Then
        fPiutang_Click
    ElseIf mHeader = "GIRO BELUM CAIR" Then
        fGiroBelumCair_Click
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer, s As String
    Caption = Caption & "---" & pTipe
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    fTanggal = pServerDate
    For i = col1("0-15 Hari").ColIndex To TDBGrid1.Columns.count - 1
        col1(i).NumberFormat = "Standard"
        col1(i).Alignment = dbgRight
        col1(i).Width = 1275
    Next
    TDBGrid1.HeadingStyle.Alignment = dbgCenter
    col1(col1.count - 1).Width = 1400
    TDBGrid1.ColumnFooters = True
    s = "select distinct BankSetor from t_STTPembayaran~"
    query s
    i = 0
    Do Until RS.EOF
        cbBank.List(i) = RS.Fields(0).value & ""
        i = i + 1
        RS.MoveNext
    Loop
    fCustomer.SetHeader "NAMA CUSTOMER@*KODE"
    fCustomer.SetWidth "3000@1000"
    fCustomer.SetType "String@Integer"
    LCustomer = False
    fPiutang_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fPiutang_Click()
Dim a As String
Dim j As Integer
Dim k As Byte
Dim seb As String
Dim i As Integer
Dim MyFilter As String
    MyFilter = ""
    If fCustomer.Text <> "" Then MyFilter = MyFilter & " and a.Kode=" & fCustomer.GetData("Kode")
    w15 = cD(add_tanggal(fTanggal, -15))
    w30 = cD(add_tanggal(fTanggal, -30))
    w45 = cD(add_tanggal(fTanggal, -45))
    w60 = cD(add_tanggal(fTanggal, -60))
    w75 = cD(add_tanggal(fTanggal, -75))
    mHeader = "PIUTANG"
    col1("No").Visible = fDetail
    col1("Tanggal").Visible = fDetail
    col1("Total").Visible = Not col1("Tanggal").Visible
    x.ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x.DeleteRows 0
    a = "select Nama, NoSJ, MataUang, TanggalSJ, iif(MataUang='USD',m_Kurs.Nilai*(SumTot),SumTot) as Sisa from " & _
    "((select NoSJ, Kode, MataUang,TanggalSJ, sum(Tot) as SumTot, sum(Tipe) from (select NoSJ, Kode, MataUang,TanggalSJ, Total-Pelunasan as Tot, 1 as Tipe from " & _
    "t_SPP~ where TanggalSJ<=" & cD(fTanggal) & " and status>1 union all select NoFaktur, KodeCustomer, MataUang,TanggalFaktur, Nilai, 0 " & _
    "from t_STTPelunasan~ where TanggalPelunasan>" & cD(fTanggal) & ") b group by NoSJ, Kode, MataUang,TanggalSJ having sum(Tot*100)>100 " & _
    " and sum(Tipe)>0)a left join m_customer on a.Kode=m_customer.Kode) left join m_Kurs on m_Kurs.Tanggal=a.TanggalSJ where 1=1" & MyFilter & " order by Nama, right(NoSJ,2), NoSJ"
    Debug.Print a
    query a
    x.ReDim 0, RS.RecordCount - 1, 0, TDBGrid1.Columns.count - 1
    j = -1
    seb = ""
    k = col1("0-15 Hari").ColIndex
    For i = 0 To RS.RecordCount - 1
        If fDetail Then
            j = j + 1
            x(j, col1("No").ColIndex) = CLng(Left(RS.Fields("NoSJ").value, 5))
            x(j, col1("Tanggal").ColIndex) = cTanggal(RS.Fields("TanggalSJ").value)
            If seb <> RS.Fields("Nama").value Then
                x(j, col1("Nama Customer").ColIndex) = RS.Fields("Nama").value
                seb = x(j, col1("Nama Customer").ColIndex)
            End If
        Else
            If seb <> RS.Fields("Nama").value Then
                j = j + 1
                x(j, col1("Nama Customer").ColIndex) = RS.Fields("Nama").value
                seb = x(j, col1("Nama Customer").ColIndex)
            End If
        End If

        If RS.Fields("TanggalSJ").value >= w15 Then
            x(j, k) = x(j, k) + RS.Fields("Sisa").value
        ElseIf RS.Fields("TanggalSJ").value > w30 Then
            x(j, k + 1) = x(j, k + 1) + RS.Fields("Sisa").value
        ElseIf RS.Fields("TanggalSJ").value > w45 Then
            x(j, k + 2) = x(j, k + 2) + RS.Fields("Sisa").value
        ElseIf RS.Fields("TanggalSJ").value > w60 Then
            x(j, k + 3) = x(j, k + 3) + RS.Fields("Sisa").value
        ElseIf RS.Fields("TanggalSJ").value > w75 Then
            x(j, k + 4) = x(j, k + 4) + RS.Fields("Sisa").value
        Else
            x(j, k + 5) = x(j, k + 5) + RS.Fields("Sisa").value
        End If
        x(j, k + 6) = x(j, k + 6) + RS.Fields("Sisa").value
        RS.MoveNext
    Next
    x.ReDim 0, j, 0, TDBGrid1.Columns.count - 1
    TDBGrid1.Rebind
    HitungTotal
End Sub

Private Sub fPrint_Click()
    TDBGrid1.PrintInfo.PageHeader = mHeader & "\t Page \p of \P Page(s)"
    TDBGrid1.PrintInfo.SettingsMarginTop = 200
    TDBGrid1.PrintInfo.SettingsMarginBottom = 200
    TDBGrid1.PrintInfo.SettingsMarginLeft = 200
    TDBGrid1.PrintInfo.SettingsMarginRight = 200
    TDBGrid1.PrintInfo.PrintPreview
End Sub

Private Sub fGiroBelumCair_Click()
Dim a As String
Dim j As Integer
Dim i As Integer
Dim seb As String
Dim k As Integer
Dim MyFilter As String
    w15 = cD(add_tanggal(fTanggal, 15))
    w30 = cD(add_tanggal(fTanggal, 30))
    w45 = cD(add_tanggal(fTanggal, 45))
    w60 = cD(add_tanggal(fTanggal, 60))
    w75 = cD(add_tanggal(fTanggal, 75))
    mHeader = "GIRO BELUM CAIR"
    col1("No").Visible = fDetail
    col1("Tanggal").Visible = fDetail
    col1("Total").Visible = Not fDetail
    x.ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x.DeleteRows 0
    If cbBank.Text <> "" Then MyFilter = " and BankSetor='" & cbBank.Text & "'"
    a = "select Nama, NoGiro, TanggalCair, NilaiRP from (t_STTPembayaran~ left join t_STT~ on t_STT~.NoSTT=t_STTPembayaran~.NoSTT) left join m_customer on t_STT~.KodeCustomer=m_customer.Kode where (TanggalCair>" & cD(fTanggal) & ") and CaraBayar='GIRO'" & MyFilter & " order by Nama"
    
    query a
    x.ReDim 0, RS.RecordCount - 1, 0, TDBGrid1.Columns.count - 1
    j = -1
    seb = ""
    k = col1("0-15 Hari").ColIndex
    Dim tglGiro As Long
    For i = 0 To RS.RecordCount - 1
        tglGiro = RS.Fields("TanggalCair").value
        If fDetail Then
            j = j + 1
            If seb <> RS.Fields("Nama").value Then
                x(j, col1("Nama Customer").ColIndex) = RS.Fields("Nama").value
                seb = x(j, col1("Nama Customer").ColIndex)
            End If
            x(j, col1("No").ColIndex) = RS.Fields("NoGiro").value
            x(j, col1("Tanggal").ColIndex) = cTanggal(tglGiro)
        ElseIf seb <> RS.Fields("Nama").value Then
            j = j + 1
            x(j, col1("Nama Customer").ColIndex) = RS.Fields("Nama").value
            seb = x(j, col1("Nama Customer").ColIndex)
        End If
        If tglGiro < w15 Then
            x(j, k) = x(j, k) + RS.Fields("NilaiRP").value
        ElseIf tglGiro < w30 Then
            x(j, k + 1) = x(j, k + 1) + RS.Fields("NilaiRP").value
        ElseIf tglGiro < w45 Then
            x(j, k + 2) = x(j, k + 2) + RS.Fields("NilaiRP").value
        ElseIf tglGiro < w60 Then
            x(j, k + 3) = x(j, k + 3) + RS.Fields("NilaiRP").value
        ElseIf tglGiro < w75 Then
            x(j, k + 4) = x(j, k + 4) + RS.Fields("NilaiRP").value
        Else
            x(j, k + 5) = x(j, k + 5) + RS.Fields("NilaiRP").value
        End If
        x(j, k + 6) = x(j, k + 6) + RS.Fields("NilaiRP").value
        RS.MoveNext
    Next
    x.ReDim 0, j, 0, TDBGrid1.Columns.count - 1
    TDBGrid1.Rebind
    HitungTotal
End Sub

Private Sub HitungTotal()
On Error Resume Next
Dim a1 As Double
Dim a2 As Double
Dim a3 As Double
Dim a4 As Double
Dim a5 As Double
Dim a6 As Double
Dim a7 As Double
Dim i As Integer
Dim k As Integer
    k = col1("0-15 Hari").ColIndex
    For i = 0 To x.UpperBound(1)
        a1 = a1 + x(i, k)
        a2 = a2 + x(i, k + 1)
        a3 = a3 + x(i, k + 2)
        a4 = a4 + x(i, k + 3)
        a5 = a5 + x(i, k + 4)
        a6 = a6 + x(i, k + 5)
        a7 = a7 + x(i, k + 6)
    Next
    TDBGrid1.Columns(k).FooterText = cDecimal(a1)
    TDBGrid1.Columns(k + 1).FooterText = cDecimal(a2)
    TDBGrid1.Columns(k + 2).FooterText = cDecimal(a3)
    TDBGrid1.Columns(k + 3).FooterText = cDecimal(a4)
    TDBGrid1.Columns(k + 4).FooterText = cDecimal(a5)
    TDBGrid1.Columns(k + 5).FooterText = cDecimal(a6)
    TDBGrid1.Columns(k + 6).FooterText = cDecimal(a7)
End Sub

Private Sub fTanggal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If mHeader = "PIUTANG" Then
            fPiutang_Click
        ElseIf mHeader = "GIRO BELUM CAIR" Then
            fGiroBelumCair_Click
        End If
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

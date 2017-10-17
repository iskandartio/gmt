VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMasterStockProses 
   BackColor       =   &H00FFC0C0&   Caption         =   "STOCK PROSES"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Tag             =   "19"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fNew 
      Caption         =   "&NEW"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton fStock 
      Caption         =   "EDIT STOCK"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5953
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "updated"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "cKey"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "cKeyIdStock"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "IdStock"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Barang"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Kode Barang"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "No Warna"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Grade"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "D"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "F"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Dept"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Satuan"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Jumlah"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Terpakai"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1296"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1217"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=767"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=688"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1111"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1032"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=3942"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3863"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1773"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1693"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(5)._MinWidth=149"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=1429"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=1349"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(6)._MinWidth=54391100"
      Splits(0)._ColumnProps(31)=   "Column(7).Width=900"
      Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=820"
      Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(35)=   "Column(8).Width=741"
      Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=661"
      Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(39)=   "Column(9).Width=688"
      Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=609"
      Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(43)=   "Column(9)._MinWidth=101421200"
      Splits(0)._ColumnProps(44)=   "Column(10).Width=2461"
      Splits(0)._ColumnProps(45)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(10)._WidthInPix=2381"
      Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(48)=   "Column(10)._MinWidth=101448608"
      Splits(0)._ColumnProps(49)=   "Column(11).Width=1085"
      Splits(0)._ColumnProps(50)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(11)._WidthInPix=1005"
      Splits(0)._ColumnProps(52)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(53)=   "Column(11)._MinWidth=101430624"
      Splits(0)._ColumnProps(54)=   "Column(12).Width=1640"
      Splits(0)._ColumnProps(55)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(12)._WidthInPix=1561"
      Splits(0)._ColumnProps(57)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(58)=   "Column(12)._MinWidth=101430624"
      Splits(0)._ColumnProps(59)=   "Column(13).Width=1746"
      Splits(0)._ColumnProps(60)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(13)._WidthInPix=1667"
      Splits(0)._ColumnProps(62)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(63)=   "Column(13)._MinWidth=101430304"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=36,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=98,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=90,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=87,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=88,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=89,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=82,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=86,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
      _StyleDefs(86)  =   "Named:id=33:Normal"
      _StyleDefs(87)  =   ":id=33,.parent=0"
      _StyleDefs(88)  =   "Named:id=34:Heading"
      _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   ":id=34,.wraptext=-1"
      _StyleDefs(91)  =   "Named:id=35:Footing"
      _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=36:Selected"
      _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=37:Caption"
      _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(97)  =   "Named:id=38:HighlightRow"
      _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=39:EvenRow"
      _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(101) =   "Named:id=40:OddRow"
      _StyleDefs(102) =   ":id=40,.parent=33"
      _StyleDefs(103) =   "Named:id=41:RecordSelector"
      _StyleDefs(104) =   ":id=41,.parent=34"
      _StyleDefs(105) =   "Named:id=42:FilterBar"
      _StyleDefs(106) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Click Update untuk Update Stock Proses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FormMasterStockProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid80.Columns
Dim x As New XArrayDB
Dim iUpdated As Integer
Dim icKey As Integer
Dim icKeyIdStock As Integer
Dim iIdStock As Integer
Dim iNamaBarang As Integer
Dim iKodeBarang As Integer
Dim iNoWarna As Integer
Dim iGrade As Integer
Dim iD1 As Integer
Dim iF1 As Integer
Dim iDept As Integer
Dim iSatuan As Integer
Dim iJumlah As Integer
Dim iTerpakai As Integer


Private Sub fDelete_Click()
    b = MsgBox("Yakin Hapus?", vbYesNo)
    If b = vbNo Then Exit Sub
    a = "delete from m_StockProses where IdDet=" & col1(icKey).Value
    ExecMe a
    TDBGrid1.Delete
End Sub

Private Sub fNew_Click()
    x.AppendRows
    j = x.UpperBound(1)
    k = TDBGrid1.Bookmark
    For i = iIdStock To x.UpperBound(2)
        x(j, i) = x(k, i)
    Next
    x(j, iUpdated) = 1
    TDBGrid1.Rebind
    TDBGrid1.MoveLast
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub DoQuery()
    a = "select 0, m_StockProses.IdDet, m_StockProses.IdStock, m_StockProses.IdStock, NamaBarang, KodeBarang, NoWarna, Grade, D, F, Dept, Satuan, m_StockProses.Jumlah, JumlahTerpakai from m_StockProses left join m_StockBeli on m_StockProses.IdStock=m_StockBeli.IdStock order by NamaBarang"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub

Private Sub Form_Load()
    iUpdated = 0
    icKey = 1
    icKeyIdStock = 2
    iIdStock = 3
    iNamaBarang = 4
    iKodeBarang = 5
    iNoWarna = 6
    iGrade = 7
    iD1 = 8
    iF1 = 9
    iDept = 10
    iSatuan = 11
    iJumlah = 12
    iTerpakai = 13

    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    TDBGridSetVisible TDBGrid1, iUpdated & "@" & icKey & "@" & icKeyIdStock
    TDBGrid1.FetchRowStyle = True
    col1(iIdStock).Locked = True
    col1(iNamaBarang).Locked = True
    col1(iJumlah).Locked = True
    col1(iTerpakai).Locked = True
    col1(iJumlah).Tag = "Decimal"
    col1(iTerpakai).Tag = "Decimal"
    TDBGridLoad TDBGrid1
    TDBGrid1.AllowAddNew = True
    DoQuery
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fStock_Click()
    FormStockBeli.LoadMe Me
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, iUpdated) = 1 Then
            If x(i, col1(icKey).ColIndex) = "" Then
                a = "insert into m_StockProses(IdStock, KodeBarang, NoWarna, Grade, D, F, Dept) values(" & _
                    x(i, col1(iIdStock).ColIndex) & _
                    ",'" & esc(x(i, col1(iKodeBarang).ColIndex)) & _
                    "','" & esc(x(i, col1(iNoWarna).ColIndex)) & _
                    "','" & esc(x(i, col1(iGrade).ColIndex)) & _
                    "'," & cNum(x(i, col1(iD1).ColIndex)) & _
                    "," & cNum(x(i, col1(iF1).ColIndex)) & _
                    ",'" & esc(x(i, col1(iDept).ColIndex)) & "')"
                ExecMe a
            Else
                a = "update m_StockProses set " & _
                        "IdStock=" & x(i, iIdStock) & _
                        ",KodeBarang='" & esc(x(i, iKodeBarang)) & _
                        "',NoWarna='" & esc(x(i, iNoWarna)) & _
                        "',Grade='" & esc(x(i, iGrade)) & _
                        "',D=" & x(i, iD1) & _
                        ",F=" & x(i, iF1) & _
                        ",Dept='" & x(i, iDept) & "' where IdDet=" & x(i, icKey)
                If ExecMe(a) = 0 Then GoTo err
                a = "update m_StockBeli set TipeStock=0 where IdStock=" & x(i, icKeyIdStock)
                If ExecMe(a) = 0 Then GoTo err
            End If
            a = "update m_StockBeli set TipeStock=1 where IdStock=" & x(i, iIdStock)
            If ExecMe(a) = 0 Then GoTo err
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoQuery
    Exit Sub
err:
    CommitTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
    col1(iUpdated).Value = 1
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("updated").ColIndex) = 1 Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Sub SetOtherRowData(ByVal tIdStock As Long)
    col1("Updated").Value = 1
    a = "select IdStock, NamaBarang, Jumlah, Satuan from m_StockBeli where IdStock=" & tIdStock
    query a
    If RS.RecordCount > 0 Then
        col1("IdStock").Value = tIdStock
        col1("Nama Barang").Value = RS.Fields("NamaBarang").Value
        col1("Jumlah").Value = RS.Fields("Jumlah").Value
        col1("Satuan").Value = RS.Fields("Satuan").Value
    Else
        col1("IdStock").Value = ""
        col1("Nama Barang").Value = ""
        col1("Jumlah").Value = ""
        col1("Satuan").Value = ""
    End If
    TDBGrid1.SetFocus
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

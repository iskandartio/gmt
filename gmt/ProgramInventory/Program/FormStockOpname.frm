VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormStockOpname 
   BackColor       =   &H00FFC0C0&
   Caption         =   "STOCK OPNAME"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Tag             =   "24"
   WindowState     =   2  'Maximized
   Begin VB.CheckBox fBerubah 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Yang Berubah Saja"
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   6
      Left            =   9060
      TabIndex        =   20
      Tag             =   "SatBesar"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   6300
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fCopy 
      Caption         =   "&COPY"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Tag             =   "Warna"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Tag             =   "Jenis"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Tag             =   "KodeBarang"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   3
      Left            =   4440
      TabIndex        =   6
      Tag             =   "NoWarna"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   4
      Left            =   6000
      TabIndex        =   5
      Tag             =   "Tube"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   5
      Left            =   7560
      TabIndex        =   4
      Tag             =   "Grade"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3555
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   6271
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IdStock"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Jenis"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kode Barang"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Warna"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "No Warna"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tube"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Grade"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Sat1"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Sat2"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Jumlah"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Jumlah2"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "NoUrut"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Kgs"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "updated"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=113"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=873"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=794"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1)._MinWidth=2508080"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2011"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1931"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(2)._MinWidth=94481980"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1085"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1005"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1455"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1376"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(4)._MinWidth=-1"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1244"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1164"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=1032"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=953"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(34)=   "Column(7).Width=1270"
      Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=1191"
      Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(38)=   "Column(8).Width=1270"
      Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=1191"
      Splits(0)._ColumnProps(41)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(42)=   "Column(9).Width=1164"
      Splits(0)._ColumnProps(43)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(9)._WidthInPix=1085"
      Splits(0)._ColumnProps(45)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(46)=   "Column(10).Width=1561"
      Splits(0)._ColumnProps(47)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(10)._WidthInPix=1482"
      Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(50)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(51)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(53)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(54)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(58)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(59)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(61)=   "Column(13).Order=14"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=93,.parent=2,.namedParent=95"
      _StyleDefs(17)  =   "FilterBarStyle:id=96,.parent=1,.namedParent=98"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=94,.parent=93"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=97,.parent=96"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=3"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13,.alignment=3"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=40,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=44,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=41,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=42,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=43,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=48,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=45,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=47,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=52,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=49,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=51,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=56,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=53,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=54,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=55,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=102,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=106,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=103,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=104,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=105,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=60,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=57,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=58,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=59,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=64,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=61,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=62,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=63,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=68,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=65,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=66,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=67,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=72,.parent=11"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=69,.parent=12"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=70,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=71,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=84,.parent=11"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=81,.parent=12"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=82,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=83,.parent=15"
      _StyleDefs(86)  =   "Named:id=29:Normal"
      _StyleDefs(87)  =   ":id=29,.parent=0"
      _StyleDefs(88)  =   "Named:id=30:Heading"
      _StyleDefs(89)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   ":id=30,.wraptext=-1"
      _StyleDefs(91)  =   "Named:id=31:Footing"
      _StyleDefs(92)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=32:Selected"
      _StyleDefs(94)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=33:Caption"
      _StyleDefs(96)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(97)  =   "Named:id=34:HighlightRow"
      _StyleDefs(98)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(99)  =   "Named:id=35:EvenRow"
      _StyleDefs(100) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(101) =   "Named:id=36:OddRow"
      _StyleDefs(102) =   ":id=36,.parent=29"
      _StyleDefs(103) =   "Named:id=95:RecordSelector"
      _StyleDefs(104) =   ":id=95,.parent=30"
      _StyleDefs(105) =   "Named:id=98:FilterBar"
      _StyleDefs(106) =   ":id=98,.parent=29"
   End
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   960
      TabIndex        =   1
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
   Begin VB.Label fLabel 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   8715
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan"
      Height          =   255
      Left            =   9060
      TabIndex        =   21
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "(Setelah Transaksi)"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Warna"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No Warna"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tube"
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FormStockOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim iIdStock As Integer
Dim iJenis As Integer
Dim iKode As Integer
Dim iWarna As Integer
Dim iNoWarna As Integer
Dim iTube As Integer
Dim iGrade As Integer
Dim iSat1 As Integer
Dim iSat2 As Integer
Dim iJumlah1 As Integer
Dim iJumlah2 As Integer
Dim iKgs As Integer
Dim iNoUrut As Integer
Dim iUpdated As Integer

Private Sub fBerubah_Click()
    DoQuery
End Sub

Private Sub fCopy_Click()
    CopyGrid TDBGrid1
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub


Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    iIdStock = 0
    iJenis = 1
    iKode = 2
    iWarna = 3
    iNoWarna = 4
    iTube = 5
    iGrade = 6
    iSat1 = 7
    iSat2 = 8
    iJumlah1 = 9
    iJumlah2 = 10
    iNoUrut = 11
    iKgs = 12
    iUpdated = 13

    col1(iJumlah1).Tag = "Integer"
    col1(iJumlah2).Tag = "Decimal"
    TDBGridLoad TDBGrid1
    'TDBGridSetVisible TDBGrid1, "updated", False
    For i = 0 To iUpdated
        col1(i).Locked = True
    Next
    x.ReDim 0, 0, 0, iUpdated
    x.DeleteRows 0
    Set TDBGrid1.Array = x
    TDBGrid1.Rebind
    TDBGrid1.FetchRowStyle = True

End Sub

Private Sub DoQuery()
'On Error GoTo err
Dim MyFilter As String
Dim sql As String
Dim tgl As Long
Dim tbl As String
    If cD(fTanggal) = "A" Then
        MsgBox "Masukkan Dulu Tanggal!!!"
        GoTo err
    End If
    sql = "select max(tgl) from mutasi where tgl<=" & cD(fTanggal.Text)
    query sql
    If IsNull(RS.Fields(0).Value) Then
        sql = "select max(tgl) from mutasi_hist where tgl<=" & cD(fTanggal.Text)
        query sql
        tbl = "mutasi_hist"
        
    Else
        
        tbl = "mutasi"
    End If
    tgl = RS.Fields(0).Value
    fTanggal.Text = cTanggal(tgl)
    sql = "select a.IDStock, b.Jenis, b.KodeBarang, b.Warna, b.NoWarna, b.Tube, b.Grade, b.SatBesar, b.SatKecil, a.akhirBox, a.akhirKg, a.akhirNoUrut, a.akhirKgs,0 from " & tbl & " as a" & _
        " left join m_stock~ as b on a.IDStock=b.IDStock" & _
        " where tgl=" & tgl
    query sql
    For i = 0 To Text.count - 1
        MyFilter = MyFilter & IIf(Trim(Text(i)) = "", "", " and b." & Text(i).Tag & " like '" & Text(i) & "'")
    Next
    sql = sql & MyFilter
    query sql
    x.ReDim 0, 0, 0, iUpdated
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    TDBGrid1.MoveFirst
err:
End Sub

Private Sub Form_Resize()
On Error GoTo err
    fLabel.Width = ScaleWidth - fLabel.Left - 100
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left - 100
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 200
err:
End Sub

Private Sub fTanggal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub deleteMutasi(ByVal tgl As Long)
On Error GoTo err
    sql = "select * into mutasi" & tgl & " from mutasi"
    ExecMe sql
    Exit Sub
err:
    sql = "drop table mutasi" & tgl
    ExecMe sql
    sql = "select * into mutasi" & tgl & " from mutasi"
    ExecMe sql
End Sub
Private Sub fUpdate_Click()
On Error GoTo err
Dim sql As String
Dim tgl As Long

    TDBGrid1.Update
    tgl = cD(fTanggal.Text)
    BeginTransaction
    sql = "select max(tgl) from mutasi where tgl=" & tgl
    query sql
    If IsNull(RS.Fields(0).Value) Then
        
        sql = "delete from mutasi_hist where tgl>" & tgl
        ExecMe sql
        sql = "delete from mutasi"
        ExecMe sql
        sql = "insert into mutasi select * from mutasi_hist where tgl=" & tgl
        ExecMe sql
    End If
    deleteMutasi tgl
    For i = 0 To x.UpperBound(1)
        If x(i, iUpdated) = "1" Then
            sql = "update mutasi set akhirBox=@jumlah1, akhirKg=@jumlah2, akhirKgs='@kgs', akhirNoUrut='@noUrut' where IDStock=@IDStock"
            sql = Replace(sql, "@IDStock", x(i, iIdStock))
            sql = Replace(sql, "@kgs", x(i, iKgs))
            sql = Replace(sql, "@noUrut", x(i, iNoUrut))
            sql = Replace(sql, "@jumlah1", x(i, iJumlah1))
            sql = Replace(sql, "@jumlah2", cNum(x(i, iJumlah2)))
            ExecMe sql
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoQuery
    Exit Sub
err:
    RollBackTransaction
    MsgBox err.Description
    MsgBox "GAGAL"
End Sub


Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(Bookmark, col1("updated").ColIndex) = "1" Then RowStyle.BackColor = vbYellow
End Sub


Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
    If KeyCode = Asc("D") Then
        FormDetailKG2.LoadMe TDBGrid1.Columns(iNoUrut).Value, TDBGrid1.Columns(iKgs).Value
        If TDBGrid1.Columns(iNoUrut).Value <> FormDetailKG2.r_NoUrut Then TDBGrid1.Columns(iUpdated).Value = "1"
        TDBGrid1.Columns(iNoUrut).Value = FormDetailKG2.r_NoUrut
        If TDBGrid1.Columns(iKgs).Value <> FormDetailKG2.r_Kgs Then TDBGrid1.Columns(iUpdated).Value = "1"
        TDBGrid1.Columns(iKgs).Value = FormDetailKG2.r_Kgs
        
        TDBGrid1.Columns(iJumlah1).Value = FormDetailKG2.r_box
        TDBGrid1.Columns(iJumlah2).Value = FormDetailKG2.r_kg
    End If
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'On Error Resume Next
    fLabel = col1(iJenis).Value & " " & col1(iKode).Value & " " & col1(iWarna).Value & " " & col1(iNoWarna).Value & " " & col1(iTube).Value & " GRADE " & col1(iGrade).Value & " (" & col1(iSat1).Value & ")"
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub


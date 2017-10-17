VERSION 5.00
Object = "{8AAEAB20-E970-42F3-9E69-BC54C54CC273}#4.0#0"; "USRCOMBO.OCX"
Object = "{5B6E0E90-AB64-4D5D-AC5E-5DC35FA1D835}#1.0#0"; "USRTEXT.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form FormProsesStock 
   Caption         =   "PROSES"
   ClientHeight    =   6510
   ClientLeft      =   -285
   ClientTop       =   675
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9360
   Tag             =   "26"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fBatal 
      Caption         =   "&BATAL"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton fOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin UsrText.IText fTanggalOK 
      Height          =   270
      Left            =   1800
      TabIndex        =   15
      Top             =   3720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
   End
   Begin UsrText.IText fNo 
      Height          =   270
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
   End
   Begin MARKETING.iToolbar iToolbar1 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   873
   End
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
      Height          =   1815
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3201
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Departemen"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=33"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=7827282"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   12632256
      ValueTranslate  =   0   'False
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=232,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3836
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Departemen"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "IdStock"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nama Barang"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "QTY"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Satuan"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Keterangan"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1799"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1720"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1191"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1111"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(2)._MinWidth=-1"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=3625"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=3545"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(3)._MinWidth=74566456"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1349"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1270"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(4)._MinWidth=6488175"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(5)._MinWidth=75715824"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(6)._MinWidth=75715824"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=58,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=24,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=50,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=62,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=54,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=15"
      _StyleDefs(58)  =   "Named:id=29:Normal"
      _StyleDefs(59)  =   ":id=29,.parent=0"
      _StyleDefs(60)  =   "Named:id=30:Heading"
      _StyleDefs(61)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=30,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=31:Footing"
      _StyleDefs(64)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=32:Selected"
      _StyleDefs(66)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=33:Caption"
      _StyleDefs(68)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(69)  =   "Named:id=34:HighlightRow"
      _StyleDefs(70)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(71)  =   "Named:id=35:EvenRow"
      _StyleDefs(72)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=36:OddRow"
      _StyleDefs(74)  =   ":id=36,.parent=29"
      _StyleDefs(75)  =   "Named:id=39:RecordSelector"
      _StyleDefs(76)  =   ":id=39,.parent=30"
      _StyleDefs(77)  =   "Named:id=42:FilterBar"
      _StyleDefs(78)  =   ":id=42,.parent=29"
   End
   Begin VB.CheckBox fCekOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton fMasterStock 
      Caption         =   "S&TOCK"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3201
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
      Columns(1).Caption=   "Nama Barang"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "QTY"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Satuan"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Keterangan"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1323"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1244"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=5398"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5318"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1693"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(2)._MinWidth=74465292"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(3)._MinWidth=-611"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2170"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2090"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(4)._MinWidth=-611"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=50,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=24,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=46,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=15"
      _StyleDefs(50)  =   "Named:id=29:Normal"
      _StyleDefs(51)  =   ":id=29,.parent=0"
      _StyleDefs(52)  =   "Named:id=30:Heading"
      _StyleDefs(53)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=30,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=31:Footing"
      _StyleDefs(56)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=32:Selected"
      _StyleDefs(58)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=33:Caption"
      _StyleDefs(60)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(61)  =   "Named:id=34:HighlightRow"
      _StyleDefs(62)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(63)  =   "Named:id=35:EvenRow"
      _StyleDefs(64)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=36:OddRow"
      _StyleDefs(66)  =   ":id=36,.parent=29"
      _StyleDefs(67)  =   "Named:id=39:RecordSelector"
      _StyleDefs(68)  =   ":id=39,.parent=30"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=29"
   End
   Begin UsrCombo.ICombo fDari 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
   End
   Begin VB.Label Label5 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "No Bukti"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Input Dari Departemen"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "PROSES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FormProsesStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid70.Columns
Dim col2 As TrueOleDBGrid70.Columns
Dim x1 As New XArrayDB
Dim x2 As New XArrayDB
Dim z As New XArrayDB
Dim m_mode As String
Dim TempQTY() As Double

Sub GetResult(ByVal tNo As String)
On Error Resume Next
    m_mode = "SEE"
    GoEvent
    fOK.Visible = False
    fBatal.Visible = False
    ClearScreen
    iToolbar1.SetQuick tNo
    a = "select t_In.Dept, NoBukti, Tanggal, t_In.IdStock, NamaBarang, QTY, Satuan, Keterangan, StatusPindahStock, TanggalOK, StatusOut from t_In left join m_StockBeli on t_In.IdStock=m_StockBeli.IdStock where NoBukti='" & tNo & "'"
    query a
    If RS.RecordCount < 1 Then Exit Sub
    fDari = RS.Fields("Dept").Value
    fNo = RS.Fields("NoBukti").Value
    fTanggal = cTanggal(RS.Fields("Tanggal").Value)
    fCekOK.Value = IIf(RS.Fields("StatusOut").Value = 0, 0, 1)
    fTanggalOK = cTanggal(RS.Fields("TanggalOK").Value)
    fOK.Visible = RS.Fields("StatusPindahStock").Value = 0 And fCekOK.Value
    fBatal.Visible = Not fOK.Visible And fCekOK.Value
    iToolbar1.SetEnabled 1, RS.Fields("StatusPindahStock").Value = 0
    iToolbar1.SetEnabled 2, RS.Fields("StatusPindahStock").Value = 0
    x1.ReDim 0, RS.RecordCount - 1, 0, TDBGrid1.Columns.Count - 1
    For i = 0 To x1.UpperBound(1)
        x1(i, 0) = RS.Fields("IdStock").Value
        x1(i, 1) = RS.Fields("NamaBarang").Value
        x1(i, 2) = RS.Fields("QTY").Value
        x1(i, 3) = RS.Fields("Satuan").Value
        x1(i, 4) = RS.Fields("Keterangan").Value
        RS.MoveNext
    Next
    TDBGrid1.Rebind
    a = "select KdDept, t_Out.Dept, t_Out.IdStock, NamaBarang, QTY, Satuan, Keterangan, StatusOK, StatusPindahStock, TanggalOK from (t_Out left join m_departemen on t_Out.Dept=m_departemen.Departemen) left join m_stockBeli on m_stockBeli.IdStock=t_Out.IdStock where NoBukti='" & tNo & "'"
    query a
    If RS.RecordCount < 1 Then Exit Sub
    fOK.Visible = False
    fBatal.Visible = False
    If RS.Fields("StatusOK").Value = 0 Then
        fCekOK.Value = 0
    Else
        fCekOK.Value = 1
        fTanggalOK = RS.Fields("TanggalOK").Value
    End If
    x2.ReDim 0, RS.RecordCount - 1, 0, TDBGrid2.Columns.Count - 1
    fTanggalOK = cTanggal(RS.Fields("TanggalOK").Value)
    For i = 0 To x2.UpperBound(1)
        x2(i, 0) = RS.Fields("KdDept").Value
        x2(i, 1) = RS.Fields("Dept").Value
        x2(i, 2) = RS.Fields("IdStock").Value
        x2(i, 3) = RS.Fields("NamaBarang").Value
        x2(i, 4) = RS.Fields("QTY").Value
        x2(i, 5) = RS.Fields("Satuan").Value
        x2(i, 6) = RS.Fields("Keterangan").Value
        RS.MoveNext
    Next
    TDBGrid2.Rebind
End Sub

Private Sub TDBDropDown1_DropDownClose()
    If Not IsNull(TDBDropDown1.Bookmark) Then
        col2("Departemen").Value = TDBDropDown1.Columns("Departemen").Value
    Else
        col2("Departemen").Value = ""
        col2("Kode").Value = ""
    End If
End Sub


Private Sub TDBGrid2_AfterColEdit(ByVal ColIndex As Integer)
    a = col2(ColIndex).Caption
    If a = "IdStock" Then
        SetOtherRowData col2("IdStock").Value
    End If
End Sub


Sub SetOtherRowData(ByVal tId As Long)
    TDBGrid2.SetFocus
    a = "Select top 1 IdStock, NamaBarang, Satuan from m_StockBeli where IdStock=" & tId & " and StatusProses=1 and Dept='" & col2("Departemen").Value & "'"
    query a
    If RS.RecordCount > 0 Then
        col2("IdStock").Value = RS.Fields("IdStock").Value
        col2("Nama Barang").Value = RS.Fields("NamaBarang").Value
        col2("Satuan").Value = RS.Fields("Satuan").Value
    Else
        col2("IdStock").Value = ""
        col2("Nama Barang").Value = ""
        col2("Satuan").Value = ""
    End If
End Sub

Private Sub fDari_Validate(Cancel As Boolean)
Dim MaxNoBukti As String
    If fDari.ListIndex = -1 Or Not fDari.Enabled Then
        Exit Sub
    End If
    a = "select max(NoBukti) from t_In where Tanggal>" & pAddNoLong & " and Dept='" & fDari & "'"
    query a
    MaxNoBukti = IIf(IsNull(RS.Fields(0).Value), 1, Left(RS.Fields(0).Value, 5) + 1)
    fNo = zerofill(MaxNoBukti, 5) & "/" & Right(pServerDate, 2) & "/" & zerofill(fDari.ItemData(fDari.ListIndex), 2)
    iToolbar1.SetQuick fNo
    a = "Select IdStock, NamaBarang, JumlahProses-JumlahProsesTerpakai, Satuan, '' from m_stockBeli where Dept='" & fDari & "' and JumlahProses>JumlahProsesTerpakai"
    query a
    x1.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x1.DeleteRows 0
    If RS.RecordCount > 0 Then x1.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub

Private Sub fMasterStock_Click()
    FormStockProses.LoadMe Me, col2("Departemen").Value & "", col2("Nama Barang").Value & ""
End Sub

Private Sub fOK_Click()
On Error GoTo err
    CN.BeginTrans
    a = "update t_In set StatusPindahStock=1 where NoBukti='" & fNo & "'"
    ExecMe a
    For i = 0 To x1.UpperBound(1)
        a = "update m_stockBeli set JumlahProses=JumlahProses-" & x1(i, col1("QTY").ColIndex) & ", JumlahProsesTerpakai=JumlahProsesTerpakai-" & x1(i, col1("QTY").ColIndex) & " where IdStock=" & x1(i, col1("IdStock").ColIndex)
        ExecMe a
    Next
    CN.CommitTrans
    MsgBox "SUKSES"
    GetResult iToolbar1.GetText
    Exit Sub
err:
    CN.RollbackTrans
    MsgBox "GAGAL"
End Sub

Private Sub fBatal_Click()
On Error GoTo err
    CN.BeginTrans
    a = "update t_In set StatusPindahStock=0 where NoBukti='" & fNo & "'"
    ExecMe a
    For i = 0 To x1.UpperBound(1)
        a = "update m_stockBeli set JumlahProses=JumlahProses+" & x1(i, col1("QTY").ColIndex) & ", JumlahProsesTerpakai=JumlahProsesTerpakai+" & x1(i, col1("QTY").ColIndex) & " where IdStock=" & x1(i, col1("IdStock").ColIndex)
        ExecMe a
    Next
    CN.CommitTrans
    MsgBox "SUKSES"
    GetResult iToolbar1.GetText
    Exit Sub
err:
    CN.RollbackTrans
    MsgBox "GAGAL"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Set TDBGrid1.Array = x1
    x2.ReDim 0, 0, 0, TDBGrid2.Columns.Count - 1
    x2.DeleteRows 0
    Set TDBGrid2.Array = x2
    Set TDBDropDown1.Array = z
    Set col1 = TDBGrid1.Columns
    Set col2 = TDBGrid2.Columns
    
    TDBGrid2.AllowAddNew = True
    TDBGrid2.AllowDelete = True
    
    col2("Kode").AutoDropDown = True
    col2("Kode").DropDown = TDBDropDown1
    
    col1("QTY").Alignment = dbgRight
    col2("QTY").Alignment = dbgRight
    
    col1("IdStock").Locked = True
    col1("Nama Barang").Locked = True
    col1("Satuan").Locked = True
    col2("Nama Barang").Locked = True
    col2("Departemen").Locked = True
    col2("Satuan").Locked = True
    z.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
    z.DeleteRows 0
    a = "select KdDept, Departemen from m_departemen where Proses=1 order by KdDept"
    query a
    If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
    TDBDropDown1.Rebind
    For i = 0 To z.UpperBound(1)
        fDari.List(i) = z(i, 1)
        fDari.ItemData(i) = z(i, 0)
    Next
    iToolbar1.SetClick 0
    iToolbar1_NewClick
End Sub

Private Sub Form_Resize()
    iToolbar1.Width = ScaleWidth - 2 * iToolbar1.Left
    TDBGrid1.Width = iToolbar1.Width
    TDBGrid2.Width = iToolbar1.Width
    TDBGrid2.Height = ScaleHeight - TDBGrid2.Top - 100
End Sub

Private Sub GoEvent()
    v = m_mode = "NEW" Or m_mode = "EDIT"
    fDari.Enabled = v
    fNo.Enabled = v
    fTanggal.Enabled = v
    TDBGrid1.AllowUpdate = v
    TDBGrid2.AllowUpdate = v
    fTanggalOK.Enabled = v
    fCekOK.Enabled = v
    fOK.Visible = Not v
    fBatal.Visible = Not v
End Sub

Private Sub ClearScreen()
    fNo = ""
    fDari = ""
    fTanggal = pServerDate
    x1.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x1.DeleteRows 0
    TDBGrid1.Rebind
    x2.ReDim 0, 0, 0, TDBGrid2.Columns.Count - 1
    x2.DeleteRows 0
    fTanggalOK = "__/__/__"
    fCekOK.Value = False
    TDBGrid2.Rebind
End Sub

Private Sub iToolbar1_BottomClick()
On Error Resume Next
    a = "select max(NoBukti) from t_In where Tanggal>" & pAddNoLong & " and Dept='" & fDari & "'"
    query a
    If Not IsNull(RS.Fields(0).Value) Then
        GetResult RS.Fields(0).Value
    Else
        GetResult "00000/" & Mid(iToolbar1.GetText, 7)
    End If
End Sub

Private Sub iToolbar1_CancelClick()
    GetResult iToolbar1.GetText
End Sub

Private Sub iToolbar1_DeleteClick()
On Error GoTo err
    CN.BeginTrans
    a = "delete from t_In where NoBukti='" & fNo & "'"
    ExecMe a
    a = "delete from t_Out where NoBukti='" & fNo & "'"
    ExecMe a
    For i = 0 To x1.UpperBound(1)
        a = "update m_StockBeli set JumlahProsesTerpakai=JumlahProsesTerpakai-" & cNum(x1(i, col1("QTY").ColIndex)) & " where IdStock=" & x1(i, col1("IdStock").ColIndex)
        ExecMe a
    Next
    CN.CommitTrans
    MsgBox "SUKSES"
    iToolbar1_BottomClick
    Exit Sub
err:
    CN.RollbackTrans
    MsgBox "GAGAL"
    iToolbar1.SetGagalSave
End Sub

Private Sub iToolbar1_EditClick()
    m_mode = "EDIT"
    GoEvent
    TDBGrid1.SetFocus
    fNo.Tag = fNo
    ReDim TempQTY(x1.UpperBound(1))
    For i = 0 To x1.UpperBound(1)
        TempQTY(i) = x1(i, col1("QTY").ColIndex)
    Next
End Sub

Private Sub iToolbar1_ExitClick()
    Unload Me
End Sub

Private Sub iToolbar1_ListClick()
    FormList.LoadMe "BELUM OK@SUDAH OK", _
        "select distinct NoBukti, Tanggal, Dept from t_In where StatusOut=0@select distinct NoBukti, Tanggal, Dept from t_In where StatusOut=1", _
        "Tanggal@Departemen", "Tanggal@Departemen", "1000@2500", "Date@String", _
        "No Bukti@Tanggal@Departemen", "1500@1500@2500", "String@Date@String", Me, "", 0
End Sub

Private Sub iToolbar1_NewClick()
On Error Resume Next
    ClearScreen
    m_mode = "NEW"
    GoEvent
    fDari.SetFocus
    ReDim TempQTY(0)
End Sub

Private Sub iToolbar1_NextClick()
On Error Resume Next
    a = iToolbar1.GetText
    Mid(a, 1) = zerofill(Left(a, 5) + 1, 5)
    GetResult a
End Sub

Private Sub iToolbar1_PrevClick()
On Error Resume Next
    a = iToolbar1.GetText
    Mid(a, 1) = zerofill(Left(a, 5) - 1, 5)
    GetResult a
End Sub

Private Sub iToolbar1_SaveClick()
On Error GoTo err
    CN.BeginTrans
    TDBGrid1.Update
    TDBGrid2.Update
    If fCekOK And cD(fTanggalOK) = 0 Then
        MsgBox "Tanggal OK Harus Diisi"
        GoTo err
    End If
    If cD(fTanggalOK) <> 0 Then fCekOK.Value = 1
    If m_mode <> "NEW" Then
        a = "delete from t_In where NoBukti='" & fNo.Tag & "'"
        ExecMe a
        a = "delete from t_Out where NoBukti='" & fNo.Tag & "'"
        ExecMe a
    End If
    For i = 0 To x1.UpperBound(1)
        If x1(i, col1("QTY").ColIndex) <> 0 Then
            a = "insert into t_In(NoBukti,Tanggal,IdStock,QTY,Keterangan,Dept, StatusOut, TanggalOK) values('" & _
                fNo & _
                "'," & cD(fTanggal) & _
                "," & x1(i, col1("IdStock").ColIndex) & _
                "," & x1(i, col1("QTY").ColIndex) & _
                ",'" & x1(i, col1("Keterangan").ColIndex) & _
                "','" & fDari & _
                "'," & IIf(fCekOK.Value, 1, 0) & _
                "," & cD(fTanggalOK) & ")"
            If ExecMe(a) = 0 Then GoTo err
            If i <= UBound(TempQTY) Then
                c = TempQTY(i)
            Else
                c = 0
            End If
            a = "update m_StockBeli set JumlahProsesTerpakai=JumlahProsesTerpakai-" & c & "+" & x1(i, col1("QTY").ColIndex) & " where IdStock=" & x1(i, col1("IdStock").ColIndex)
            ExecMe a
        End If
    Next
    For i = 0 To x2.UpperBound(1)
        a = "insert into t_Out(NoBukti,Tanggal,IdStock,QTY,Keterangan,Dept,StatusOK,TanggalOK,DeptAsal) values('" & _
            fNo & _
            "'," & cD(fTanggal) & _
            "," & x2(i, col2("IdStock").ColIndex) & _
            "," & x2(i, col2("QTY").ColIndex) & _
            ",'" & x2(i, col2("Keterangan").ColIndex) & _
            "','" & x2(i, col2("Departemen").ColIndex) & _
            "'," & IIf(fCekOK.Value, 1, 0) & _
            "," & cD(fTanggalOK) & _
            ",'" & fDari & "')"
        If ExecMe(a) = 0 Then GoTo err
    Next
    CN.CommitTrans
    MsgBox "SUKSES"
    GetResult fNo
    Exit Sub
err:
    CN.RollbackTrans
    iToolbar1.SetGagalSave
    MsgBox "GAGAL"
End Sub

Private Sub iToolbar1_TextEnter()
    GetResult iToolbar1.GetText
End Sub

Private Sub iToolbar1_TopClick()
On Error Resume Next
    a = "select min(NoBukti) from t_In where left(NoBukti,5)='" & Left(iToolbar1.GetText, 5) & "'"
    query a
    a = IIf(IsNull(RS.Fields(0).Value), Left(iToolbar1.GetText, 6), RS.Fields(0).Value)
    GetResult a
End Sub


VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormSJ 
   BackColor       =   &H00FFC0C0&
   Caption         =   "SURAT JALAN"
   ClientHeight    =   7890
   ClientLeft      =   -30
   ClientTop       =   180
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11880
   Tag             =   "2"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fPrintSPP 
      Caption         =   "PRINT SPP"
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      Top             =   240
      Width           =   1095
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3836
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Jenis"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Kode"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Warna"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No Warna"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tube"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Grade"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Harga"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "SatBesar"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Jumlah1"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Jumlah2"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Dtl"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Keterangan"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "NoSC"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "IdSC"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "SatKecil"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "IdStock"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=64478624"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=64365648"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1005"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=926"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=1482"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1402"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1376"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1296"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(4)._MinWidth=1312901971"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=1005"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=926"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=1402"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=1323"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=1349"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=1270"
      Splits(0)._ColumnProps(35)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(36)=   "Column(8).Width=1217"
      Splits(0)._ColumnProps(37)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(8)._WidthInPix=1138"
      Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(40)=   "Column(9).Width=1455"
      Splits(0)._ColumnProps(41)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(9)._WidthInPix=1376"
      Splits(0)._ColumnProps(43)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(44)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(45)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(48)=   "Column(11).Width=2778"
      Splits(0)._ColumnProps(49)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(11)._WidthInPix=2699"
      Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(52)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(53)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(55)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(56)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(57)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(59)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(60)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(61)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(63)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(64)=   "Column(14)._MinWidth=1573376"
      Splits(0)._ColumnProps(65)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(69)=   "Column(15)._MinWidth=1573376"
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
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=86,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=98,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=95,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=96,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=97,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=74,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=78,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=82,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=90,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=94,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
      _StyleDefs(94)  =   "Named:id=33:Normal"
      _StyleDefs(95)  =   ":id=33,.parent=0"
      _StyleDefs(96)  =   "Named:id=34:Heading"
      _StyleDefs(97)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(98)  =   ":id=34,.wraptext=-1"
      _StyleDefs(99)  =   "Named:id=35:Footing"
      _StyleDefs(100) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(101) =   "Named:id=36:Selected"
      _StyleDefs(102) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(103) =   "Named:id=37:Caption"
      _StyleDefs(104) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(105) =   "Named:id=38:HighlightRow"
      _StyleDefs(106) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(107) =   "Named:id=39:EvenRow"
      _StyleDefs(108) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(109) =   "Named:id=40:OddRow"
      _StyleDefs(110) =   ":id=40,.parent=33"
      _StyleDefs(111) =   "Named:id=41:RecordSelector"
      _StyleDefs(112) =   ":id=41,.parent=34"
      _StyleDefs(113) =   "Named:id=42:FilterBar"
      _StyleDefs(114) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton fTotal 
      Caption         =   "TOTAL"
      Height          =   375
      Left            =   4620
      TabIndex        =   31
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox fKota 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   9
      Top             =   1920
      Width           =   4935
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   1500
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   300
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton fList 
      Caption         =   "LIST"
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   240
      Width           =   375
   End
   Begin UsrText.IText fNoSJ 
      Height          =   270
      Left            =   2160
      TabIndex        =   26
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.TextBox fAlamatPenerima 
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   1560
      Width           =   4935
   End
   Begin UsrText.IText fNamaPenerima 
      Height          =   270
      Left            =   4920
      TabIndex        =   7
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fTanggalSPP 
      Height          =   270
      Left            =   7320
      TabIndex        =   10
      Top             =   2520
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
   Begin UsrText.IText fNoSPP 
      Height          =   270
      Left            =   4920
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fTanggalKirim 
      Height          =   270
      Left            =   8280
      TabIndex        =   12
      Top             =   2520
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
   Begin UsrText.IText fKeteranganSPP 
      Height          =   270
      Left            =   6600
      TabIndex        =   24
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fTanggalSJ 
      Height          =   270
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
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
   Begin UsrText.IText fSopir 
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fNoKendaraan 
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fAngkutan 
      Height          =   270
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fQuick 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   3960
      Width           =   9735
   End
   Begin VB.Label fUpdateSJBy 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   34
      Top             =   3540
      Width           =   4335
   End
   Begin VB.Label fUpdatedBy 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      TabIndex        =   33
      Top             =   3060
      Width           =   4215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN SPP"
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA ANGKUTAN"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NOMOR KENDARAAN"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA SOPIR"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL SJ"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NO SJ"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SURAT JALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5460
      TabIndex        =   18
      Top             =   180
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
      Height          =   255
      Left            =   7320
      TabIndex        =   17
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "NO SPP"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT PENERIMA"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA PENERIMA"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL KIRIM"
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "FormSJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mKodeCustomer As Integer
Dim x As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim rs1() As Variant
Dim iJenis As Integer
Dim iKode As Integer
Dim iWarna As Integer
Dim iNoWarna As Integer
Dim iTube As Integer
Dim iGrade As Integer
Dim iHarga As Integer
Dim iJumlah1 As Integer
Dim iJumlah2 As Integer
Dim iKeterangan As Integer
Dim iNoSC As Integer
Dim iIdSC As Integer
Dim iSatBesar As Integer
Dim iSatKecil As Integer
Dim iDtl As Integer

Dim iIdStock As Integer
Dim tReadTime As Date
Dim tags() As ArrayList

Dim coll As ClassProperties

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub GoEvent(ByVal tEvent As String)
    If tEvent = "SETUJU" Then
        m_mode = "SETUJU"
        tProp = 1
    ElseIf tEvent = "SJ" Then
        m_mode = "SJ"
        tProp = 2
    ElseIf tEvent = "BELUM SETUJU" Then
        m_mode = "BELUM SETUJU"
        tProp = 0
    End If
    v = IIf((tProp And 1) = 0, False, True)
        fTanggalSJ.Enabled = v
        fSopir.Enabled = v
        fAngkutan.Enabled = v
        fNoKendaraan.Enabled = v
        TDBGrid1.AllowUpdate = v
        fSave.Enabled = v
        fPrint.Enabled = Not v
    v = IIf((tProp And 2) = 0, False, True)
        fDelete.Enabled = v
End Sub

Private Sub fDelete_Click()
On Error GoTo err
    BeginTransaction
    If Not cekValid("DELETE", Tag) Then GoTo err
    b = 0
    a = "update t_SPP~ set Status=1, WaktuUpdate=now where Pelunasan=0 and NoSPP='" & esc(fNoSPP) & "' and Status=2"
    b = ExecMe(a)
    If b = 0 Then
        MsgBox "Sudah Bikin Kwitansi"
        GoTo err
    End If
    For i = 0 To x.UpperBound(1)
        a = "update t_SPPDetail" & pTipe & " set StatusDetail=1 where NoSPP='" & esc(fNoSPP) & "' and IdSPP=" & i
        ExecMe a
        a = "update t_SCDetail" & pTipe & " set Terpakai=Terpakai-" & cNum(x(i, iJumlah2)) & " where NoSC='" & esc(x(i, iNoSC)) & "' and IdSC=" & x(i, iIdSC)
        ExecMe a
        a = "update m_stock~ set JumlahBox=JumlahBox+" & x(i, iJumlah1) & ",JumlahKG=JumlahKG+" & cNum(x(i, iJumlah2)) & " where IdStock=" & x(i, iIdStock)
        If ExecMe(a) = 0 Then GoTo err
    Next
    s = "select now() as a"
    query s
    tReadTime = RS.Fields(0).value
    CommitTransaction
    MsgBox "SUKSES"
    m_mode = "SETUJU"
    GoEvent "SETUJU"
    'SendData "1HAPUS SJ NO: " & fNoSJ & Chr(8)
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fList_Click()
    FormList.LoadMe "BELUM SJ@SUDAH SJ@BELUM SETUJU", _
        "select NoSPP, TanggalSPP, TanggalKirim, NamaPenerima, 0, '' from t_SPP" & pTipe & " left join m_customer on m_customer.Kode=t_SPP" & pTipe & ".Kode where status=1@" & _
        "select NoSPP, TanggalSPP, TanggalKirim, NamaPenerima, TanggalSJ, NamaSopir from t_SPP" & pTipe & " left join m_customer on m_customer.Kode=t_SPP" & pTipe & ".Kode where status>=2@" & _
        "select NoSPP, TanggalSPP, TanggalKirim, NamaPenerima, TanggalSJ, '' from t_SPP" & pTipe & " left join m_customer on m_customer.Kode=t_SPP" & pTipe & ".Kode where status=0", _
        "Nama Penerima@Tanggal Kirim@Tanggal SJ", "NamaPenerima@TanggalKirim@TanggalSJ", "2500@1000@1000", "String@Date@Date", _
        "No SPP@Tanggal SPP@Tanggal Kirim@Nama Penerima@Tanggal SJ@Nama Sopir", _
        "1500@1000@1000@3000@1000@2000", _
        "String@Date@Date@String@Date@String", Me, " order by TanggalSPP\10000 Desc, NoSPP Desc"
    FormList.Show , Me
End Sub

Private Sub ClearScreen()
    fUpdatedBy = ""
    fNoSJ = ""
    fTanggalSJ = pServerDate
    fSopir = ""
    fNoKendaraan = ""
    fAngkutan = ""
    fNamaPenerima = ""
    fAlamatPenerima = ""
    fKota = ""
    fNoSPP = ""
    fTanggalSPP = ""
    fTanggalKirim = ""
    fKeteranganSPP = ""
    x.ReDim 0, 0, 0, col1.count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
End Sub

Sub GetResult(ByVal tNo As String)
On Error Resume Next
    s = "select now() as a"
    query s
    tReadTime = RS.Fields(0).value
    ClearScreen
    fQuick = Left(tNo, 5) & "/" & Right(tNo, 2)
    fUpdateSJBy.Caption = ""
    
    a = "select Kode, IdtSJ, IdtSPP, Status, NoSJ, TanggalSJ, NamaSopir, NamaAngkutan, NoKendaraan, NamaPenerima, AlamatPenerima, NoSPP, TanggalSPP, TanggalKirim, KeteranganSPP, AlamatPendek, Pengupdate, WaktuUpdate from t_SPP~ where ShortSPP='" & esc(fQuick) & "'"
    query a
    If RS.RecordCount = 0 Then
        Exit Sub
    End If
    mKodeCustomer = RS!Kode
    fUpdateSJBy.Caption = GetID(RS!IdtSJ)
    fNoSJ = RS.Fields("NoSJ").value & ""
    a = cTanggal(RS.Fields("TanggalSJ").value)
    If a <> "__/__/__" Then fTanggalSJ = a
    fSopir = RS.Fields("NamaSopir").value & ""
    fAngkutan = RS.Fields("NamaAngkutan").value & ""
    fNoKendaraan = RS.Fields("NoKendaraan").value & ""
    fNamaPenerima = RS.Fields("NamaPenerima").value
    fAlamatPenerima = RS.Fields("AlamatPenerima").value
    fNoSPP = RS.Fields("NoSPP").value
    fTanggalSPP = cTanggal(RS.Fields("TanggalSPP").value)
    fTanggalKirim = cTanggal(RS.Fields("TanggalKirim").value)
    fKeteranganSPP = RS.Fields("KeteranganSPP").value
    fKota = RS.Fields("AlamatPendek").value
    fUpdatedBy = RS!Pengupdate & " " & RS!WaktuUpdate & " " & RS!IdtSPP
    If RS.Fields("Status").value = 1 Then
        fNoSJ = Left(fNoSPP, 5)
        fTanggalSJ_LostFocus
        GoEvent "SETUJU"
    ElseIf RS.Fields("Status").value = 2 Then
        GoEvent "SJ"
    Else
        GoEvent "BELUM SETUJU"
    End If
    a = "select Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, Harga, SatBesar, t_SPPDetail" & pTipe & ".JumlahBox, t_SPPDetail" & pTipe & ".JumlahKg, Dtl, KeteranganSPPDetail, NoSC, IdSC, SatKecil, t_SPPDetail" & pTipe & ".IdStock from t_SPPDetail" & pTipe & " left join m_stock" & pTipe & " on t_SPPDetail" & pTipe & ".IdStock=m_stock" & pTipe & ".IdStock " & _
            "where NoSPP='" & esc(fNoSPP) & "'"
    query a
    If RS.RecordCount > 0 Then
        x.LoadRows RS.GetRows
        TDBGrid1.Rebind
    End If
    ReDim tags(x.UpperBound(1))
End Sub

Private Sub fNext_Click()
On Error Resume Next
    a = fQuick
    Mid(a, 1) = zerofill(Left(a, 5) + 1, 5)
    fQuick = a
    GetResult fQuick
err:
End Sub

Private Sub Form_Load()
    iJenis = 0
    iKode = 1
    iWarna = 2
    iNoWarna = 3
    iTube = 4
    iGrade = 5
    iHarga = 6
    iSatBesar = 7
    iJumlah1 = 8
    iJumlah2 = 9
    iDtl = 10

    iKeterangan = 11
    iNoSC = 12
    iIdSC = 13
    iSatKecil = 14
    iIdStock = 15
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    fNoSJ.Enabled = False
    fSopir.Enabled = False
    fNoKendaraan.Enabled = False
    fAngkutan.Enabled = False
    fNamaPenerima.Enabled = False
    fAlamatPenerima.Enabled = False
    fNoSPP.Enabled = False
    fTanggalSPP.Enabled = False
    fTanggalKirim.Enabled = False
    fKeteranganSPP.Enabled = False
    For i = 0 To col1.count - 1
        col1(i).Locked = True
    Next
    'Col1(iJUMLAH1).Locked = False
    'col1(iJumlah2).Locked = False
    col1(iJumlah2).Tag = "Decimal"
    col1(iJumlah1).Tag = "Integer"
    TDBGridLoad TDBGrid1
    TDBGridSetVisible TDBGrid1, "Harga@NoSC@IdSC@SatKecil@IdStock", False
    a = "select min(NoSPP) from t_SPP~ where status=1"
    query a
    If Not IsNull(RS.Fields(0).value) Then
        fQuick = Left(RS.Fields(0).value, 5) & "/" & Right(RS.Fields(0).value, 2)
        fQuick_KeyDown 13, 0
        Exit Sub
    End If
    a = "select max(NoSPP) from t_SPP~ where TanggalSPP>" & pAddNoLong
    query a
    If Not IsNull(RS.Fields(0).value) Then
        fQuick = Left(RS.Fields(0).value, 5) & "/" & Right(RS.Fields(0).value, 2)
        fQuick_KeyDown 13, 0
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    fKet.Width = ScaleWidth - 2 * fKet.Left
End Sub

Private Sub fPrev_Click()
On Error Resume Next
    a = fQuick
    Mid(a, 1) = zerofill(Left(a, 5) - 1, 5)
    fQuick = a
    GetResult fQuick
err:
End Sub

Private Sub fPrint_Click()
    FormPreview.LoadMe Me, "SJ", fNoSJ
End Sub

Private Sub fPrintSPP_Click()
    FormPreview.LoadMe Me, "SPP", fNoSPP & "@HideHarga"
End Sub

Private Sub fSave_Click()
'On Error GoTo err
Dim total As Double
    BeginTransaction
    If Not cekValid("EDIT", Tag) Then GoTo err
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        
        
        If x(i, iSatBesar) = x(i, iSatKecil) Then
            If CLng(x(i, iJumlah1)) <> CLng(x(i, iJumlah2)) Then
                MsgBox "Jumlah Harus sama karena Satuan juga sama"
                GoTo err
            End If
        End If
        If x(i, iJumlah1) <> 0 Then
            Jumlah = x(i, iJumlah2)
            If Jumlah = 0 Then
                MsgBox "JUMLAH TIDAK BOLEH NOL"
                GoTo err
            End If
        End If
    Next
    
    a = "update t_SPP~ set status=2," & _
        "NoSJ='" & fNoSJ & _
        "',TanggalSJ=" & cD(fTanggalSJ) & _
        ",NamaSopir='" & fSopir & _
        "',NamaAngkutan='" & fAngkutan & _
        "',IdtSJ=" & cNum(Rnd * 2000000000, 0) & _
        ",NoKW='', NoKendaraan='" & fNoKendaraan & "', WaktuUpdate=now where NoSPP='" & esc(fNoSPP) & "' and WaktuUpdate<=#" & Format(tReadTime, "mm/dd/yyyy hh:nn:ss") & "#"
    b = ExecMe(a)
    
    If b <= 0 Then
        MsgBox "Data surat jalan sudah berubah"
        GoTo err
    End If
    For i = 0 To x.UpperBound(1)
        Dim NoBukti As Long
        Dim jumlahKG As Double
        Dim jumlahBox As Long
        'jumlahKG = 0
        'NoBukti = 0
        'Dim d3() As String
        'Dim d2() As String
        'Dim d() As String
        'd3 = Split(x(i, iDtl3), "_")
        'd2 = Split(x(i, iDtl2), "_")
        'd = Split(x(i, iDtl), "_")
        'For j = 0 To UBound(d3)
        '    If d3(j) = -1 Then
        '        If NoBukti = 0 Then
        '            a = "select min(NoBukti) from t_InputStock~ where NoBukti<0"
        '            query a
        '            If Not IsNull(RS.Fields(0).value) Then
        '                NoBukti = RS.Fields(0).value
        '            End If
        '        End If
        '        NoBukti = NoBukti - 1
        '
        '       a = "insert into t_InputStockDetail~(NoUrut, Kg, NoBukti)"
        '        a = a & " values(" & d(j) & "," & cNum(d2(j)) & "," & NoBukti & ")"
        '        ExecMe a
        '        a = "SELECT @@Identity"
        '        query a
        '        d3(j) = RS.Fields(0).value
        '        jumlahKG = jumlahKG + d2(j)
        '        JumlahBox = JumlahBox + 1
        '    End If
        'Next
        'If NoBukti < 0 Then
        '    a = "insert into t_InputStock~(NoBukti, Tanggal, TanggalGudang, IDStock, n1, n2, status)"
        '    a = a & " values(" & NoBukti & "," & cD(fTanggalSJ) & "," & cD(fTanggalSJ) & "," & x(i, iIdStock) & "," & JumlahBox & "," & cNum(jumlahKG) & ", 10)"
        '    ExecMe a
        'End If
        
        
        If x(i, iJumlah1) <> 0 Then
            a = "update t_SPPDetail~ set NamaSopirDetail='" & fSopir & _
            "', NamaAngkutanDetail='" & fAngkutan & _
            "', NoKendaraanDetail='" & fNoKendaraan & _
            "',StatusDetail=2" & _
            ", TanggalDetail=" & cD(fTanggalSJ) & _
            ", JumlahKg=" & cNum(x(i, iJumlah2)) & _
            ", dtl='" & x(i, iDtl) & "'" & _
            ", NoSJ='" & fNoSJ & "' where NoSPP='" & esc(fNoSPP) & "' and IdSPP=" & i
            If ExecMe(a) = 0 Then GoTo err
            
            a = "update t_SCDetail~ set Terpakai=Terpakai+" & cNum(x(i, iJumlah2)) & " where NoSC='" & esc(x(i, iNoSC)) & "' and IdSC=" & x(i, iIdSC)
            ExecMe a
            s = "select * from t_SCDetail~ where NoSC='" & esc(x(i, iNoSC)) & "' and IdSC=" & x(i, iIdSC)
            query s
            If Round(RS!Jumlah - RS!Terpakai, 0) < 0 Then
                MsgBox "Penjualan Barang melebihi Kontrak (" & x(i, iJumlah2) & ")"
                GoTo err
            End If
            a = "update m_stock~ set JumlahBox=JumlahBox-" & x(i, iJumlah1) & ",JumlahKG=JumlahKG-" & cNum(x(i, iJumlah2)) & " where IdStock=" & x(i, iIdStock)
            ExecMe a
        End If
    Next
    
    s = "select sum(Harga*JumlahKG) as Total from t_SPPDetail~ where NoSPP='" & esc(fNoSPP) & "'"
    query s
    s = "update t_SPP~ set Total=" & cNum(RS!total) & " where NoSPP='" & esc(fNoSPP) & "'"
    ExecMe s
    
    total = 0
    s = "select sum(Total-Pelunasan) as a from t_SPPPE where Kode=" & mKodeCustomer
    query s
    total = IIf(IsNull(RS!a), 0, RS!a)
    s = "select sum(Total-Pelunasan) as a from t_SPPDTY where Kode=" & mKodeCustomer
    query s
    total = total + IIf(IsNull(RS!a), 0, RS!a)
    
    s = "select sum(Nilai) as a from t_STTPelunasanDTY where KodeCustomer=" & mKodeCustomer & " and TanggalPelunasan>" & cD(fTanggalSJ.Text)
    query s
    total = total + IIf(IsNull(RS!a), 0, RS!a)
    
    s = "select sum(Nilai) as a from t_STTPelunasanPE where KodeCustomer=" & mKodeCustomer & " and TanggalPelunasan>" & cD(fTanggalSJ.Text)
    query s
    total = total + IIf(IsNull(RS!a), 0, RS!a)
    
    s = "select Limit from m_Customer where Kode=" & mKodeCustomer
    query s
    If total > RS!Limit Then
        MsgBox "Lewat Limit!"
        GoTo err
    End If
    
    'total = 0
    'Dim y As Long
    'y = (Year(Now) Mod 100) * 10000
    
    's = "select sum(Total) as a from t_SPPPE where NamaPenerima='" & esc(fNamaPenerima.Text) & "' and TanggalSJ between " & y & " and " & (y + 10000)
    'query s
    'total = total + IIf(IsNull(RS!a), 0, RS!a)
    's = "select sum(Total) as a from t_SPPDTY where NamaPenerima='" & esc(fNamaPenerima.Text) & "' and TanggalSJ between " & y & " and " & (y + 10000)
    'query s
    'total = total + IIf(IsNull(RS!a), 0, RS!a)
    'If total / 1000000 > 590 Then
    '    MsgBox "Lewat Limit 600 juta!"
    '    GoTo err
    'End If
    
    CommitTransaction
    MsgBox "SUKSES"
    'SendData "1BUAT SJ NO: " & fNoSJ & Chr(8)
    GetResult fNoSJ
    DoEvents
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fTanggalSJ_LostFocus()
    BuatNomor fNoSJ, fTanggalSJ, pNomorSJ, fQuick, "select ''"
End Sub

Private Sub fQuick_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        a = fQuick
        If Left(Right(a, 3), 1) = "/" Then
            b = Right(a, 2)
        Else
            b = Right(pServerDate, 2)
        End If
        a = zerofill(Left(a, 5), 5) & "/" & b
        fQuick = a
        GetResult fQuick
    End If
    fQuick.Cancel = True
End Sub

Private Sub fTotal_Click()
    TDBGrid1.Update
    a = 0
    For i = 0 To x.UpperBound(1)
       a = a + x(i, iJumlah2)
    Next
    MsgBox cDecimal(a)
End Sub

Private Sub TDBGrid1_DblClick()
    If Not TDBGrid1.AllowUpdate Then Exit Sub
    FormDetailKG.LoadMe TDBGrid1.Columns(iJumlah1).value, TDBGrid1.Columns(iDtl).value
    TDBGrid1.Columns(iDtl).value = FormDetailKG.r_Kgs
    TDBGrid1.Columns(iJumlah2).value = FormDetailKG.r_kg
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 68 Then
        TDBGrid1_DblClick
        Exit Sub
    End If
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'fKet.Caption = TDBGrid1.Columns(iDtl2).value
End Sub

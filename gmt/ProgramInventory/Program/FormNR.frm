VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormNR 
   BackColor       =   &H00FFC0C0&
   Caption         =   "RETUR"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "42"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fPrint3 
      Caption         =   "PRINT 3"
      Height          =   375
      Left            =   10200
      TabIndex        =   26
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton fPrint4 
      Caption         =   "PRINT 4"
      Height          =   375
      Left            =   10200
      TabIndex        =   25
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton fPrint2 
      Caption         =   "PRINT 2"
      Height          =   375
      Left            =   9240
      TabIndex        =   24
      Top             =   600
      Width           =   855
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2655
      Left            =   180
      TabIndex        =   13
      Top             =   2880
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No SJ"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Jenis"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kode"
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
      Columns(7).Caption=   "SatBesar"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "SatKecil"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Harga"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Jumlah"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "idstock"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3440"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3360"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0)._MinWidth=12632256"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1693"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1535"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1455"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(2)._MinWidth=221"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=1323"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1244"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(3)._MinWidth=221"
      Splits(0)._ColumnProps(28)=   "Column(4).Width=1508"
      Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(34)=   "Column(4)._MinWidth=221"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=1482"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=1402"
      Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(41)=   "Column(5)._MinWidth=221"
      Splits(0)._ColumnProps(42)=   "Column(6).Width=1138"
      Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=1058"
      Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(48)=   "Column(6)._MinWidth=221"
      Splits(0)._ColumnProps(49)=   "Column(7).Width=1270"
      Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=1191"
      Splits(0)._ColumnProps(52)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(54)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(55)=   "Column(7)._MinWidth=221"
      Splits(0)._ColumnProps(56)=   "Column(8).Width=1270"
      Splits(0)._ColumnProps(57)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(8)._WidthInPix=1191"
      Splits(0)._ColumnProps(59)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(61)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(62)=   "Column(8)._MinWidth=221"
      Splits(0)._ColumnProps(63)=   "Column(9).Width=1799"
      Splits(0)._ColumnProps(64)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(9)._WidthInPix=1720"
      Splits(0)._ColumnProps(66)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(67)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(68)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(69)=   "Column(9)._MinWidth=221"
      Splits(0)._ColumnProps(70)=   "Column(10).Width=1773"
      Splits(0)._ColumnProps(71)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(10)._WidthInPix=1693"
      Splits(0)._ColumnProps(73)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(74)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(75)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(76)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(77)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(79)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(80)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(81)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(11).Order=12"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   2
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
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=125,.parent=2,.namedParent=127"
      _StyleDefs(17)  =   "FilterBarStyle:id=128,.parent=1,.namedParent=130"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=126,.parent=125"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=129,.parent=128"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=76,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=73,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=74,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=75,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=24,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=40,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=37,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=38,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=39,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=44,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=41,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=42,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=43,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=48,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=45,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=46,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=47,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=52,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=49,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=51,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=56,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=53,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=54,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=55,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=60,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=57,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=58,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=59,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=72,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=69,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=71,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=92,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=89,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=90,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=91,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=124,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=121,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=122,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=123,.parent=15"
      _StyleDefs(78)  =   "Named:id=29:Normal"
      _StyleDefs(79)  =   ":id=29,.parent=0"
      _StyleDefs(80)  =   "Named:id=30:Heading"
      _StyleDefs(81)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=30,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=31:Footing"
      _StyleDefs(84)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=32:Selected"
      _StyleDefs(86)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=33:Caption"
      _StyleDefs(88)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(89)  =   "Named:id=34:HighlightRow"
      _StyleDefs(90)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(91)  =   "Named:id=35:EvenRow"
      _StyleDefs(92)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=36:OddRow"
      _StyleDefs(94)  =   ":id=36,.parent=29"
      _StyleDefs(95)  =   "Named:id=127:RecordSelector"
      _StyleDefs(96)  =   ":id=127,.parent=30"
      _StyleDefs(97)  =   "Named:id=130:FilterBar"
      _StyleDefs(98)  =   ":id=130,.parent=29"
   End
   Begin VB.CommandButton fTotal 
      Caption         =   "TOTAL"
      Height          =   375
      Left            =   10920
      TabIndex        =   23
      Top             =   1080
      Width           =   735
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No SJ"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Jenis"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kode"
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
      Columns(7).Caption=   "SatBesar"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "SatKecil"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Jumlah1"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Jumlah2"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Harga"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "QTY Disc"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Disc"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "IdStock"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=3296"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=900"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=820"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=64659712"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2434"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2355"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=64368576"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=979"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=900"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1138"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1058"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(4)._MinWidth=34"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1138"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1058"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=900"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=820"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=1270"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=1191"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=1217"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=1138"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=1164"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=1085"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(45)=   "Column(10).Width=1349"
      Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=1270"
      Splits(0)._ColumnProps(48)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(49)=   "Column(11).Width=1931"
      Splits(0)._ColumnProps(50)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(11)._WidthInPix=1852"
      Splits(0)._ColumnProps(52)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(53)=   "Column(12).Width=1111"
      Splits(0)._ColumnProps(54)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(12)._WidthInPix=1032"
      Splits(0)._ColumnProps(56)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(57)=   "Column(12)._MinWidth=-1"
      Splits(0)._ColumnProps(58)=   "Column(13).Width=1773"
      Splits(0)._ColumnProps(59)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(13)._WidthInPix=1693"
      Splits(0)._ColumnProps(61)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(62)=   "Column(13)._MinWidth=-1"
      Splits(0)._ColumnProps(63)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(64)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(66)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(67)=   "Column(14)._MinWidth=-1"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=98,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=95,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=96,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=97,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=102,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=99,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=100,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=101,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
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
      _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(90)  =   "Named:id=33:Normal"
      _StyleDefs(91)  =   ":id=33,.parent=0"
      _StyleDefs(92)  =   "Named:id=34:Heading"
      _StyleDefs(93)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(94)  =   ":id=34,.wraptext=-1"
      _StyleDefs(95)  =   "Named:id=35:Footing"
      _StyleDefs(96)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(97)  =   "Named:id=36:Selected"
      _StyleDefs(98)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=37:Caption"
      _StyleDefs(100) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(101) =   "Named:id=38:HighlightRow"
      _StyleDefs(102) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(103) =   "Named:id=39:EvenRow"
      _StyleDefs(104) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(105) =   "Named:id=40:OddRow"
      _StyleDefs(106) =   ":id=40,.parent=33"
      _StyleDefs(107) =   "Named:id=41:RecordSelector"
      _StyleDefs(108) =   ":id=41,.parent=34"
      _StyleDefs(109) =   "Named:id=42:FilterBar"
      _StyleDefs(110) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   9240
      TabIndex        =   19
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton fList 
      Caption         =   "&LIST"
      Height          =   375
      Left            =   10320
      TabIndex        =   18
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox fMataUang 
      Height          =   285
      Left            =   6240
      TabIndex        =   17
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton fNew 
      Caption         =   "&NEW"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox fQuick 
      Height          =   315
      Left            =   7320
      TabIndex        =   14
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   2280
      TabIndex        =   7
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
   Begin UsrText.IText fNo 
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
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
   Begin UsrText.IText fNoKW 
      Height          =   270
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
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
   Begin UsrText.IText fCustomer 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
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
   Begin UsrText.IText fTanggalKW 
      Height          =   270
      Left            =   4920
      TabIndex        =   12
      Top             =   360
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
   Begin UsrText.IText fKetNR 
      Height          =   270
      Left            =   3480
      TabIndex        =   20
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN"
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "MATA UANG"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL KW"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LIHAT RETUR LAIN"
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL NR"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NO RETUR"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NO KW"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA CUSTOMER"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FormNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As New XArrayDB
Dim fAlamat As String
Dim x As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim iNoSJ As Integer
Dim iJenis As Integer
Dim iKode As Integer
Dim iWarna As Integer
Dim iNoWarna As Integer
Dim iTube As Integer
Dim iGrade As Integer
Dim iSatBesar As Integer
Dim iSatKecil As Integer
Dim iJumlah1 As Integer
Dim iJumlah2 As Integer
Dim iHarga As Integer
Dim iDiscKG As Integer
Dim iDisc As Integer
Dim iIdStock As Integer

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub


Private Sub GoEvent(ByVal tEvent As String)
    If tEvent = "NEW" Then
        m_mode = "NEW"
        tProp = 2
    ElseIf tEvent = "SEE" Then
        m_mode = "SEE"
        tProp = 1
    End If
    v = IIf((tProp And 1) = 0, False, True)
        fDelete.Enabled = v
        fPrint.Enabled = v
    v = IIf((tProp And 2) = 0, False, True)
        fNo.Enabled = v
        fTanggal.Enabled = v
        fSave.Enabled = v
        TDBGrid1.AllowUpdate = v
        fMataUang.Enabled = v
        fKetNR.Enabled = v
End Sub

Private Sub fList_Click()
    FormList.LoadMe "NOTA RETUR", _
"select distinct NoNR, TanggalNR, t_NR" & pTipe & ".NoKW, TanggalKW, m_customer.Nama, t_NR" & pTipe & ".Total, left(TanggalNR,2) from t_NR" & pTipe & " left join m_customer on m_customer.Kode=t_NR" & pTipe & ".KodeCustomer where 1=1", _
"Nama Customer@No NR@Tanggal NR@No KW@Tanggal KW", _
"m_Customer.Nama@NoNR@TanggalNR@t_NR" & pTipe & ".NoKW@TanggalKW", _
"2500@1000@1000@1000@1000", _
"String@String@Date@String@Date", _
"NO NR@TANGGAL NR@NO KW@TANGGAL KW@NAMA CUSTOMER@TOTAL", _
"1700@1000@1700@1000@2500@1500", "String@Date@String@Date@String@Decimal", Me, _
" order by left(TanggalNR,2), NoNR"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 200
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
End Sub

Private Sub fPrint_Click()
    FormPreview.LoadMe Me, "NR", fNo & "@"
End Sub

Private Sub UpdatefQuick()
    temp(0) = fQuick
    a = "select NoNR from t_NR" & pTipe & " where NoKW='" & fNoKW & "'"
    query a
    fQuick.Clear
    If RS.RecordCount > 0 Then
        For i = 0 To RS.RecordCount - 1
            fQuick.List(i) = RS.Fields(0).Value
            RS.MoveNext
        Next
    End If
    fQuick = temp(0)
End Sub

Sub GetResult(ByVal tNo As String)
    If tNo = "" Then Exit Sub
    fQuick = Left(tNo, 5) & "/" & Right(tNo, 2)
    a = "select Nama, t_NR~.NoKW, TanggalKW, NoNR, TanggalNR, MataUang, StatusNR, Ket " & _
        "from t_NR~ left join m_customer on t_NR~.KodeCustomer=m_customer.Kode where ShortNR='" & fQuick & "'"
    query a
    GoEvent "SEE"
    If RS.RecordCount = 0 Then Exit Sub
    If RS.Fields("StatusNR").Value = 1 Then fDelete.Enabled = False
    fCustomer = RS.Fields("Nama").Value
    fNo = RS.Fields("NoNR").Value
    fTanggal = cTanggal(RS.Fields("TanggalNR").Value)
    fNoKW = RS.Fields("NoKW").Value
    fTanggalKW = cTanggal(RS.Fields("TanggalKW").Value)
    fMataUang = RS.Fields("MataUang").Value
    fKetNR = RS.Fields("Ket").Value & ""
    UpdatefQuick
    a = "select t_NRDetail~.NoSJ,Jenis,KodeBarang,Warna,NoWarna,Tube,Grade, SatBesar, SatKecil, ReturBox,ReturKg," & _
        "t_NRDetail~.Harga,DiscKg,Discount,t_NRDetail~.IdStock from t_NRDetail~ left join m_stock~ on m_stock~.IdStock=t_NRDetail~.IdStock where NoNR='" & fNo & "'"
    query a
    TDBGridClear TDBGrid1
    If RS.RecordCount = 0 Then Exit Sub
    x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    TDBDropDown1.Columns("HARGA").NumberFormat = "Standard"
    TDBDropDown1.Columns("HARGA").Alignment = dbgRight
    a = "select t_SPPDetail~.NoSJ,Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, Harga, t_SPPDetail~.JumlahKg, m_stock~.IdStock from t_SPPDetail" & pTipe & " left join m_stock" & pTipe & " on t_SPPDetail" & pTipe & ".IdStock=m_stock~.IdStock where NoKWDetail='" & fNoKW & "'"
    query a
    z.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
    z.DeleteRows 0
    If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
    TDBDropDown1.Rebind
End Sub

Private Sub Form_Load()
    If pUpdateHargaSC = 0 And UCase(pUsr) <> "EULIS" Then fPrint2.Visible = False Else fPrint2.Visible = True
    iNoSJ = 0
    iJenis = 1
    iKode = 2
    iWarna = 3
    iNoWarna = 4
    iTube = 5
    iGrade = 6
    iSatBesar = 7
    iSatKecil = 8
    iJumlah1 = 9
    iJumlah2 = 10
    iHarga = 11
    iDiscKG = 12
    iDisc = 13
    iIdStock = 14
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    col1(iJumlah1).Tag = "Integer"
    col1(iJumlah2).Tag = "Decimal"
    col1(iHarga).Tag = "Decimal"
    col1(iDiscKG).Tag = "Decimal"
    col1(iDisc).Tag = "Decimal"
    TDBGridLoad TDBGrid1
    TDBGridSetVisible TDBGrid1, iIdStock
    col1(iNoSJ).AutoDropDown = True
    col1(iNoSJ).DropDown = TDBDropDown1
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
    TDBGrid1.HeadLines = 2
End Sub

Sub LoadMe(ByVal tNo As String)
    Form_Load
    Set TDBDropDown1.Array = z
    fCustomer.Enabled = False
    fNoKW.Enabled = False
    fTanggalKW.Enabled = False
    a = "select top 1 Nama, Alamat,NoKW, TanggalKW, MataUang from t_SPP" & pTipe & " left join m_customer on t_SPP" & pTipe & ".Kode=m_customer.Kode where NoKW='" & tNo & "' order by Nama"
    query a
    If RS.RecordCount = 0 Then Exit Sub
    fCustomer = RS.Fields(0).Value
    fAlamat = RS.Fields(1).Value
    fNoKW = RS.Fields(2).Value
    fTanggalKW = cTanggal(RS.Fields(3).Value)
    fMataUang = RS.Fields("MataUang").Value
    fNo_LostFocus
    TDBDropDown1.Columns("HARGA").NumberFormat = "Standard"
    TDBDropDown1.Columns("HARGA").Alignment = dbgRight
    a = "select t_SPPDetail~.NoSJ,Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, Harga, t_SPPDetail~.JumlahKg, m_stock~.IdStock from t_SPPDetail" & pTipe & " left join m_stock" & pTipe & " on t_SPPDetail" & pTipe & ".IdStock=m_stock" & pTipe & ".IdStock where NoKWDetail='" & tNo & "'"
    query a
    z.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
    z.DeleteRows 0
    If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
    TDBDropDown1.Rebind
    UpdatefQuick
    fNew_Click
    fTanggal = pServerDate
End Sub

Private Sub fDelete_Click()
On Error GoTo err
    BeginTransaction
    a = "delete from t_NR" & pTipe & " where NoNR='" & fNo & "'"
    ExecMe a
    a = "delete from t_NRDetail" & pTipe & " where NoNR='" & fNo & "'"
    ExecMe a
    'For i = 0 To x.UpperBound(1)
    '    a = "update m_stock" & pTipe & " set JumlahBox=JumlahBox-" & x(i, iJumlah1) & _
    '        ",JumlahKG=JumlahKG-" & cNum(x(i, iJumlah2)) & " where IdStock=" & x(i, iIdStock)
    '    ExecMe a
    'Next
    CommitTransaction
    MsgBox "SUKSES"
    UpdatefQuick
    GoEvent "NEW"
    'SendData "1HAPUS NR NO: " & fNo & Chr(8)
    DoEvents
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub ClearScreen()
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
    fNo = ""
    fKetNR = ""
    fTanggal = pServerDate
    fNo_LostFocus
    GoEvent "NEW"
    a = "select t_SPPDetail" & pTipe & ".NoSJ,Jenis, KodeBarang, Warna, NoWarna,Tube, Grade, SatBesar, SatKecil, Harga,t_SPPDetail" & pTipe & ".JumlahKg, m_stock" & pTipe & ".IdStock from t_SPPDetail" & pTipe & " left join m_stock" & pTipe & " on t_SPPDetail" & pTipe & ".IdStock=m_stock" & pTipe & ".IdStock where NoKWDetail='" & fNoKW & "'"
    query a
    z.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
    z.DeleteRows 0
    If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
    TDBDropDown1.Rebind
End Sub

Private Sub fNew_Click()
    ClearScreen
End Sub

Private Sub fNo_LostFocus()
    BuatNomor fNo, fTanggal, pNomorNR, fQuick, "select max(NoNR) from t_NR" & pTipe & " where TanggalNR>" & pAddNoLong
End Sub

Private Sub fPrint2_Click()
    FormPreview.LoadMe Me, "NR", fNo & "@PrintNamaPenerima"
End Sub

Private Sub fPrint3_Click()
    FormPreview.LoadMe Me, "NR", fNo & "@HideCompany"
End Sub
Private Sub fPrint4_Click()
    FormPreview.LoadMe Me, "NR", fNo & "@PrintNamaPenerima@HideCompany"
End Sub

Private Sub fQuick_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fQuick_Validate False
End Sub

Private Sub fQuick_Validate(Cancel As Boolean)
On Error Resume Next
    'If fDelete.Enabled Then
    GetResult fQuick
End Sub


Private Sub fSave_Click()
On Error GoTo err
    BeginTransaction
    fNo.Validate
    fTanggal.Validate
    fNoKW.Validate
    TDBGrid1.Update
Dim tot As Double
    tot = 0
    For i = 0 To x.UpperBound(1)
        tot = tot + x(i, iJumlah2) * x(i, iHarga) + x(i, iDiscKG) * x(i, iDisc)
    Next
    a = "select Kode from m_customer where IsActive=1 and Nama='" & fCustomer & "'"
    query a
    Kode = RS.Fields(0).Value
    a = "insert into t_NR" & pTipe & "(NoNR,TanggalNR,NoKW,Ket,KodeCustomer,Total, TanggalKW, MataUang, ShortNR) values('" & _
        fNo & _
        "'," & cD(fTanggal) & _
        ",'" & fNoKW & _
        "','" & fKetNR & _
        "'," & Kode & _
        "," & cNum(tot) & _
        "," & cD(fTanggalKW) & _
        ",'" & fMataUang & _
        "','" & Left(fNo, 5) & "/" & Right(fNo, 2) & "')"
    If ExecMe(a) = 0 Then GoTo err
    Dim cIdStock As Long
    For i = 0 To x.UpperBound(1)
        a = "select top 1 IdStock from m_Stock~ where Jenis='" & x(i, iJenis) & "' and KodeBarang='" & x(i, iKode) & "' and Warna='" & x(i, iWarna) & "' and NoWarna='" & x(i, iNoWarna) & "' and Tube='" & x(i, iTube) & "' and Grade='" & x(i, iGrade) & "' and SatBesar='" & x(i, iSatBesar) & "'"
        query a
        If RS.RecordCount = 0 Then
            MsgBox "Stock tidak Terdaftar"
            GoTo err
        End If
        cIdStock = RS.Fields(0).Value
        a = "insert into t_NRDetail~(NoNR,KodeCustomerDetail,TanggalNRDetail,IdNR,IdStock,KetDetail,ReturBox,ReturKg,Harga,DiscKg,Discount,NoSJ,NoKWDetail, TanggalKWDetail, MataUangDetail) values('" & _
            fNo & _
            "'," & Kode & _
            "," & cD(fTanggal) & _
            "," & i & _
            "," & cIdStock & _
            ",'" & fKetNR & _
            "'," & cNum(x(i, iJumlah1)) & _
            "," & cNum(x(i, iJumlah2)) & _
            "," & cNum(x(i, iHarga)) & _
            "," & cNum(x(i, iDiscKG)) & _
            "," & cNum(x(i, iDisc)) & _
            ",'" & x(i, iNoSJ) & _
            "','" & fNoKW & _
            "'," & cD(fTanggalKW) & _
            ",'" & fMataUang & "')"
        If ExecMe(a) = 0 Then GoTo err
        'a = "update m_stock" & pTipe & " set JumlahBox=JumlahBox+" & cNum(x(i, iJumlah1)) & _
        '    ",JumlahKG=JumlahKG+" & cNum(x(i, iJumlah2)) & " where IdStock=" & x(i, iIdStock)
        'If ExecMe(a) = 0 Then GoTo err
    Next
    CommitTransaction
    MsgBox "SUKSES"
    UpdatefQuick
    'SendData "1BUAT NR NO: " & fNo & Chr(8)
    DoEvents
    GetResult fNo
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fTanggal_LostFocus()
    fNo_LostFocus
End Sub

Private Sub fTotal_Click()
    TDBGrid1.Update
    n = 0
    For i = 0 To x.UpperBound(1)
        n = n + x(i, iJumlah2) * x(i, iHarga) + x(i, iDisc) * x(i, iDiscKG)
    Next
    MsgBox cDecimal(n)
End Sub

Private Sub TDBDropDown1_Paint()
On Error Resume Next
    TDBGrid1.SelLength = Len(TDBGrid1.Text)
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 0 Then TDBDropDown1_DropDownClose
End Sub

Private Sub TDBDropDown1_DropDownClose()
    col1(iNoSJ).Value = TDBDropDown1.Columns("No SJ").Value
    col1(iJenis).Value = TDBDropDown1.Columns("JENIS").Value
    col1(iKode).Value = TDBDropDown1.Columns("KODE").Value
    col1(iWarna).Value = TDBDropDown1.Columns("WARNA").Value
    col1(iNoWarna).Value = TDBDropDown1.Columns("NO WARNA").Value
    col1(iTube).Value = TDBDropDown1.Columns("TUBE").Value
    col1(iGrade).Value = TDBDropDown1.Columns("GRADE").Value
    col1(iSatBesar).Value = TDBDropDown1.Columns("SatBesar").Value
    col1(iSatKecil).Value = TDBDropDown1.Columns("SatKecil").Value
    col1(iHarga).Value = TDBDropDown1.Columns("HARGA").Value
    col1(iIdStock).Value = TDBDropDown1.Columns("IdStock").Value
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Begin VB.Form FormWaste 
   BackColor       =   &H00FFC0C0&
   Caption         =   "WASTE"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Tag             =   "19"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fLaporan 
      Caption         =   "LAPORAN"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton fKW 
      Caption         =   "KW"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton fSJ 
      Caption         =   "SJ"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton fSPP 
      Caption         =   "SPP"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   2415
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama Barang"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "IdStock"
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
      Columns(4).Caption=   "Harga"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4207"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4128"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=10"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1270"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1191"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=79399504"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1429"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1349"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=78250968"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1349"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1270"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=1905"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1826"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2415
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4260
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
      Columns(1).Caption=   "Telepon"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "KodeCustomer"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4207"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4128"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=10"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1429"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1349"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=78250968"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin UsrText.IText fCustomer 
      Height          =   270
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.CheckBox fLunas 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lunas"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
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
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
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
      Columns(2).Caption=   "QTYTag"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "cKeyIdStock"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "cKodeCustomer"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Print?"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "NoWaste"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Tanggal"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Nama Customer"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Id Stock"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Nama Barang"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "QTY"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Satuan"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Curr"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Harga"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Total"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Cara Bayar"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Pelunasan"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Tanggal Lunas"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Keterangan"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   20
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=20"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=926"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=847"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1376"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1296"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1482"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1402"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(8)._MinWidth=149"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=1217"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=1138"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(9)._MinWidth=149"
      Splits(0)._ColumnProps(43)=   "Column(10).Width=3466"
      Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=3387"
      Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(47)=   "Column(10)._MinWidth=54391100"
      Splits(0)._ColumnProps(48)=   "Column(11).Width=1032"
      Splits(0)._ColumnProps(49)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(11)._WidthInPix=953"
      Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(52)=   "Column(12).Width=1296"
      Splits(0)._ColumnProps(53)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(12)._WidthInPix=1217"
      Splits(0)._ColumnProps(55)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(56)=   "Column(13).Width=979"
      Splits(0)._ColumnProps(57)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(13)._WidthInPix=900"
      Splits(0)._ColumnProps(59)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(60)=   "Column(13)._MinWidth=64"
      Splits(0)._ColumnProps(61)=   "Column(14).Width=2117"
      Splits(0)._ColumnProps(62)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(14)._WidthInPix=2037"
      Splits(0)._ColumnProps(64)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(65)=   "Column(15).Width=2328"
      Splits(0)._ColumnProps(66)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(15)._WidthInPix=2249"
      Splits(0)._ColumnProps(68)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(69)=   "Column(16).Width=1720"
      Splits(0)._ColumnProps(70)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(16)._WidthInPix=1640"
      Splits(0)._ColumnProps(72)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(73)=   "Column(16)._MinWidth=1"
      Splits(0)._ColumnProps(74)=   "Column(17).Width=1773"
      Splits(0)._ColumnProps(75)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(17)._WidthInPix=1693"
      Splits(0)._ColumnProps(77)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(78)=   "Column(18).Width=2037"
      Splits(0)._ColumnProps(79)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(18)._WidthInPix=1958"
      Splits(0)._ColumnProps(81)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(82)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(83)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(85)=   "Column(19).Order=20"
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
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=106,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=110,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=114,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=118,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=115,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=116,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=117,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=98,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=46,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=50,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=66,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=70,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=74,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=71,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=72,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=73,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=78,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=75,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=76,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=77,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(16).Style:id=82,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(16).HeadingStyle:id=79,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(16).FooterStyle:id=80,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(16).EditorStyle:id=81,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(17).Style:id=86,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(18).Style:id=90,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(18).HeadingStyle:id=87,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(18).FooterStyle:id=88,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(18).EditorStyle:id=89,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(19).Style:id=94,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(19).HeadingStyle:id=91,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(19).FooterStyle:id=92,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(19).EditorStyle:id=93,.parent=17"
      _StyleDefs(110) =   "Named:id=33:Normal"
      _StyleDefs(111) =   ":id=33,.parent=0"
      _StyleDefs(112) =   "Named:id=34:Heading"
      _StyleDefs(113) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(114) =   ":id=34,.wraptext=-1"
      _StyleDefs(115) =   "Named:id=35:Footing"
      _StyleDefs(116) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(117) =   "Named:id=36:Selected"
      _StyleDefs(118) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(119) =   "Named:id=37:Caption"
      _StyleDefs(120) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(121) =   "Named:id=38:HighlightRow"
      _StyleDefs(122) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(123) =   "Named:id=39:EvenRow"
      _StyleDefs(124) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(125) =   "Named:id=40:OddRow"
      _StyleDefs(126) =   ":id=40,.parent=33"
      _StyleDefs(127) =   "Named:id=41:RecordSelector"
      _StyleDefs(128) =   ":id=41,.parent=34"
      _StyleDefs(129) =   "Named:id=42:FilterBar"
      _StyleDefs(130) =   ":id=42,.parent=33"
   End
   Begin UsrText.IText fAwal 
      Height          =   270
      Left            =   3000
      TabIndex        =   7
      Top             =   840
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
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   3960
      TabIndex        =   8
      Top             =   840
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
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Click Update untuk Update Data Waste"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FormWaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid80.Columns
Dim coly1 As TrueOleDBGrid80.Columns
Dim coly2 As TrueOleDBGrid80.Columns
Dim y1 As New XArrayDB
Dim y2 As New XArrayDB
Dim x As New XArrayDB
Dim LCustomer As Boolean
Dim LNamaBarang As Boolean

Private Sub fAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub fCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub fLaporan_Click()
Dim y As New XArrayDB
    b = col1("NoWaste").ColIndex
    y.ReDim 0, x.UpperBound(1), 0, x.UpperBound(2) - b
    For j = 0 To x.UpperBound(1)
        For i = b To col1.Count - 1
            y(j, i - b) = x(j, i)
        Next
    Next
    'FormPreview.SetData Me, "LaporanWaste", IIf(fLunas.Value = 1, "LUNAS", "") & " " & fCustomer & " " & fAwal & " - " & fAkhir, res
    FormPreview.LoadFromData Me, "LaporanWaste", y, IIf(fLunas.Value = 1, "LUNAS", "") & " " & fCustomer & " " & fAwal & " - " & fAkhir
End Sub

Private Sub fLunas_Click()
    DoQuery
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LCustomer = False
    LNamaBarang = False
End Sub

Private Function MyFilter() As String
    If fLunas Then
        MyFilter = MyFilter & " and Pelunasan>=Total"
    End If
    If Trim(fCustomer) <> "" Then MyFilter = MyFilter & " and Nama like '%" & fCustomer & "%'"
    If cD(fAwal) <> "A" Then MyFilter = MyFilter & " and Tanggal>=" & cD(fAwal)
    If cD(fAkhir) <> "A" Then MyFilter = MyFilter & " and Tanggal<=" & cD(fAkhir)
    MyFilter = " where " & Mid(MyFilter, 6)
End Function

Private Sub DoQuery()
    a = "select 0, NoWaste, QTY, t_Waste.IdStock, KodeCustomer, Printed-1, NoWaste, Tanggal, Nama, t_Waste.IdStock, m_StockBeli.NamaBarang, QTY, Satuan, t_Waste.MataUang, Harga, Total, CaraBayar, Pelunasan, TanggalLunas, Keterangan from (t_Waste left join m_StockBeli on t_Waste.IdStock=m_StockBeli.IdStock) left join m_Customer on m_Customer.Kode=t_Waste.KodeCustomer " & MyFilter & " order by NoWaste"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then
        x.LoadRows RS.GetRows
        For i = 0 To x.UpperBound(1)
            x(i, col1("NoWaste").ColIndex) = CLng(Left(x(i, col1("NoWaste").ColIndex), 5))
            x(i, col1("Tanggal").ColIndex) = cTanggal2(x(i, col1("Tanggal").ColIndex))
            x(i, col1("Tanggal Lunas").ColIndex) = cTanggal2(x(i, col1("Tanggal Lunas").ColIndex))
        Next
    End If
    TDBGrid1.Rebind
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    Set coly1 = TDBDropDown1.Columns
    Set coly2 = TDBDropDown2.Columns
    Set TDBGrid1.Array = x
    Set TDBDropDown1.Array = y1
    Set TDBDropDown2.Array = y2
    col1("updated").Visible = False
    col1("cKey").Visible = False
    col1("QTYTag").Visible = False
    col1("cKeyIdStock").Visible = False
    col1("cKodeCustomer").Visible = False
    coly1("KodeCustomer").Visible = False
    TDBGrid1.FetchRowStyle = True
    col1("Tanggal").NumberFormat = "Edit Mask"
    col1("Tanggal").EditMask = "##/##/##"
    col1("Tanggal Lunas").NumberFormat = "Edit Mask"
    col1("Tanggal Lunas").EditMask = "##/##/##"
    col1("Harga").NumberFormat = "Standard"
    col1("Total").NumberFormat = "Standard"
    col1("Pelunasan").NumberFormat = "Standard"
    col1("Total").NumberFormat = "Standard"
    col1("QTY").Alignment = dbgRight
    col1("Harga").Alignment = dbgRight
    col1("Nama Customer").AutoDropDown = True
    col1("Nama Customer").DropDown = "TDBDropDown1"
    col1("Nama Barang").AutoDropDown = True
    col1("Nama Barang").DropDown = "TDBDropDown2"
    col1("Pelunasan").Alignment = dbgRight
    col1("Total").Alignment = dbgRight
    col1("Total").Locked = True
    col1("Print?").ValueItems.Presentation = dbgCheckBox
    col1("Print?").Alignment = dbgCenter
    coly2("QTY").Alignment = dbgRight
    coly2("Harga").Alignment = dbgRight
    coly2("Harga").NumberFormat = "Standard"
    col1("Id Stock").Locked = True
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
    fAkhir = pServerDate
    fAwal = cTanggal((cD(fAkhir) \ 100) * 100 + 1)
    DoQuery
End Sub

Private Sub Form_Resize()
On Error Resume Next
    fKet.Width = ScaleWidth - fKet.Left - 100
    TDBGrid1.Width = ScaleWidth
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fSJ_Click()
    b = GetParam
    If b = "" Then Exit Sub
    a = "select Nama as NamaPenerimaDetail, Alamat as AlamatPenerimaDetail, NoWaste as NoSJ, Tanggal as TanggalDetail, NoWaste as NoSPP, Tanggal as TanggalSPPDetail , NamaBarang, Satuan as SatBesar, QTY as JumlahBox, Satuan as SatKecil, QTY as JumlahKG, Keterangan as KeteranganSPPDetail, '' as NoKendaraanDetail, '' as NamaAngkutanDetail, '' as NamaSopirDetail from (t_Waste left join m_StockBeli on m_StockBeli.IdStock=t_Waste.IdStock) left join m_Customer on t_Waste.KodeCustomer=m_Customer.Kode where t_Waste.NoWaste in" & b
    FormPreview.LoadMe Me, "SJ", "", a
End Sub

Private Sub fKW_Click()
    b = GetParam
    If b = "" Then Exit Sub
    c = Mid(b, 2)
    c = Split(c, ",")
    For i = 0 To UBound(c)
        nokw = nokw & "." & CLng(Left(Mid(c(i), 2), 5))
    Next
    nokw = Mid(nokw, 2) & "/" & Left(Right(b, 4), 2)
    a = "select '" & nokw & "' as NoKWDetail, Tanggal as TanggalKWDetail, t_Waste.MataUang as MataUangDetail, Nama, Alamat, NoWaste as NoSJ, Tanggal as TanggalDetail, NamaBarang as JenisBarang, QTY as JumlahKG, Satuan as SatKecil, Harga, Total from (t_Waste left join m_StockBeli on m_StockBeli.IdStock=t_Waste.IdStock) left join m_Customer on m_Customer.Kode=t_Waste.KodeCustomer where NoWaste in" & b & " order by NoWaste"
    FormPreview.LoadMe Me, "KW", "", a
End Sub

Private Function GetParam() As String
    TDBGrid1.Update
    b = ""
    For i = 0 To x.UpperBound(1)
        If x(i, col1("Print?").ColIndex) <> 0 And x(i, col1("Nama Customer").ColIndex) = col1("Nama Customer").Value And x(i, col1("Tanggal").ColIndex) = col1("Tanggal").Value Then
            b = b & "','" & zerofill(x(i, col1("NoWaste").ColIndex), 5) & "/" & Right(x(i, col1("Tanggal").ColIndex), 2)
        End If
    Next
    If b = "" Then
        MsgBox "Beri Tanda v untuk data yang ingin di Print"
        GetParam = ""
        Exit Function
    End If
    GetParam = "(" & Mid(b, 3) & "')"
End Function

Private Sub fSPP_Click()
    b = GetParam
    If b = "" Then Exit Sub
    a = "select Nama, Alamat, NoWaste as NoSPP, Tanggal as TanggalSPPDetail, Nama as NamaPenerimaDetail, Alamat as AlamatPenerimaDetail, Tanggal as TanggalKirimDetail, Keterangan as KeteranganSPPHeader, NamaBarang, Satuan as SatBesar, QTY as JumlahBox from (t_Waste left join m_Customer on m_Customer.Kode=t_Waste.KodeCustomer) left join m_StockBeli on m_StockBeli.IdStock=t_Waste.IdStock where NoWaste in" & b
    FormPreview.LoadMe Me, "SPP", "", a, "update t_Waste set Printed=1 where NoWaste in" & b
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) = "1" Then
            If x(i, col1("cKey").ColIndex) = "" Then
                a = "insert into t_Waste(NoWaste, Tanggal, KodeCustomer, IdStock, QTY, MataUang, Harga, Total, CaraBayar, Pelunasan, TanggalLunas, Keterangan) values('" & _
                    zerofill(x(i, col1("NoWaste").ColIndex), 5) & Right(x(i, col1("Tanggal").ColIndex), 3) & _
                    "'," & cD(x(i, col1("Tanggal").ColIndex)) & _
                    ",'" & esc(x(i, col1("cKodeCustomer").ColIndex)) & _
                    "'," & x(i, col1("Id Stock").ColIndex) & _
                    "," & cNum(x(i, col1("QTY").ColIndex)) & _
                    ",'" & esc(x(i, col1("Curr").ColIndex)) & _
                    "'," & cNum(x(i, col1("Harga").ColIndex)) & _
                    "," & cNum(x(i, col1("Total").ColIndex)) & _
                    ",'" & esc(x(i, col1("Cara Bayar").ColIndex)) & _
                    "'," & cNum(x(i, col1("Pelunasan").ColIndex)) & _
                    "," & cD(x(i, col1("Tanggal Lunas").ColIndex)) & _
                    ",'" & esc(x(i, col1("Keterangan").ColIndex)) & _
                    "')"
                If ExecMe(a) = 0 Then GoTo err
            Else
                a = "update t_Waste set NoWaste='" & zerofill(x(i, col1("NoWaste").ColIndex), 5) & pAddNo & _
                    "', Tanggal=" & cD(x(i, col1("Tanggal").ColIndex)) & _
                    ", KodeCustomer=" & cNum(x(i, col1("cKodeCustomer").ColIndex)) & _
                    ", IdStock=" & x(i, col1("Id Stock").ColIndex) & _
                    ", QTY=" & cNum(x(i, col1("QTY").ColIndex)) & _
                    ", MataUang='" & esc(x(i, col1("Curr").ColIndex)) & _
                    "', Harga=" & cNum(x(i, col1("Harga").ColIndex)) & _
                    ", Total=" & cNum(x(i, col1("Total").ColIndex)) & _
                    ", CaraBayar='" & esc(x(i, col1("Cara Bayar").ColIndex)) & _
                    "', Pelunasan=" & cNum(x(i, col1("Pelunasan").ColIndex)) & _
                    ", TanggalLunas=" & cD(x(i, col1("Tanggal Lunas").ColIndex)) & _
                    ", Keterangan='" & esc(x(i, col1("Keterangan").ColIndex)) & _
                    "' where NoWaste='" & x(i, col1("cKey").ColIndex) & "'"
                If ExecMe(a) = 0 Then GoTo err
                a = "update m_StockWaste set Jumlah=Jumlah+" & cNum(x(i, col1("QTYTag").ColIndex)) & " where IdStock=" & x(i, col1("cKeyIdStock").ColIndex)
                If ExecMe(a) = 0 Then GoTo err
            End If
            a = "update m_StockWaste set Jumlah=Jumlah-" & cNum(x(i, col1("QTY").ColIndex)) & " where IdStock=" & x(i, col1("Id Stock").ColIndex)
            If ExecMe(a) = 0 Then GoTo err
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoQuery
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBDropDown1_DropDownClose()
On Error Resume Next
    If TDBDropDown1.Bookmark = "" Or TDBDropDown1.Bookmark < 0 Then
        col1("Nama Customer").Value = ""
        col1("cKodeCustomer").Value = 0
    Else
        col1("Nama Customer").Value = coly1("Nama Customer").Value
        col1("cKodeCustomer").Value = coly1("KodeCustomer").Value
    End If
End Sub

Private Sub TDBDropDown2_DropDownClose()
On Error Resume Next
    If TDBDropDown2.Bookmark = "" Or TDBDropDown2.Bookmark < 0 Then
        col1("Nama Barang").Value = ""
        col1("QTY").Value = ""
        col1("Satuan").Value = ""
        col1("Harga").Value = ""
        col1("Total").Value = ""
        col1("Id Stock").Value = ""
    Else
        col1("Nama Barang").Value = coly2("Nama Barang").Value
        col1("QTY").Value = coly2("QTY").Value
        col1("Satuan").Value = coly2("Satuan").Value
        col1("Harga").Value = coly2("Harga").Value
        col1("Total").Value = coly2("Harga").Value * coly2("QTY").Value
        col1("Id Stock").Value = coly2("IdStock").Value
    End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
    If col1(ColIndex).Caption <> "Print?" Then col1("updated").Value = "1"
    a = col1(ColIndex).Caption
    If a = "QTY" Or a = "Harga" Then
        col1("Total").Value = col1("QTY").Value * col1("Harga").Value
    End If
End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
On Error Resume Next
    a = col1(ColIndex).Caption
    If a = "Nama Customer" Then
        If Not LCustomer Then
            a = "select Nama, Telepon, Kode from m_Customer where IsActive=1 order by Nama"
            query a
            y1.ReDim 0, 0, 0, coly1.Count - 1
            y1.DeleteRows 0
            If RS.RecordCount > 0 Then y1.LoadRows RS.GetRows
            TDBDropDown1.Rebind
            LCustomer = True
        End If
    ElseIf a = "Nama Barang" Then
        If Not LNamaBarang Then
            a = "select NamaBarang, m_StockWaste.IdStock, m_StockWaste.Jumlah, Satuan, Harga from m_StockWaste inner join m_StockBeli on m_StockWaste.IdStock=m_StockBeli.IdStock order by NamaBarang"
            query a
            y2.ReDim 0, 0, 0, coly2.Count - 1
            y2.DeleteRows 0
            If RS.RecordCount > 0 Then y2.LoadRows RS.GetRows
            TDBDropDown2.Rebind
            LNamaBarang = True
        End If
    End If
End Sub

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)
    b = MsgBox("Yakin Hapus?", vbYesNo)
    If b = vbNo Then
        Cancel = True
        Exit Sub
    End If
    BeginTransaction
    If col1("cKey").Value <> "" Then
        a = "delete from t_Waste where NoWaste='" & esc(col1("cKey").Value) & "'"
        If ExecMe(a) = 0 Then GoTo err
        a = "update m_StockBeli set JumlahProses=JumlahProses+" & cNum(col1("QTYTag").Value) & " where IdStock=" & col1("cKeyIdStock").Value
        If ExecMe(a) = 0 Then Exit Sub
    End If
    CommitTransaction
    Exit Sub
err:
    RollBackTransaction
    Cancel = True
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("updated").ColIndex) = "1" Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    fKet = col1("Nama Customer").Value & ": " & col1("Nama Barang").Value
End Sub


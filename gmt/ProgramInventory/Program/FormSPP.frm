VERSION 5.00
Object = "{8AAEAB20-E970-42F3-9E69-BC54C54CC273}#4.0#0"; "usrcombo.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{BD09B73E-A5EF-4CAB-A002-921A8335B40E}#1.0#0"; "usrtruecombo.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormSPP 
   BackColor       =   &H00FFC0C0&
   Caption         =   "SURAT PERINTAH PENGIRIMAN"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   1995
      Left            =   6360
      TabIndex        =   44
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3519
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Satuan"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=416"
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
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=228,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(34)  =   "Named:id=33:Normal"
      _StyleDefs(35)  =   ":id=33,.parent=0"
      _StyleDefs(36)  =   "Named:id=34:Heading"
      _StyleDefs(37)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(38)  =   ":id=34,.wraptext=-1"
      _StyleDefs(39)  =   "Named:id=35:Footing"
      _StyleDefs(40)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(41)  =   "Named:id=36:Selected"
      _StyleDefs(42)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(43)  =   "Named:id=37:Caption"
      _StyleDefs(44)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(45)  =   "Named:id=38:HighlightRow"
      _StyleDefs(46)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=39:EvenRow"
      _StyleDefs(48)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(49)  =   "Named:id=40:OddRow"
      _StyleDefs(50)  =   ":id=40,.parent=33"
      _StyleDefs(51)  =   "Named:id=41:RecordSelector"
      _StyleDefs(52)  =   ":id=41,.parent=34"
      _StyleDefs(53)  =   "Named:id=42:FilterBar"
      _StyleDefs(54)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      CausesValidation=   0   'False
      Height          =   3375
      Left            =   360
      TabIndex        =   32
      Top             =   4680
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5953
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No SC"
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
      Columns(7).Caption=   "Sisa"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Harga"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Keterangan"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "IdSC"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "DP"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3043"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2963"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1826"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1746"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=2328"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2249"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1270"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1191"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1852"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1773"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1508"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1429"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=1138"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1058"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=1931"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1852"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=1773"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1693"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(56)=   "Column(9).Width=5212"
      Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=5133"
      Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(62)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(67)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(68)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(69)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(73)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(74)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(75)=   "Column(11).Order=12"
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
      DeadAreaBackColor=   12632256
      ValueTranslate  =   0   'False
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=81,.parent=2,.namedParent=83"
      _StyleDefs(23)  =   "FilterBarStyle:id=84,.parent=1,.namedParent=86"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=82,.parent=81"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=85,.parent=84"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=40,.parent=11"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=44,.parent=11"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=41,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=42,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=43,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=48,.parent=11"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=45,.parent=12"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=46,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=47,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=52,.parent=11"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=49,.parent=12"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=50,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=51,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=56,.parent=11"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=53,.parent=12"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=54,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=55,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=60,.parent=11"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=57,.parent=12"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=58,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=59,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=64,.parent=11"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=61,.parent=12"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=62,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=63,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=68,.parent=11"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=65,.parent=12"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=66,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=67,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=80,.parent=11"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=77,.parent=12"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=78,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=79,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=72,.parent=11"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=69,.parent=12"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=70,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=71,.parent=15"
      _StyleDefs(84)  =   "Named:id=29:Normal"
      _StyleDefs(85)  =   ":id=29,.parent=0"
      _StyleDefs(86)  =   "Named:id=30:Heading"
      _StyleDefs(87)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=30,.wraptext=-1"
      _StyleDefs(89)  =   "Named:id=31:Footing"
      _StyleDefs(90)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=32:Selected"
      _StyleDefs(92)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=33:Caption"
      _StyleDefs(94)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(95)  =   "Named:id=34:HighlightRow"
      _StyleDefs(96)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(97)  =   "Named:id=35:EvenRow"
      _StyleDefs(98)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(99)  =   "Named:id=36:OddRow"
      _StyleDefs(100) =   ":id=36,.parent=29"
      _StyleDefs(101) =   "Named:id=83:RecordSelector"
      _StyleDefs(102) =   ":id=83,.parent=30"
      _StyleDefs(103) =   "Named:id=86:FilterBar"
      _StyleDefs(104) =   ":id=86,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   2990
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No SC"
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
      Columns(7).Caption=   "Satuan"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Harga"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Jumlah"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Keterangan"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "IdSC"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "DP"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "IdDet"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3440"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3360"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1244"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1164"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1085"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1005"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1482"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1402"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(4)._MinWidth=1312901971"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1191"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1111"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=979"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=900"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=1535"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=1455"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=1879"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=1799"
      Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=1191"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=1111"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(10).Width=2831"
      Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=2752"
      Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(46)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(50)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(51)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(53)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(54)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(58)=   "Column(13)._MinWidth=-1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
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
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
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
   Begin VB.CommandButton fMasterStock 
      Caption         =   "MASTER STOCK"
      Height          =   375
      Left            =   7080
      TabIndex        =   41
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton fHistoris 
      Caption         =   "HISTORIS"
      Height          =   375
      Left            =   9840
      TabIndex        =   40
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox fSetuju 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SETUJU"
      Height          =   255
      Left            =   7800
      TabIndex        =   39
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox fAlamatPendek 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   15
      Top             =   2040
      Width           =   4935
   End
   Begin VB.CommandButton fStock 
      Caption         =   "STOCK"
      Height          =   375
      Left            =   8760
      TabIndex        =   36
      Top             =   120
      Width           =   975
   End
   Begin UsrText.IText fWaktuPembayaran 
      Height          =   270
      Left            =   8280
      TabIndex        =   6
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
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
   Begin VB.TextBox fAlamatPenerima 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   14
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fNew 
      Caption         =   "NEW"
      Height          =   375
      Left            =   9000
      TabIndex        =   17
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox fAlamat 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   6375
   End
   Begin VB.CommandButton fSave 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   9840
      TabIndex        =   10
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   10800
      TabIndex        =   12
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton fList 
      Caption         =   "LIST"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin UsrText.IText fQuick 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
   Begin UsrText.IText fKeterangan 
      Height          =   270
      Left            =   1800
      TabIndex        =   7
      Top             =   3720
      Width           =   4695
      _ExtentX        =   8281
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
   Begin UsrCombo.ICombo fMataUang 
      Height          =   315
      Left            =   6720
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Enabled         =   0   'False
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
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   9240
      TabIndex        =   4
      Top             =   2760
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
      Left            =   6720
      TabIndex        =   3
      Top             =   2760
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
   End
   Begin UsrTrueCombo.ITrueCombo fCustomer 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
      _ExtentX        =   11245
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
   Begin UsrText.IText fTelepon 
      Height          =   270
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   3135
      _ExtentX        =   5530
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
   Begin UsrText.IText fFax 
      Height          =   270
      Left            =   3360
      TabIndex        =   21
      Top             =   2640
      Width           =   3135
      _ExtentX        =   5530
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
   Begin UsrText.IText fContactPerson 
      Height          =   270
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   6375
      _ExtentX        =   11245
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
   Begin UsrText.IText fTanggalKirim 
      Height          =   270
      Left            =   10200
      TabIndex        =   5
      Top             =   2760
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
   Begin UsrTrueCombo.ITrueCombo fNamaPenerima 
      Height          =   285
      Left            =   6720
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
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
   Begin UsrText.IText fNoKW 
      Height          =   270
      Left            =   4800
      TabIndex        =   42
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Label fUpdatedBy 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2280
      TabIndex        =   45
      Top             =   660
      Width           =   4215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NO KW"
      Height          =   255
      Left            =   4080
      TabIndex        =   43
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL KIRIM"
      Height          =   255
      Left            =   10200
      TabIndex        =   38
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA PENERIMA"
      Height          =   255
      Left            =   6720
      TabIndex        =   37
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "HARI"
      Height          =   255
      Left            =   9240
      TabIndex        =   35
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "WAKTU PEMBAYARAN"
      Height          =   255
      Left            =   8280
      TabIndex        =   34
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT PENERIMA"
      Height          =   255
      Left            =   6720
      TabIndex        =   33
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NO SPP"
      Height          =   255
      Left            =   6720
      TabIndex        =   31
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
      Height          =   255
      Left            =   9240
      TabIndex        =   30
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "MATA UANG"
      Height          =   255
      Left            =   6720
      TabIndex        =   29
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN SPP"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT PERSON"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   1755
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      Height          =   255
      Left            =   3360
      TabIndex        =   26
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TELEPON"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA CUSTOMER"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "FormSPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fKode As Long
Dim m_mode As String
Dim m_status As Byte
Dim z As New XArrayDB
Dim zSat As New XArrayDB
Dim LCustomer As Boolean
Dim LPenerima As Boolean
Dim LDetail As Boolean
Dim LSat As Boolean
Dim col1 As TrueOleDBGrid80.Columns
Dim iNoSC As Integer
Dim iJenis As Integer
Dim iKode As Integer
Dim iWarna As Integer
Dim iNoWarna As Integer
Dim iTube As Integer
Dim iGrade As Integer
Dim iSatuan As Integer
Dim iHarga As Integer
Dim iJumlah As Integer
Dim iKeterangan As Integer
Dim iIdSC As Integer
Dim iDP As Integer
Dim iIdDet As Integer
Dim x As New XArrayDB

Private Sub fCustomer_Validate(Cancel As Boolean)
On Error Resume Next
    LPenerima = False
    If fCustomer = "" Then Exit Sub
    fCustomer = fCustomer.GetData("Nama Customer")
    fAlamat = fCustomer.GetData("Alamat")
    fMataUang = fCustomer.GetData("MataUang")
    fTelepon = fCustomer.GetData("Telepon")
    fFax = fCustomer.GetData("Fax")
    fContactPerson = fCustomer.GetData("ContactPerson")
    fKode = ""
    fKode = fCustomer.GetData("Kode")
    fWaktuPembayaran = fCustomer.GetData("WaktuPembayaran")
    fNamaPenerima = ""
    fAlamatPenerima = ""
    fAlamatPendek = ""
    fNamaPenerima_KeyDown 0, 0
End Sub

Private Sub fHistoris_Click()
    FormHistoryPenjualan.LoadMe fCustomer, col1(iJenis).value & " " & col1(iKode).value & " " & col1(iWarna).value & " " & col1(iNoWarna).value & " " & col1(iTube).value & " GRADE " & col1(iGrade).value, Me
End Sub

Private Sub fMasterStock_Click()
    FormMasterStock.Show
End Sub

Private Sub fNamaPenerima_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    fNamaPenerima.Cancel
    j = fNamaPenerima.ListCount - 1
    For i = 0 To j
        If StrComp(Left(fNamaPenerima.zz(i, "Nama Penerima"), Len(fNamaPenerima)), fNamaPenerima, vbTextCompare) = 0 Then
            fNamaPenerima.SetListIndex i
            Exit Sub
        End If
    Next
    fNamaPenerima.SetListIndex -1
End Sub


Private Sub GoEvent(ByVal tEvent As String)
    If tEvent = "NEW" Then
        m_mode = "NEW"
        tProp = 1
    ElseIf tEvent = "EDIT" Then
        m_mode = "EDIT"
        tProp = 3
    ElseIf tEvent = "SEE" Then
        m_mode = "SEE"
        tProp = 0
    End If
    v = IIf((tProp And 1) = 0, False, True)
        fCustomer.Enabled = v
        fNamaPenerima.Enabled = v
        fNo.Enabled = v
        fTanggal.Enabled = v
        fTanggalKirim.Enabled = v
        fWaktuPembayaran.Enabled = v
        fKeterangan.Enabled = v
        fSetuju.Enabled = v
        fSave.Enabled = v
        TDBGrid1.AllowUpdate = v
        
    v = IIf((tProp And 2) = 0, False, True)
        fDelete.Enabled = v
        fPrint.Enabled = v
    If pUpdateHargaSC Then fPrint.Enabled = tEvent <> "NEW"
End Sub

Private Sub ClearScreen()
    fUpdatedBy = ""
    fCustomer = ""
    fAlamat = ""
    fNamaPenerima = ""
    fAlamatPenerima = ""
    fAlamatPendek = ""
    fTelepon = ""
    fFax = ""
    fContactPerson = ""
    fKode = -1
    fNo = ""
    fNo.Tag = ""
    fNoKW = ""
    fTanggal = pServerDate
    fTanggalKirim = fTanggal
    fKeterangan = ""
    fWaktuPembayaran = ""
    fSetuju.value = 0
    x.ReDim 0, 0, 0, col1.count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
End Sub

Private Sub fAlamat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys Chr(9)
End Sub

Private Sub fAlamatPenerima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys Chr(9)
End Sub

Private Sub fCustomer_GotFocus()
    LDetail = False
End Sub

Private Sub fCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    If Not LCustomer Then
        fCustomer.SetHeader "Nama Customer@Alamat@MataUang@*Telepon@*Fax@*ContactPerson@*Kode@*WaktuPembayaran@*AlamatPendek"
        fCustomer.SetWidth "2500@4000@1000"
        Dim rs1() As Variant
        a = "select distinct Nama, Alamat, t_SC~.MataUang, Telepon, Fax, ContactPerson, m_Customer.Kode, t_SC~.WaktuPembayaran, AlamatPendek from t_SC~ left join m_customer on m_customer.Kode=t_SC~.Kode where Status=0 and Disetujui=1 order by Nama"
        query a
        If Not RS.EOF Then
            rs1 = RS.GetRows
            fCustomer.SetDB rs1
            LCustomer = True
            fCustomer.SetType "String@String@String@String@String@String"
        End If
    End If
    If KeyCode = 116 Then LCustomer = False
err:
End Sub

Private Sub fNamaPenerima_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    If LPenerima Then Exit Sub
    LPenerima = True
    fNamaPenerima.SetHeader "Nama Penerima@Alamat Penerima@Alamat Pendek"
    fNamaPenerima.SetWidth "2000@2000@1000"
    Dim rs1() As Variant
    If fCustomer.GetData("Kode") <> "" Then
        a = "select Nama,Alamat,AlamatPendek from m_Penerima where Kode=" & fCustomer.GetData("Kode") & " ORDER BY NoUrut"
        query a
    End If
    ReDim rs1(0, 0)
    v = RS.RecordCount > 0
    If v Then
        rs1 = RS.GetRows
        fNamaPenerima.SetDB rs1
        fNamaPenerima.SetType "String@String@String"
        fNamaPenerima.SetListIndex 0
    Else
        fNamaPenerima = fCustomer
        fAlamatPenerima = fCustomer.GetData("Alamat")
        fAlamatPendek = fCustomer.GetData("AlamatPendek")
    End If
    fNamaPenerima.Enabled = v And m_mode <> "SEE"
    Exit Sub
err:
    LPenerima = False
End Sub

Private Sub fNamaPenerima_LostFocus()
On Error Resume Next
    If fNamaPenerima.Enabled Then
        fNamaPenerima = fNamaPenerima.GetData("Nama Penerima")
        fAlamatPenerima = fNamaPenerima.GetData("Alamat Penerima")
        fAlamatPendek = fNamaPenerima.GetData("Alamat Pendek")
    End If
End Sub

Private Sub LoadGrid1()
    Set TDBDropDown1.Array = z
    z.ReDim 0, 0, 0, TDBDropDown1.Columns.count - 1
    z.DeleteRows 0
    a = "select t_SC~.NoSC, JenisBarang, KodeBarang, Warna, NoWarna, Tube, Grade, Jumlah-Terpakai as Sisa,Harga,t_SCDetail~.Keterangan,IdSC, DP from t_SCDetail~ left join t_SC~ on t_SC~.NoSC=t_SCDetail~.NoSC where Jumlah-Terpakai>0 and Kode=" & fKode & " and MataUang='" & esc(fMataUang) & "' and  status=0 and disetujui=1  ORDER BY TANGGALSC, T_SC~.NOSC"
    'a = "select t_SC~.NoSC, JenisBarang, KodeBarang, Warna, NoWarna, Tube, Grade, Jumlah-Terpakai as Sisa,Harga,t_SCDetail~.Keterangan,IdSC, DP from t_SCDetail~ left join t_SC~ on t_SC~.NoSC=t_SCDetail~.NoSC where Kode=" & fKode & " and MataUang='" & fMataUang & "' and  status=0 and disetujui=1  ORDER BY TANGGALSC, T_SC~.NOSC"
    query a
    If RS.RecordCount <> 0 Then z.LoadRows RS.GetRows
    TDBDropDown1.Rebind
End Sub

Private Sub fDelete_Click()
On Error GoTo err
    BeginTransaction
    a = "delete from t_SPP~ where NoSPP='" & esc(fNo.Tag) & "'"
    ExecMe a
    a = "delete from t_SPPDetail~ where NoSPP='" & esc(fNo.Tag) & "'"
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    For i = 0 To x.UpperBound(1)
        x(i, iIdDet) = ""
    Next
    TDBGrid1.Rebind
    GoEvent "NEW"
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fList_Click()
    FormList.Show , Me
    FormList.LoadMe "BELUM SETUJU@BELUM SURAT JALAN@SUDAH SURAT JALAN", _
        "select NoSPP,Nama,TanggalSPP,TanggalKirim,MataUang,KeteranganSPP from t_SPP~ left join m_Customer on t_SPP~.Kode=m_Customer.Kode where status=0@" & _
        "select NoSPP,Nama,TanggalSPP,TanggalKirim,MataUang,KeteranganSPP from t_SPP~ left join m_Customer on t_SPP~.Kode=m_Customer.Kode where status=1@" & _
        "select NoSPP,Nama,TanggalSPP,TanggalKirim,MataUang,KeteranganSPP from t_SPP~ left join m_Customer on t_SPP~.Kode=m_Customer.Kode where status>=2", _
        "No SPP@Nama", "NoSPP@Nama", "1000@3500", "String@String", _
        "No SPP@Nama Customer@Tanggal SPP@Tanggal Kirim@Mata Uang @Keterangan", _
        "2000@2500@1000@1000@700@2500", _
        "String@String@Date@Date@String@String", Me, " order by TanggalSPP\10000 desc, NoSPP desc"
End Sub

Private Sub fNew_Click()
On Error Resume Next
    ClearScreen
    fNo_LostFocus
    m_status = 0
    GoEvent "NEW"
    fCustomer.Tag = ""
    fCustomer.SetFocus
End Sub

Private Sub fNext_Click()
On Error Resume Next
    a = fQuick
    Mid(a, 1) = zerofill(Left(a, 5) + 1, 5)
    fQuick = a
    GetResult fQuick
End Sub

Private Sub fNo_LostFocus()
    BuatNomor fNo, fTanggal, pNomorSPP, fQuick, "select max(NoSPP) from t_SPP" & pTipe & " where TanggalSPP>" & pAddNoLong
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LCustomer = False
    LPenerima = False
    LDetail = False
    LSat = False
End Sub

Sub SetOtherRowData(ByVal tNo As Long)
On Error GoTo err
    If Not TDBGrid1.AllowUpdate Then Exit Sub
    a = "select Jenis,KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar from m_stock~ where IsActive=1 and IdStock=" & tNo
    query a
    col1(iJenis).value = RS.Fields("Jenis").value
    col1(iKode).value = RS.Fields("KodeBarang").value
    col1(iWarna).value = RS.Fields("Warna").value
    col1(iNoWarna).value = RS.Fields("NoWarna").value
    col1(iTube).value = RS.Fields("Tube").value
    col1(iGrade).value = RS.Fields("Grade").value
    col1(iSatuan).value = RS.Fields("SatBesar").value
    TDBGrid1.SetFocus
err:
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    fCustomer.ZOrder 0
    fNamaPenerima.ZOrder 0
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    iNoSC = 0
    iJenis = 1
    iKode = 2
    iWarna = 3
    iNoWarna = 4
    iTube = 5
    iGrade = 6
    iSatuan = 7
    iHarga = 8
    iJumlah = 9
    iKeterangan = 10
    iIdSC = 11
    iDP = 12
    iIdDet = 13
    col1(iHarga).Tag = "Decimal"
    col1(iJumlah).Tag = "Integer"
    TDBGridLoad TDBGrid1
    col1(iNoSC).DropDown = TDBDropDown1
    
    col1(iNoSC).AutoDropDown = True
    col1(iHarga).Locked = True
    TDBGridSetVisible TDBGrid1, "IdSC@DP@IdDet", False
    TDBDropDown1.Columns("HARGA").Alignment = dbgRight
    TDBDropDown1.Columns("HARGA").NumberFormat = "Standard"
    TDBDropDown1.Columns("SISA").Alignment = dbgRight
    TDBDropDown1.Columns("SISA").NumberFormat = "Standard"
    a = "select kode from m_matauang order by Kode"
    query a
    For i = 0 To RS.RecordCount - 1
        fMataUang.List(i) = RS.Fields(0).value
        RS.MoveNext
    Next
    TDBGrid1.Columns("Harga").Visible = True
    fNew_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fPrev_Click()
On Error Resume Next
    a = fQuick
    Mid(a, 1) = zerofill(Left(a, 5) - 1, 5)
    fQuick = a
    GetResult fQuick
End Sub

Private Sub fPrint_Click()
    FormPreview.LoadMe Me, "SPP", fNo
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

Sub GetResult(ByVal tNo As String)
On Error Resume Next
    LDetail = False
    b = Left(tNo, 5) & "/" & Right(tNo, 2)
    fQuick = b
    a = "select IdtSPP, t_SPP~.Kode, Nama, Alamat, Telepon, Fax, ContactPerson, t_SPP~.NamaPenerima, t_SPP~.AlamatPenerima, t_SPP~.AlamatPendek, NoSPP, TanggalSPP, TanggalKirim, MataUang, t_SPP~.WaktuPembayaran, KeteranganSPP, Status, NoKW, Pengupdate, WaktuUpdate from t_SPP~ left join m_customer on t_SPP~.Kode=m_Customer.Kode where ShortSPP='" & esc(b) & "'"
    ClearScreen
    query a
    If RS.RecordCount = 0 Then
        GoEvent "NEW"
        fNo = fQuick
        fNo_LostFocus
        Exit Sub
    End If
    fKode = RS.Fields("Kode").value
    fCustomer = RS.Fields("Nama").value
    fNamaPenerima = RS.Fields("NamaPenerima").value
    fAlamat = RS.Fields("Alamat").value
    fAlamatPenerima = RS.Fields("AlamatPenerima").value
    fAlamatPendek = RS.Fields("AlamatPendek").value
    fTelepon = RS.Fields("Telepon").value
    fFax = RS.Fields("Fax").value
    fContactPerson = RS.Fields("ContactPerson").value
    fNo = RS.Fields("NoSPP").value
    fNo.Tag = fNo
    fNoKW = RS.Fields("NoKW").value
    fTanggal = cTanggal(RS.Fields("TanggalSPP").value)
    fTanggalKirim = cTanggal(RS.Fields("TanggalKirim").value)
    fMataUang = RS.Fields("MataUang").value
    fWaktuPembayaran = RS.Fields("WaktuPembayaran").value
    fKeterangan = RS.Fields("KeteranganSPP").value
    fUpdatedBy = RS!Pengupdate & " " & RS!WaktuUpdate & " " & GetID(RS!IdtSPP)
    m_status = RS.Fields("Status").value
    If m_status >= 2 Then
        GoEvent "SEE"
    Else
        GoEvent "EDIT"
    End If
    fSetuju = IIf(m_status = 0, 0, 1)
    a = "select NoSC,  Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, Harga, t_SPPDetail~.JumlahBox, KeteranganSPPDetail, IdSC, DP, IdDet from t_SPPDetail~ left join m_stock~ on m_stock~.IdStock=t_SPPDetail~.IdStock " & _
        "where NoSPP='" & esc(fNo) & "' order by IdSPP"
    query a
    If RS.RecordCount > 0 Then
        x.LoadRows RS.GetRows
        TDBGrid1.Rebind
    End If
    a = fCustomer
    fCustomer_KeyDown 0, 0
    fCustomer = a
    fCustomer.FindIndex
    If fCustomer.ListIndex > -1 Then
        LPenerima = False
        a = fNamaPenerima
        fNamaPenerima_KeyDown 0, 0
        fNamaPenerima = a
        fNamaPenerima.FindIndex
    End If
End Sub

Private Sub fSave_Click()
On Error GoTo err
Dim tCurrent As Double
    tCurrent = 0
    BeginTransaction
    TDBGrid1.Update
    seb = IIf(x(0, iDP) & "" = "", 0, x(0, iDP))
    sebNoSC = x(0, iNoSC)
'    For i = 1 To x.UpperBound(1)
'        If x(i, iJumlah) <> "" Then
'            seb1 = IIf(x(i, iDP) & "" = "", 0, x(i, iDP))
'            If CByte(seb) <> CByte(seb1) And x(i, iJumlah) <> "" Then
'                MsgBox "DP Tidak Sama"
'                GoTo err
'            End If
'            If sebNoSC <> x(i, iNoSC) Then
'                MsgBox "No SC Tidak Sama"
'                GoTo err
'            End If
'        End If
'    Next
    m_status = fSetuju
    If m_mode = "EDIT" Then
        a = "update t_SPP~ set " & _
            "Kode=" & fKode & _
            ", NamaPenerima='" & esc(fNamaPenerima) & _
            "', AlamatPenerima='" & esc(fAlamatPenerima) & _
            "', NoSPP='" & esc(fNo) & _
            "', ShortSPP='" & esc(Left(fNo, 5) & "/" & Right(fNo, 2)) & _
            "', TanggalSPP=" & cD(fTanggal) & _
            ", TanggalKirim=" & cD(fTanggalKirim) & _
            ", MataUang='" & esc(fMataUang) & _
            "', WaktuPembayaran='" & esc(fWaktuPembayaran) & _
            "', KeteranganSPP='" & esc(fKeterangan) & _
            "', status=" & m_status & _
            ", AlamatPendek='" & fAlamatPendek & _
            "', WaktuUpdate=now" & _
            ", IdtSPP=" & cNum(Rnd * 2000000000, 0) & _
            ", PengUpdate='" & FormLogin.fUsr & "' where NoSPP='" & fNo.Tag & "'"
        If ExecMe(a) = 0 Then GoTo err
        'SendData "1EDIT SPP NO: " & fNo & Chr(8)
    ElseIf m_mode = "NEW" Then
        a = "insert into t_SPP~(Kode,NamaPenerima,AlamatPenerima,NoSPP,ShortSPP,TanggalSPP,TanggalKirim,MataUang,WaktuPembayaran,KeteranganSPP,Status,AlamatPendek,IdtSPP, PengUpdate) values(" & _
            fKode & _
            ",'" & fNamaPenerima & _
            "','" & fAlamatPenerima & _
            "','" & fNo & _
            "','" & Left(fNo, 5) & "/" & Right(fNo, 2) & _
            "'," & cD(fTanggal) & _
            "," & cD(fTanggalKirim) & _
            ",'" & fMataUang & _
            "'," & fWaktuPembayaran & _
            ",'" & fKeterangan & _
            "'," & m_status & _
            ",'" & fAlamatPendek & _
            "'," & cNum(Rnd * 2000000000) & _
            ",'" & FormLogin.fUsr & "')"
        If ExecMe(a) = 0 Then GoTo err
        'SendData "1BUAT SPP NO: " & fNo & Chr(8)
    End If
    Dim IdStock As Long
    For i = 0 To x.UpperBound(1)
        If x(i, iJumlah) <> "" Then
            If x(i, iJumlah) > 0 Then
                If x(i, iSatuan) = "" Then
                    MsgBox "Satuan Harus Diisi"
                    TDBGrid1.SetFocus
                    GoTo err
                End If
                c = "select top 1 IdStock from m_stock~ where NoWarna='" & esc(x(i, iNoWarna)) & _
                    "' and Jenis='" & esc(x(i, iJenis)) & _
                    "' and KodeBarang='" & esc(x(i, iKode)) & _
                    "' and Warna='" & esc(x(i, iWarna)) & _
                    "' and Tube='" & esc(x(i, iTube)) & _
                    "' and Grade='" & esc(x(i, iGrade)) & _
                    "' and SatBesar='" & esc(x(i, iSatuan)) & "'"
                query c
                If RS.RecordCount = 0 Then
                    MsgBox (x(i, iJenis) & " " & x(i, iKode) & " " & x(i, iWarna) & " " & x(i, iNoWarna) & " " & x(i, iTube) & " " & x(i, iGrade) & " TIDAK ADA DI DATA STOCK")
                    GoTo err
                Else
                    IdStock = RS.Fields(0).value
                End If
                IdSC = IIf(x(i, iIdSC) = "", 0, x(i, iIdSC))
                If x(i, iIdDet) <> "" Then
                    a = "update t_SPPDetail~ set " & _
                            "StatusDetail=" & m_status & _
                            ", KodeCustomerDetail=" & fKode & _
                            ", NoSPP='" & esc(fNo) & _
                            "', IdSPP=" & i & _
                            ", NoSC='" & esc(x(i, iNoSC)) & _
                            "', IdSC=" & IdSC & _
                            ", IdStock=" & IdStock & _
                            ", Satuan='" & esc(x(i, iSatuan)) & _
                            "', JumlahBox=" & cNum(x(i, iJumlah)) & _
                            ", Harga=" & cNum(x(i, iHarga)) & _
                            ", KeteranganSPPHeader='" & esc(fKeterangan) & _
                            "', KeteranganSPPDetail='" & esc(x(i, iKeterangan)) & _
                            "', MataUangDetail='" & esc(fMataUang) & _
                            "', TanggalKirimDetail=" & cD(fTanggalKirim) & _
                            ", TanggalSPPDetail=" & cD(fTanggal) & _
                            ", PengUpdate='" & esc(FormLogin.fUsr) & _
                            "', WaktuUpdate=now()" & _
                            ", NamaPenerimaDetail='" & esc(fNamaPenerima) & _
                            "', AlamatPenerimaDetail='" & esc(fAlamatPendek) & "' where IdDet=" & x(i, iIdDet)
                    If ExecMe(a) = 0 Then GoTo err
                Else
                    a = "insert into t_SPPDetail~(StatusDetail,KodeCustomerDetail,NoSPP, IdSPP, NoSC, IdSC, IdStock, Satuan, JumlahBox, Harga, KeteranganSPPHeader, KeteranganSPPDetail, MataUangDetail, TanggalKirimDetail,TanggalSPPDetail, PengUpdate, NamaPenerimaDetail, AlamatPenerimaDetail) values(" & m_status & "," & fKode & ",'" & fNo & _
                        "'," & i & _
                        ",'" & x(i, iNoSC) & _
                        "'," & IdSC & _
                        "," & IdStock & _
                        ",'" & x(i, iSatuan) & _
                        "'," & cNum(x(i, iJumlah)) & _
                        "," & cNum(x(i, iHarga)) & _
                        ",'" & fKeterangan & _
                        "','" & x(i, iKeterangan) & _
                        "','" & fMataUang & _
                        "'," & cD(fTanggalKirim) & _
                        "," & cD(fTanggal) & _
                        ",'" & FormLogin.fUsr & _
                        "','" & fNamaPenerima & _
                        "','" & fAlamatPendek & "')"
                    If ExecMe(a) = 0 Then GoTo err
                End If
                tCurrent = tCurrent + x(i, iJumlah) * 26 * x(i, iHarga)
            End If
        End If
    Next
    fNo.Tag = fNo
    s = "select sum(Total-Pelunasan) as a from t_SPPPE where Kode=" & fKode
    query s
    total = IIf(IsNull(RS!a), 0, RS!a)
    s = "select sum(Total-Pelunasan) as a from t_SPPDTY where Kode=" & fKode
    query s
    total = total + IIf(IsNull(RS!a), 0, RS!a)
    s = "select sum(Nilai) as a from t_STTPelunasanPE where TanggalPelunasan>" & cD(fTanggal.Text) & " and KodeCustomer=" & fKode
    query s
    total = total + IIf(IsNull(RS!a), 0, RS!a)
    s = "select sum(Nilai) as a from t_STTPelunasanDTY where TanggalPelunasan>" & cD(fTanggal.Text) & " and KodeCustomer=" & fKode
    query s
    total = total + IIf(IsNull(RS!a), 0, RS!a)
    s = "select Limit from m_Customer where Kode=" & fKode
    query s
    If total + tCurrent > RS!Limit Then
        If MsgBox("Lewat Limit! Lanjutkan?", vbYesNo) = vbNo Then
            GoTo err
        End If
    End If

    CommitTransaction
    MsgBox "SUKSES"
    DoEvents
    GetResult fQuick
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fSetuju_Validate(Cancel As Boolean)
    If m_status < 2 Then
        a = "update t_SPP~ set Status=" & fSetuju.value & ", WaktuUpdate=now, Pengupdate='" & FormLogin.fUsr & "' where NoSPP='" & esc(fNo) & "'"
        ExecMe a
        a = "update t_SPPDetail~ set StatusDetail=" & fSetuju.value & ", WaktuUpdate='" & Now & "', Pengupdate='" & FormLogin.fUsr & "' where NoSPP='" & esc(fNo) & "'"
        ExecMe a
    End If
End Sub

Private Sub fStock_Click()
On Error Resume Next
    FormStock.LoadMe Me, col1(iJenis).value, col1(iKode).value, col1(iNoWarna).value, col1(iTube).value, col1(iGrade).value, col1(iSatuan).value
End Sub

Private Sub fTanggal_LostFocus()
    fNo_LostFocus
End Sub

Private Sub fUpdateCustomer_Click()
On Error GoTo err
    If fCustomer = "" Then Exit Sub
    a = "update m_customer set " & _
        "Nama='" & fCustomer & _
        "',Alamat='" & fAlamat & _
        "',AlamatPenerima='" & fAlamatPenerima & _
        "',NamaPenerima='" & fNamaPenerima & _
        "',Telepon='" & fTelepon & _
        "',Fax='" & fFax & _
        "',ContactPerson='" & fcontact & _
        "' where Kode=" & fCustomer.GetData("Kode")
    ExecMe a
    LCustomer = False
    MsgBox "SUKSES"
    DoEvents
    'SendData "1UPDATE CUSTOMER " & fCustomer & Chr(8)
    Exit Sub
err:
    MsgBox "GAGAL"
End Sub

Private Sub TDBDropDown1_Paint()
On Error Resume Next
    TDBGrid1.SelLength = Len(TDBGrid1.Text)
End Sub

Private Sub TDBDropDown2_DropDownClose()
On Error Resume Next
    col1(iSatuan).value = TDBDropDown2.Columns("Satuan").value
End Sub

Private Sub TDBDropDown2_Paint()
On Error Resume Next
    TDBGrid1.SelLength = Len(TDBGrid1.Text)
End Sub


Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo err
    If ColIndex = iNoSC Then
        i = TDBDropDown1.Bookmark
        If i < 0 Or IsNull(i) Then
            For j = 1 To 6
                col1(j).value = ""
            Next
            col1(iHarga).value = ""
            col1(iIdSC).value = ""
            col1(iSatuan).value = ""
            col1(iDP).value = ""
            Exit Sub
        End If
        For j = 0 To 6
            col1(j).value = z(i, j)
        Next
        col1(iHarga).value = TDBDropDown1.Columns("HARGA").value
        col1(iIdSC).value = TDBDropDown1.Columns("IdSC").value
        col1(iSatuan).value = ""
        col1(iDP).value = TDBDropDown1.Columns("DP").value
    End If
err:
End Sub

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)
On Error GoTo err
    If MsgBox("Yakin Hapus?", vbYesNo) = vbNo Then Exit Sub
    If col1("IdDet").value = "" Then Exit Sub
    BeginTransaction
    a = "delete from t_SPPDetail~ where IdDet=" & col1("IdDet").value
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBGrid1_GotFocus()
    If Not LDetail Then
        LDetail = True
        LoadGrid1
    End If
    If Not LSat Then
        col1(iSatuan).DropDown = TDBDropDown2
        
        col1(iSatuan).AutoDropDown = True
        TDBDropDown2.ColumnHeaders = False
        Set TDBDropDown2.Array = zSat
        a = "select distinct SatBesar from m_Stock~ where IsActive=1 order by SatBesar"
        query a
        If RS.RecordCount > 0 Then zSat.LoadRows RS.GetRows
        TDBDropDown2.Rebind
        LSat = True
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
    TDBGrid1.Columns(iNoSC).AutoDropDown = True
    'TDBGrid1.Columns(iNoSC).AutoDropDown = False
    
    
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err
    If col1(iSatuan).value <> "" Then Exit Sub
    If LastCol = iWarna Or LastCol = iKode Or LastCol = iJenis Then
        a = "select top 1 SatBesar from m_stock~ where IsActive=1 and KodeBarang='" & esc(col1(iKode).value) & "' and Warna='" & esc(col1(iWarna).value) & "' and Jenis='" & esc(col1(iJenis).value) & "'"
        query a
        If RS.RecordCount <= 0 Then
            col1(iSatuan).value = ""
            Exit Sub
        End If
        col1(iSatuan).value = RS.Fields(0).value
    End If
    Exit Sub
err:
End Sub

Private Sub TDBDropDown1_DropDownClose()
On Error GoTo err
    i = TDBDropDown1.Bookmark
    If i < 0 Or IsNull(i) Then
        For j = 1 To 6
            col1(j).value = ""
        Next
        col1(iHarga).value = ""
        col1(iIdSC).value = ""
        col1(iSatuan).value = ""
        col1(iDP).value = ""
        Exit Sub
    End If
    For j = 0 To 6
        col1(j).value = z(i, j)
    Next
    col1(iHarga).value = TDBDropDown1.Columns("HARGA").value
    col1(iIdSC).value = TDBDropDown1.Columns("IdSC").value
    col1(iSatuan).value = ""
    col1(iDP).value = TDBDropDown1.Columns("DP").value
err:
End Sub


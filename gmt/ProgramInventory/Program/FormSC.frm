VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{8AAEAB20-E970-42F3-9E69-BC54C54CC273}#4.0#0"; "usrcombo.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{BD09B73E-A5EF-4CAB-A002-921A8335B40E}#1.0#0"; "UsrTrueCombo.ocx"
Begin VB.Form FormSC 
   BackColor       =   &H00FFC0C0&
   Caption         =   "SALES CONTRACT"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fPrint2 
      Caption         =   "PRINT2"
      Height          =   375
      Left            =   5040
      TabIndex        =   48
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox fKeterangan 
      Height          =   735
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CommandButton fDetailSC 
      Caption         =   "Detail SC"
      Height          =   375
      Left            =   4800
      TabIndex        =   43
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton fOutstandingSC 
      Caption         =   "Outstanding SC"
      Height          =   375
      Left            =   3240
      TabIndex        =   42
      Top             =   600
      Width           =   1455
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2415
      Left            =   1980
      TabIndex        =   41
      Top             =   4560
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode Barang"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Warna"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tube"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Satuan"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=145"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1323"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1244"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=74843660"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=64"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1588"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1508"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4471
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
      Columns(6).Caption=   "Satuan"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Jumlah"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Terpakai"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Harga"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "OK?"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Keterangan"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=953"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2619"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2540"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1058"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=979"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1693"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1614"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1323"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1244"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(4)._MinWidth=1312901971"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=979"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=900"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=1111"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=1032"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=1323"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=1244"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=1482"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=1402"
      Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=2170"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=2090"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(10).Width=847"
      Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=767"
      Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(46)=   "Column(11).Width=2434"
      Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2355"
      Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
      _StyleDefs(84)  =   "Named:id=33:Normal"
      _StyleDefs(85)  =   ":id=33,.parent=0"
      _StyleDefs(86)  =   "Named:id=34:Heading"
      _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=34,.wraptext=-1"
      _StyleDefs(89)  =   "Named:id=35:Footing"
      _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=36:Selected"
      _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=37:Caption"
      _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(95)  =   "Named:id=38:HighlightRow"
      _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=39:EvenRow"
      _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(99)  =   "Named:id=40:OddRow"
      _StyleDefs(100) =   ":id=40,.parent=33"
      _StyleDefs(101) =   "Named:id=41:RecordSelector"
      _StyleDefs(102) =   ":id=41,.parent=34"
      _StyleDefs(103) =   "Named:id=42:FilterBar"
      _StyleDefs(104) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton fMasterCustomer 
      Caption         =   "CUSTOMER"
      Height          =   375
      Left            =   1920
      TabIndex        =   40
      Top             =   600
      Width           =   1215
   End
   Begin UsrText.IText fNilaiKontrak 
      Height          =   270
      Left            =   6960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
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
   Begin VB.CheckBox fDP 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DP"
      Height          =   255
      Left            =   8640
      TabIndex        =   39
      Top             =   3000
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CommandButton fStock 
      Caption         =   "STOCK"
      Height          =   375
      Left            =   9360
      TabIndex        =   38
      Top             =   120
      Width           =   975
   End
   Begin UsrTrueCombo.ITrueCombo fCustomer 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
      _ExtentX        =   8493
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
   Begin UsrText.IText fQuick 
      Height          =   330
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.CommandButton fClose 
      Caption         =   "&CLOSE"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton fList 
      Caption         =   "LIST"
      Height          =   375
      Left            =   3360
      TabIndex        =   36
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   10680
      TabIndex        =   35
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox fAlamat 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   6375
   End
   Begin UsrText.IText fWaktuPembayaran 
      Height          =   270
      Left            =   9840
      TabIndex        =   5
      Top             =   1440
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
   Begin UsrText.IText fLamaKontrak 
      Height          =   270
      Left            =   8280
      TabIndex        =   4
      Top             =   1440
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
   Begin UsrCombo.ICombo fMataUang 
      Height          =   315
      Left            =   6960
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
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
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   9000
      TabIndex        =   2
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
   Begin UsrText.IText fNo 
      Height          =   270
      Left            =   6960
      TabIndex        =   1
      Top             =   840
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
   Begin VB.CommandButton fPrint 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton fNew 
      Caption         =   "&NEW"
      Height          =   375
      Left            =   8880
      TabIndex        =   18
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin UsrText.IText fTelepon 
      Height          =   270
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
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
      TabIndex        =   12
      TabStop         =   0   'False
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4455
      _ExtentX        =   7858
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
   Begin VB.CheckBox fSetuju 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Disetujui"
      Height          =   375
      Left            =   6480
      TabIndex        =   37
      Top             =   3600
      Width           =   1215
   End
   Begin UsrText.IText fNamaMarketing 
      Height          =   270
      Left            =   120
      TabIndex        =   44
      Top             =   3840
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
   Begin UsrText.IText fNamaCustomerSC 
      Height          =   270
      Left            =   3000
      TabIndex        =   46
      Top             =   3840
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
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA CONTACT CUSTOMER"
      Height          =   255
      Left            =   3000
      TabIndex        =   47
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA MARKETING"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA CUSTOMER"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TELEPON"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT PERSON"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "NILAI CONTRACT"
      Height          =   255
      Left            =   6960
      TabIndex        =   29
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "HARI"
      Height          =   255
      Left            =   10800
      TabIndex        =   28
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "HARI"
      Height          =   255
      Left            =   9240
      TabIndex        =   27
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN"
      Height          =   255
      Left            =   6960
      TabIndex        =   26
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "WAKTU PEMBAYARAN"
      Height          =   255
      Left            =   9840
      TabIndex        =   25
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "LAMA KONTRAK"
      Height          =   255
      Left            =   8280
      TabIndex        =   24
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MATA UANG"
      Height          =   255
      Left            =   6960
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
      Height          =   255
      Left            =   9000
      TabIndex        =   22
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "NO SC"
      Height          =   255
      Left            =   6960
      TabIndex        =   21
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SALES CONTRACT"
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
      Left            =   6480
      TabIndex        =   20
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FormSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LCustomer As Boolean
Dim LMataUang As Boolean
Dim LKodeBarang As Boolean
Dim fKode As Long
Dim m_mode As String
Dim m_status As Byte
Dim z As New XArrayDB
Dim x As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim iJenis As Integer
Dim iKode As Integer
Dim iWarna As Integer
Dim iNoWarna As Integer
Dim iTube As Integer
Dim iGrade As Integer
Dim iSatuan As Integer
Dim iJumlah As Integer
Dim iTerpakai As Integer
Dim iHarga As Integer
Dim iOK As Integer
Dim iKeterangan As Integer

Private Sub fDetailSC_Click()
    FormPreview.LoadMe Me, "LaporanDetailSC", fNo
End Sub

Private Sub fKeterangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        HelpMe "Keterangan", Me
    End If
End Sub

Sub FormHelpKeyDown(ByVal tVal As String)
    If ActiveControl.Name = "fKeterangan" Then
        ActiveControl.Text = tVal
    End If
End Sub

Private Sub GoEvent(ByVal tEvent As String)
    If tEvent = "NEW" Then
        fCustomer.Enabled = True
        fNo.Enabled = True
        m_mode = "NEW"
        tProp = 0
    ElseIf tEvent = "EDIT" Then
        m_mode = "EDIT"
        tProp = 1
    End If
    v = IIf((tProp And 1) = 0, False, True)
        fDelete.Enabled = v
        fPrint.Enabled = v
        TDBGridSetVisible TDBGrid1, "Terpakai@OK?", v
End Sub

Private Sub ClearScreen()
    fCustomer = ""
    fAlamat = ""
    fTelepon = ""
    fFax = ""
    fContactPerson = ""
    fKode = -1
    fNo = ""
    fNo.Tag = ""
    fTanggal = pServerDate
    fMataUang = "RP"
    fLamaKontrak = ""
    fNamaCustomerSC = ""
    fNamaMarketing = ""
    
    fWaktuPembayaran = "30"
    fKeterangan = ""
    fDP.Value = 0
    fNilaiKontrak = "0"
    fSetuju = 0
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
End Sub


Private Sub fAlamat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys Chr(9)
End Sub

Private Sub fClose_Click()
On Error GoTo err
    BeginTransaction
    st = IIf(fClose.Caption = "&CLOSE", 1, 0)
    a = "update t_SCDetail" & pTipe & " set StatusDetail=" & st & " where NoSC='" & esc(fNo.Tag) & "'"
    ExecMe a
    b = "update t_SC" & pTipe & " set Status=" & st & " where NoSC='" & esc(fNo.Tag) & "'"
    ExecMe b
    CommitTransaction
    MsgBox "SUKSES"
    If st = 0 Then fClose.Caption = "&CLOSE" Else fClose.Caption = "&UNCLOSE"
    GetResult fNo.Tag
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    If KeyCode = 116 Then LCustomer = False
    If Not LCustomer Then
        fCustomer.SetHeader "Nama Customer@Alamat@*Telepon@*Fax@*ContactPerson@*Kode@*WaktuPembayaran"
        fCustomer.SetWidth "2500@4000"
        Dim rs1() As Variant
        a = "select Nama,Alamat,Telepon,Fax,ContactPerson,Kode,WaktuPembayaran from m_customer where IsActive=1 ORDER BY NAMA"
        query a
        rs1 = RS.GetRows
        fCustomer.SetDB rs1
        fCustomer.SetType "String@String@String@String@String"
        LCustomer = True
    End If
err:
End Sub

Private Sub fCustomer_LostFocus()
On Error Resume Next
    If fCustomer = "" Then Exit Sub
    fCustomer = fCustomer.GetData("Nama Customer")
    fAlamat = fCustomer.GetData("Alamat")
    fTelepon = fCustomer.GetData("Telepon")
    fFax = fCustomer.GetData("Fax")
    fContactPerson = fCustomer.GetData("ContactPerson")
    fKode = fCustomer.GetData("Kode")
    fWaktuPembayaran = fCustomer.GetData("WaktuPembayaran")
End Sub

Private Sub fDelete_Click()
On Error GoTo err
    BeginTransaction
    a = "select top 1 NoSC from t_SPPDetail" & pTipe & " where NoSC='" & esc(fNo) & "'"
    query a
    If RS.RecordCount > 0 Then
        MsgBox "SUDAH ADA DI SPP"
        GoTo err
    End If
    a = "delete from t_SC" & pTipe & " where nosc='" & esc(fNo.Tag) & "'"
    ExecMe a
    a = "delete from t_SCDetail" & pTipe & " where nosc='" & esc(fNo.Tag) & "'"
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    GoEvent "NEW"
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fList_Click()
    FormList.LoadMe "NOT CLOSED@CLOSED", _
        "select Disetujui, NoSC,Nama,TanggalSC,MataUang,LamaKontrak,NilaiKontrak,Keterangan from t_SC" & pTipe & " left join m_Customer on t_SC" & pTipe & ".Kode=m_Customer.Kode where status=0@" & _
        "select Disetujui, NoSC,Nama,TanggalSC,MataUang,LamaKontrak,NilaiKontrak,Keterangan from t_SC" & pTipe & " left join m_Customer on t_SC" & pTipe & ".Kode=m_Customer.Kode where status=1", _
        "No SC@Nama", "NoSC@Nama", "1000@3500", "String@String", _
        "@No SC@Nama Customer@Tanggal SC@Mata Uang @Lama Kontrak@Nilai Kontrak@Keterangan", _
        "700@2000@2500@1000@700@700@1500@10", _
        "YesNo@String@String@Date@String@Currency@Decimal@String", Me, " order by TanggalSC\10000 Desc, NoSC Desc", 1
    FormList.Show
End Sub

Private Sub fMasterCustomer_Click()
    FormMasterCustomer.Show , Me
End Sub

Private Sub fMasterMataUang_Click()
    FormInputMaster.LoadMe _
        "select Kode,Nama,Negara from m_matauang", _
        "Kode@Nama@Negara", "Kode@Nama@Negara", "700@1500@1500", "String@String@String", _
        "Kode@Nama@Negara", "KODE@NAMA@NEGARA", "700@1500@1500", "String@String@String", "Kode"
End Sub

Private Sub fMataUang_DropDown()
    If Not LMataUang Then
        a = "select Kode from m_MataUang"
        query a
        t = fMataUang
        fMataUang.Clear
        For i = 0 To RS.RecordCount - 1
            fMataUang.List(i) = RS.Fields("Kode").Value
            RS.MoveNext
        Next
        fMataUang = t
        LMataUang = True
    End If
End Sub

Private Sub fMataUang_KeyDown(KeyCode As Integer, Shift As Integer)
    fMataUang_DropDown
End Sub

Private Sub fNew_Click()
On Error Resume Next
    ClearScreen
    fNo_LostFocus
    m_status = 0
    GoEvent "NEW"
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
    BuatNomor fNo, fTanggal, pNomorSC, fQuick, "select max(nosc) from t_SC" & pTipe & " where TanggalSC>" & pAddNoLong
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LCustomer = False
    LMataUang = False
End Sub

Private Sub Form_Load()
    Set TDBDropDown1.Array = z
    Caption = Caption & "---" & pTipe
    fCustomer.ZOrder 0
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
    iJenis = 0
    iKode = 1
    iWarna = 2
    iNoWarna = 3
    iTube = 4
    iGrade = 5
    iSatuan = 6
    iJumlah = 7
    iTerpakai = 8
    iHarga = 9
    iOK = 10
    iKeterangan = 11
    col1(iKode).DropDown = TDBDropDown1
    col1(iKode).AutoDropDown = True
    TDBGridSetTag TDBGrid1, "Jumlah@Terpakai@Harga", "Decimal"
    TDBGridSetLock TDBGrid1, iTerpakai, True
    col1(iOK).Tag = "OK?"
    TDBGridLoad TDBGrid1
    LCustomer = False
    LMataUang = False
    LKodeBarang = False
    fMataUang = "RP"
    fMataUang_DropDown
    fNew_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fOutstandingSC_Click()
    FormPreview.LoadMe Me, "LaporanOutstandingSC"
End Sub

Private Sub fPrev_Click()
On Error Resume Next
    a = fQuick
    Mid(a, 1) = zerofill(Left(a, 5) - 1, 5)
    fQuick = a
    GetResult fQuick
End Sub


Private Sub fPrint_Click()
    FormPreview.LoadMe Me, "SC", fNo
End Sub
Private Sub fPrint2_Click()
    FormPreview.LoadMe Me, "SC2", fNo
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

Function GetResult(ByVal tNo As String) As Long
On Error Resume Next
    fQuick = Left(tNo, 5) & "/" & Right(tNo, 2)
    a = "select NoSC, t_SC~.Kode, Nama, Alamat, Telepon, Fax, ContactPerson, NoSC, TanggalSC, MataUang, LamaKontrak, t_SC" & pTipe & ".WaktuPembayaran, Keterangan, NilaiKontrak, Status, Disetujui, DP, NamaMarketing, NamaCustomerSC from t_SC" & pTipe & " left join m_customer on t_SC" & pTipe & ".Kode=m_Customer.Kode where ShortSC='" & esc(fQuick) & "'"
    ClearScreen
    query a
    If RS.RecordCount = 0 Then
        GetResult = 0
        GoEvent "NEW"
        fNo = fQuick
        fNo_LostFocus
        Exit Function
    End If
    fKode = RS.Fields("Kode").Value
    fCustomer = RS.Fields("Nama").Value
    fAlamat = RS.Fields("Alamat").Value
    fTelepon = RS.Fields("Telepon").Value
    fFax = RS.Fields("Fax").Value & ""
    fContactPerson = RS.Fields("ContactPerson").Value & ""
    fNo = RS.Fields("NoSC").Value
    fNo.Tag = fNo
    fTanggal = cTanggal(RS.Fields("TanggalSC").Value)
    fMataUang = RS.Fields("MataUang").Value
    fLamaKontrak = RS.Fields("LamaKontrak").Value
    fWaktuPembayaran = RS.Fields("WaktuPembayaran").Value
    fKeterangan = RS.Fields("Keterangan").Value
    fNamaMarketing = RS.Fields("NamaMarketing").Value
    fNamaCustomerSC = RS.Fields("NamaCustomerSC").Value
    m_status = RS.Fields("Status").Value
    If m_status = 0 Then fClose.Caption = "&CLOSE" Else fClose.Caption = "&UNCLOSE"
    fSetuju = RS.Fields("Disetujui").Value
    fDP = RS.Fields("DP").Value
    a = "select JenisBarang, KodeBarang, Warna, NoWarna, Tube, Grade, Satuan, Jumlah, Terpakai, Harga, StatusDetail, Keterangan from t_SCDetail" & pTipe & " where NoSC='" & esc(RS.Fields("NoSC").Value) & "' order by IdSC"
    query a
    If RS.RecordCount > 0 Then
        x.LoadRows RS.GetRows
        TDBGrid1.Rebind
        HitungNilaiKontrak
    End If
    fNo.Enabled = True
    fCustomer.Enabled = True
    GoEvent "EDIT"
    For i = 0 To x.UpperBound(1)
        If x(i, iTerpakai) > 0 Then
            fNo.Enabled = False
            fCustomer.Enabled = False
            Exit For
        End If
    Next
End Function

Private Sub fSave_Click()
'On Error GoTo err
    Dim Mode As String
    BeginTransaction
    If Not fCustomer.Validate Then GoTo err
    If Not fMataUang.Validate Then GoTo err
    If Not fWaktuPembayaran.Validate Then GoTo err
    TDBGrid1.Update
    If m_mode <> "NEW" Then
        a = "delete from t_SC~ where nosc='" & esc(fNo.Tag) & "'"
        ExecMe a
        a = "delete from t_SCDetail~ where nosc='" & esc(fNo.Tag) & "'"
        ExecMe a
    End If
    a = "insert into t_SC~(Kode,NamaCustomer, NamaMarketing, NamaCustomerSC, NoSC,ShortSC,TanggalSC,MataUang,LamaKontrak,WaktuPembayaran,NilaiKontrak,Keterangan,Status,Disetujui,DP,PengUpdate) values(" & _
        fKode & _
        ",'" & fCustomer & _
        "','" & fNamaMarketing & _
        "','" & fNamaCustomerSC & _
        "','" & fNo & _
        "','" & Left(fNo, 5) & "/" & Right(fNo, 2) & _
        "'," & cD(fTanggal) & _
        ",'" & fMataUang & _
        "'," & cNum(fLamaKontrak) & _
        "," & fWaktuPembayaran & _
        "," & cNum(fNilaiKontrak.Tag) & _
        ",'" & fKeterangan & _
        "'," & m_status & "," & fSetuju & "," & fDP & ",'" & FormLogin.fUsr & _
        "')"
    If ExecMe(a) = 0 Then GoTo err
    If m_mode = "NEW" Then
        'SendData "1BUAT SC NO: " & fNo & Chr(8)
    Else
        'SendData "1EDIT SC NO: " & fNo & Chr(8)
    End If
    a = "insert into t_SCDetail~(NoSC, IdSC, KetDetail, MataUangDetail, KodeCustomerDetail, NamaMarketingDetail, NamaCustomerSCDetail, WaktuPembayaranDetail, LamaKontrakDetail, TanggalSCDetail, JenisBarang, KodeBarang, Warna, NoWarna, Tube, Grade, Satuan, Jumlah, Terpakai, Harga, StatusDetail, Keterangan, DPDetail, PengUpdate) values('" & fNo & "'"
    For i = 0 To x.UpperBound(1)
        b = "," & i & _
            ",'" & fKeterangan & _
            "','" & fMataUang & _
            "'," & fKode & _
            ",'" & fNamaMarketing & _
            "','" & fNamaCustomerSC & _
            "'," & cNum(fWaktuPembayaran) & _
            "," & cNum(fLamaKontrak) & _
            "," & cD(fTanggal) & _
            ",'" & x(i, iJenis) & _
            "','" & x(i, iKode) & _
            "','" & x(i, iWarna) & _
            "','" & x(i, iNoWarna) & _
            "','" & x(i, iTube) & _
            "','" & x(i, iGrade) & _
            "','" & x(i, iSatuan) & _
            "'," & cNum(x(i, iJumlah)) & _
            "," & cNum(x(i, iTerpakai)) & _
            "," & cNum(x(i, iHarga)) & _
            "," & IIf(x(i, iOK) = 0, 0, 1) & _
            ",'" & esc(x(i, iKeterangan)) & _
            "'," & fDP & _
            ",'" & FormLogin.fUsr & "')"
        If ExecMe(a & b) = 0 Then GoTo err
        b = "update t_SPPDetail~ set MataUangDetail='" & fMataUang & "', DP=" & fDP.Value & ", Harga=" & cNum(x(i, iHarga)) & " where NoSC='" & esc(fNo) & "' and IdSC=" & i
        ExecMe b
    Next
    Dim tSPP() As String
    a = "select NoSPP from t_SPPDetail~ where NoSC='" & esc(fNo.Text) & "' group by NoSPP"
    query a
    If RS.RecordCount > 0 Then
        ReDim tSPP(RS.RecordCount - 1)
        For i = 0 To RS.RecordCount - 1
            tSPP(i) = RS.Fields(0).Value
            RS.MoveNext
        Next
        For i = 0 To UBound(tSPP)
            a = "select sum(JumlahKG*Harga) as Total from t_SPPDetail~ where NoSPP='" & tSPP(i) & "'"
            query a
            a = "update t_SPP~ set Total=" & cNum(RS!total) & " where NoSPP='" & tSPP(i) & "'"
            query a
        Next
    End If
    fNo.Tag = fNo
    CommitTransaction
    MsgBox "SUKSES"
    DoEvents
    GetResult fNo
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fSetuju_Validate(Cancel As Boolean)
On Error Resume Next
    a = "update t_SC~ set Disetujui=" & fSetuju.Value & " where NoSC='" & esc(fNo) & "'"
    ExecMe a
End Sub

Private Sub fStock_Click()
On Error Resume Next
    FormStock.LoadMe Me, col1(iJenis).Value, col1(iKode).Value, col1(iNoWarna).Value, col1(iTube).Value, col1(iGrade).Value
End Sub

Private Sub fTanggal_LostFocus()
    fNo_LostFocus
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
    a = TDBGrid1.Columns(ColIndex).Caption
    If a = "Jumlah" Then
        If m_mode = "NEW" Or col1(iTerpakai).Value = "" Then col1(iTerpakai).Value = 0
        HitungNilaiKontrak
    ElseIf a = "Harga" Then
        HitungNilaiKontrak
    ElseIf a = "Warna" Then
        a = "select top 1 satkecil from m_stock~ where IsActive=1 and Jenis='" & esc(col1(iJenis).Value) & "' and KodeBarang='" & esc(col1(iKode).Value) & "' and Warna='" & esc(col1(iWarna).Value) & "'"
        query a
        If Not IsNull(RS.Fields(0).Value) Then
            col1(iSatuan).Value = RS.Fields(0).Value
        End If
    End If
End Sub

Private Sub HitungNilaiKontrak()
On Error Resume Next
    TDBGrid1.Update
    a = 0
    For i = 0 To x.UpperBound(1)
        a = a + x(i, iJumlah) * x(i, iHarga)
    Next
    fNilaiKontrak.Tag = a
    fNilaiKontrak = cDecimal(fNilaiKontrak.Tag)
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = iKode Then
        col1(iKode).Value = TDBDropDown1.Columns("Kode Barang").Value
        col1(iWarna).Value = TDBDropDown1.Columns("Warna").Value
        col1(iTube).Value = TDBDropDown1.Columns("Tube").Value
        col1(iSatuan).Value = TDBDropDown1.Columns("Satuan").Value
    End If
End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    a = TDBGrid1.Columns(ColIndex).Caption
    If a = "Jumlah" Or a = "Harga" Then
        If pUpdateHargaSC = 0 And m_mode = "EDIT" Then Cancel = True
    End If
End Sub

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)
    If col1(iTerpakai).Value > 0 Then Cancel = True
End Sub

Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
    Response = False
End Sub

Private Sub TDBGrid1_GotFocus()
    If Not LKodeBarang Then
        a = "select distinct KodeBarang, Warna, Tube, SatKecil from m_Stock~ where IsActive=1 order by KodeBarang"
        query a
        z.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
        z.DeleteRows 0
        If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
        TDBDropDown1.Rebind
        LKodeBarang = True
    End If
End Sub

Private Sub TDBDropDown1_DropDownClose()
    col1(iKode).Value = TDBDropDown1.Columns("Kode Barang").Value
    col1(iWarna).Value = TDBDropDown1.Columns("Warna").Value
    col1(iTube).Value = TDBDropDown1.Columns("Tube").Value
    col1(iSatuan).Value = TDBDropDown1.Columns("Satuan").Value
End Sub

Private Sub TDBDropDown1_Paint()
    TDBGrid1.SelLength = Len(col1(TDBGrid1.Col).Text)
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub
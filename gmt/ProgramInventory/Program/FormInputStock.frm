VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Begin VB.Form FormInputStock 
   BackColor       =   &H00FFC0C0&
   Caption         =   "INPUT STOCK"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Tag             =   "28"
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5318
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&H8000000F&"
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
   Begin UsrText.IText fNoNota 
      Height          =   270
      Left            =   5280
      TabIndex        =   14
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2415
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   5535
      _ExtentX        =   9763
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
      Columns(3).Caption=   "Sat Besar"
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
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3863"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3784"
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
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
   Begin VB.CommandButton fFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton fLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6165
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
      Columns(1).Caption=   "Printed"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "OK?"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No Nota"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Jenis"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Kode"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Warna"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "No Warna"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Tube"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Grade"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Sat Besar"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Sat Kecil"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "IdStock"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "n1"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "n2"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Keterangan"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Printed Code"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "d1"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "d2"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "cKey"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "cIdStock"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "cOK"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   22
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=22"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1005"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=926"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=49"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=741"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=661"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=49"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1217"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1138"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(3)._MinWidth=49"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=926"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=847"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(4)._MinWidth=49"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2831"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2752"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(5)._MinWidth=-1"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1005"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=926"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=1455"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=1376"
      Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=1270"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=1191"
      Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(52)=   "Column(8)._MinWidth=-1"
      Splits(0)._ColumnProps(53)=   "Column(9).Width=900"
      Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=820"
      Splits(0)._ColumnProps(56)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(58)=   "Column(10).Width=1455"
      Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=1376"
      Splits(0)._ColumnProps(61)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(62)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(63)=   "Column(10)._MinWidth=176210976"
      Splits(0)._ColumnProps(64)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(65)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(67)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(68)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(69)=   "Column(11)._MinWidth=176210976"
      Splits(0)._ColumnProps(70)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(71)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(73)=   "Column(12)._ColStyle=8708"
      Splits(0)._ColumnProps(74)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(75)=   "Column(12)._MinWidth=176210976"
      Splits(0)._ColumnProps(76)=   "Column(13).Width=847"
      Splits(0)._ColumnProps(77)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(13)._WidthInPix=767"
      Splits(0)._ColumnProps(79)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(80)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(81)=   "Column(14).Width=1455"
      Splits(0)._ColumnProps(82)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(14)._WidthInPix=1376"
      Splits(0)._ColumnProps(84)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(85)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(86)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(87)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(89)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(90)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(91)=   "Column(16).Width=2223"
      Splits(0)._ColumnProps(92)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(16)._WidthInPix=2143"
      Splits(0)._ColumnProps(94)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(95)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(96)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(97)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(99)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(100)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(101)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(102)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(104)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(105)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(106)=   "Column(18)._MinWidth=75880512"
      Splits(0)._ColumnProps(107)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(108)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(110)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(111)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(112)=   "Column(19)._MinWidth=75908624"
      Splits(0)._ColumnProps(113)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(117)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(118)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(119)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(120)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(121)=   "Column(21)._ColStyle=516"
      Splits(0)._ColumnProps(122)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(123)=   "Column(21)._MinWidth=-1"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=77,.parent=2,.namedParent=79"
      _StyleDefs(17)  =   "FilterBarStyle:id=80,.parent=1,.namedParent=82"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=78,.parent=77"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=81,.parent=80"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=90,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=52,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=49,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=50,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=51,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=106,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=56,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=53,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=54,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=55,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=76,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=73,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=74,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=75,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=24,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=28,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=40,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=37,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=38,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=39,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=44,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=41,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=42,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=43,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=48,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=45,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=46,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=47,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=114,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=111,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=112,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=113,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=122,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=119,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=120,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=121,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=64,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=61,.parent=12,.alignment=2"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=62,.parent=13,.alignment=3"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=63,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=68,.parent=11"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=65,.parent=12"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=66,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=67,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=72,.parent=11"
      _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=69,.parent=12"
      _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=70,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=71,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=60,.parent=11"
      _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=57,.parent=12"
      _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=58,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=59,.parent=15"
      _StyleDefs(94)  =   "Splits(0).Columns(16).Style:id=86,.parent=11"
      _StyleDefs(95)  =   "Splits(0).Columns(16).HeadingStyle:id=83,.parent=12"
      _StyleDefs(96)  =   "Splits(0).Columns(16).FooterStyle:id=84,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(16).EditorStyle:id=85,.parent=15"
      _StyleDefs(98)  =   "Splits(0).Columns(17).Style:id=94,.parent=11"
      _StyleDefs(99)  =   "Splits(0).Columns(17).HeadingStyle:id=91,.parent=12"
      _StyleDefs(100) =   "Splits(0).Columns(17).FooterStyle:id=92,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(17).EditorStyle:id=93,.parent=15"
      _StyleDefs(102) =   "Splits(0).Columns(18).Style:id=98,.parent=11"
      _StyleDefs(103) =   "Splits(0).Columns(18).HeadingStyle:id=95,.parent=12"
      _StyleDefs(104) =   "Splits(0).Columns(18).FooterStyle:id=96,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(18).EditorStyle:id=97,.parent=15"
      _StyleDefs(106) =   "Splits(0).Columns(19).Style:id=102,.parent=11"
      _StyleDefs(107) =   "Splits(0).Columns(19).HeadingStyle:id=99,.parent=12"
      _StyleDefs(108) =   "Splits(0).Columns(19).FooterStyle:id=100,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(19).EditorStyle:id=101,.parent=15"
      _StyleDefs(110) =   "Splits(0).Columns(20).Style:id=110,.parent=11"
      _StyleDefs(111) =   "Splits(0).Columns(20).HeadingStyle:id=107,.parent=12"
      _StyleDefs(112) =   "Splits(0).Columns(20).FooterStyle:id=108,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(20).EditorStyle:id=109,.parent=15"
      _StyleDefs(114) =   "Splits(0).Columns(21).Style:id=118,.parent=11"
      _StyleDefs(115) =   "Splits(0).Columns(21).HeadingStyle:id=115,.parent=12"
      _StyleDefs(116) =   "Splits(0).Columns(21).FooterStyle:id=116,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(21).EditorStyle:id=117,.parent=15"
      _StyleDefs(118) =   "Named:id=29:Normal"
      _StyleDefs(119) =   ":id=29,.parent=0"
      _StyleDefs(120) =   "Named:id=30:Heading"
      _StyleDefs(121) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(122) =   ":id=30,.wraptext=-1"
      _StyleDefs(123) =   "Named:id=31:Footing"
      _StyleDefs(124) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(125) =   "Named:id=32:Selected"
      _StyleDefs(126) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(127) =   "Named:id=33:Caption"
      _StyleDefs(128) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(129) =   "Named:id=34:HighlightRow"
      _StyleDefs(130) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(131) =   "Named:id=35:EvenRow"
      _StyleDefs(132) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(133) =   "Named:id=36:OddRow"
      _StyleDefs(134) =   ":id=36,.parent=29"
      _StyleDefs(135) =   "Named:id=79:RecordSelector"
      _StyleDefs(136) =   ":id=79,.parent=30"
      _StyleDefs(137) =   "Named:id=82:FilterBar"
      _StyleDefs(138) =   ":id=82,.parent=29"
   End
   Begin VB.CommandButton fMasterStock 
      Caption         =   "S&TOCK"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   120
      TabIndex        =   3
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari No Nota"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   120
      Width           =   1095
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
      Height          =   375
      Left            =   2580
      TabIndex        =   11
      Top             =   720
      Width           =   5595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   765
   End
End
Attribute VB_Name = "FormInputStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim x2 As New XArrayDB
Dim m_mode As String
Dim m_status As Byte
Dim col1 As TrueOleDBGrid80.Columns
Dim LLoadDropDown As Boolean
Dim z As New XArrayDB
Dim iUpdated As Integer
Dim iPrinted As Integer
Dim iOK As Integer
Dim iNoNota As Integer
Dim iJenis As Integer
Dim iKode As Integer
Dim iWarna As Integer
Dim iNoWarna As Integer
Dim iTube As Integer
Dim iGrade As Integer
Dim iSatB As Integer
Dim iSatK As Integer
Dim iIdStock As Integer
Dim in1 As Integer
Dim in2 As Integer
Dim iKet As Integer
Dim iPrintedCode As Integer
Dim iD1 As Integer
Dim id2 As Integer
Dim icKey As Integer
Dim icIdStock As Integer
Dim icOK As Integer
Dim cArr As New Collection

Private Sub fDelete_Click()
    If InStr(col1(iKet).Value, "RETUR DARI") <> 0 Then
        MsgBox "Hapus Dari Retur Stock!!!"
        Exit Sub
    End If
    b = MsgBox("Hapus Data?", vbYesNo)
    If b = vbNo Then Exit Sub
    If col1(icKey).Value = "" Then
        cArr.Remove CStr(TDBGrid1.Bookmark)
        TDBGrid1.Delete
        
        Exit Sub
    End If
    BeginTransaction
    If Not cekValid("DELETE", Tag) Then GoTo err
    a = "delete from t_InputStock~ where TransID=" & col1(icKey).Value & " and Tanggal=" & cD(fTanggal)
    If ExecMe(a) = 0 Then GoTo err
    If CLng(col1(icOK).Value) <> 0 Then
        a = "update m_Stock~ set JumlahBox=JumlahBox-" & cNum(col1(iD1).Value) & ", JumlahKG=JumlahKG-" & cNum(col1(id2).Value) & " where IdStock=" & col1(icIdStock).Value
        If ExecMe(a) = 0 Then GoTo err
    End If
    CommitTransaction
    MsgBox "SUKSES"
    TDBGrid1.Delete
    fTanggal_Validate False
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fFirst_Click()
On Error Resume Next
    fTanggal = pServerDate
    a = "select min(Tanggal) from t_InputStock~"
    query a
    If Not IsNull(RS.Fields(0).Value) Then fTanggal = cTanggal(RS.Fields(0).Value)
    GetResult cD(fTanggal)
End Sub

Private Sub fLast_Click()
On Error Resume Next
    fTanggal = pServerDate
    a = "select max(Tanggal) from t_InputStock~"
    query a
    If Not IsNull(RS.Fields(0).Value) Then fTanggal = cTanggal(RS.Fields(0).Value)
    GetResult cD(fTanggal)
End Sub

Private Sub fMasterStock_Click()
    FormStock.LoadMe Me, col1(iJenis).Value & "", col1(iKode).Value & "", col1(iNoWarna).Value & "", col1(iTube).Value & "", col1(iGrade).Value & ""
End Sub

Private Sub fNext_Click()
On Error Resume Next
    fTanggal = add_tanggal(fTanggal, 1)
    GetResult cD(fTanggal)
End Sub

Private Sub fNoNota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        fNoNota.Cancel = True
        a = "select top 1 Tanggal from t_InputStock~ where NoBukti=" & fNoNota
        query a
        If RS.RecordCount = 0 Then
            MsgBox "Tidak Ada"
            Exit Sub
        End If
        fTanggal = cTanggal(RS.Fields(0).Value)
        GetResult RS.Fields(0).Value
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Caption = Caption & "--" & pTipe
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Load()
    iUpdated = 0
    iPrinted = 1
    iOK = 2
    iNoNota = 3
    iJenis = 4
    iKode = 5
    iWarna = 6
    iNoWarna = 7
    iTube = 8
    iGrade = 9
    iSatB = 10
    iSatK = 11
    iIdStock = 12
    in1 = 13
    in2 = 14
    iKet = 15
    iPrintedCode = 16
    iD1 = 17
    id2 = 18
    icKey = 19
    icIdStock = 20
    icOK = 21
    Set TDBGrid1.Array = x
    Set TDBGrid2.Array = x2
    Set TDBDropDown1.Array = z
    Set col1 = TDBGrid1.Columns
    TDBGrid1.AllowAddNew = True
    col1(in1).Tag = "Integer"
    col1(in2).Tag = "Decimal"
    col1(iOK).Tag = "OK?"
    col1(iPrinted).Tag = "OK?"
    
    TDBGridLoad TDBGrid1
    'TDBGridSetVisible TDBGrid1, "d1@d2@updated@cKey@cIdStock@cOK@IdStock@Sat Kecil", False
    TDBGridSetVisible TDBGrid1, "d1@d2@updated@cIdStock@cOK@IdStock@Sat Kecil", False
    TDBGridSetLock TDBGrid1, "Printed Code@Warna@Tube@Sat Besar", True
    fTanggal = pServerDate
    x2.ReDim 0, 9, 0, 9
    TDBGrid2.Columns.Clear
    For i = 1 To 10
        TDBGrid2.Columns.Add(i - 1).Caption = i
        TDBGrid2.Columns(i - 1).Visible = True
        TDBGrid2.Columns(i - 1).Width = 1000
        TDBGrid2.Columns(i - 1).Alignment = dbgRight
    Next
    For i = 0 To TDBGrid2.Columns.Count - 1
        TDBGrid2.Columns(i).Tag = "Decimal"
    Next
    TDBGrid2.Rebind
    GetResult cD(fTanggal)
    TDBGrid1.FetchRowStyle = True
    col1(iKode).DropDown = TDBDropDown1
    col1(iKode).AutoDropDown = True

End Sub

Sub GetResult(ByVal tDate As Long)
On Error Resume Next
    Set cArr = Nothing
    a = "select 0, 1-Status, Status, NoBukti, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, m_Stock~.IdStock, n1,n2, Keterangan, PrintedCode,n1 ,n2, TransID, m_Stock~.IdStock, Status from t_InputStock~ left join m_Stock~ on t_InputStock~.IdStock=m_Stock~.IdStock where Tanggal=" & tDate & " order by NoBukti"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    If Not LLoadDropDown Then
        a = "select distinct KodeBarang, Warna, Tube, SatBesar from m_Stock~ where IsActive=1 order by KodeBarang"
        query a
        z.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
        z.DeleteRows 0
        If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
        TDBDropDown1.Rebind
        LLoadDropDown = True
    End If
    TDBGrid1.MoveLast
    TDBGrid1.SetFocus
    GetFromCollection x2, TDBGrid1.Bookmark
    TDBGrid2.Rebind
    fPrint.Enabled = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid2.Width = ScaleWidth - 2 * TDBGrid2.Left
    'TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
    fKet.Width = ScaleWidth - fKet.Left - 100
End Sub

Private Sub fPrev_Click()
On Error Resume Next
    fTanggal = add_tanggal(fTanggal, -1)
    GetResult cD(fTanggal)
End Sub

Private Sub fPrint_Click()
'On Error GoTo err
Dim y As New XArrayDB
Dim j As Integer
    'iNoNota = 0
    'iNamaBarang = 1
    'in1 = 2
    'iSatB = 3
    'in2 = 4
    'iSatK = 5
    'iPrintedCode = 6
    If Not fPrint.Enabled Then Exit Sub
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, iUpdated) <> 0 Then
            MsgBox "Ada yang belum Diupdate!!!"
            Exit Sub
        End If
    Next
    y.ReDim 0, x.UpperBound(1), 0, x.UpperBound(2) + 1
    j = 0
    For i = 0 To x.UpperBound(1)
        If x(i, iPrinted) <> 0 Then
            y(j, 0) = x(i, iNoNota)
            y(j, 1) = x(i, iJenis) & " " & x(i, iKode) & " " & x(i, iWarna) & " " & x(i, iNoWarna) & " " & x(i, iTube) & " GRADE " & x(i, iGrade)
            y(j, 2) = x(i, in1)
            y(j, 3) = x(i, iSatB)
            y(j, 4) = x(i, in2)
            y(j, 5) = x(i, iSatK)
            y(j, 6) = x(i, iPrintedCode)
            If x(i, iSatB) = x(i, iSatK) Then
                y(j, 7) = ""
            Else
                s = "select nDet from t_InputStock~ where TransID=" & x(i, icKey)
                query s
                s2 = Split(RS.Fields(0).Value, "__")
                s3 = ""
                For k = 0 To UBound(s2) - 1
                    s3 = s3 & s2(k) & "__"
                    If k = 9 Then s3 = s3 & vbCrLf
                Next
                y(j, 7) = s3
            End If
            j = j + 1
        End If
    Next
    y.ReDim 0, j - 1, 0, y.UpperBound(2)
    FormPreview.LoadFromData Me, "InputStock", y, fTanggal
err:
End Sub

Private Function CreateDetail(ByVal tKey As String) As String
    Dim k As Integer
    Dim a As New XArrayDB
    k = 0
    a.ReDim 0, 9, 0, 9
    Set a = cArr(tKey)
    For i = 0 To a.UpperBound(1)
        For j = 0 To a.UpperBound(2)
            If a(i, j) <> "" And Not IsEmpty(a(i, j)) Then
                k = k + 1
                If k = 11 Then CreateDetail = CreateDetail & vbCrLf
                CreateDetail = CreateDetail & Format(a(i, j), "0.00") & "__"
            End If
        Next
    Next
    
End Function
Private Sub fSave_Click()
'On Error GoTo err
Dim PrintedCode As Long
Dim IdStock As Long
Dim nDet As String
    TDBGrid1.Update
    Randomize
    AddToCollection x2, TDBGrid1.Bookmark
    BeginTransaction
    If Not cekValid("EDIT", Tag) Then GoTo err
    For i = 0 To x.UpperBound(1)
        If x(i, iUpdated) = 1 Then
            nDet = ""
            
            a = "select top 1 IdStock, SatBesar, SatKecil from m_Stock~ where Jenis='" & esc(x(i, iJenis)) & "' and KodeBarang='" & esc(x(i, iKode)) & "' and Warna='" & esc(x(i, iWarna)) & _
                "' and NoWarna='" & esc(x(i, iNoWarna)) & "' and Tube='" & esc(x(i, iTube)) & "' and Grade='" & esc(x(i, iGrade)) & "' and SatBesar='" & esc(x(i, iSatB)) & "'"
            query a
            
            If RS.RecordCount < 1 Then
                MsgBox "Nama Stock(" & x(i, iNoWarna) & ") Salah"
                GoTo err
            End If
            IdStock = RS.Fields(0).Value
            
            If RS.Fields(1).Value = RS.Fields(2).Value And CInt(x(i, in1)) <> CInt(x(i, in2)) Then
                MsgBox "Jumlah harus sama untuk satuan yang sama!!!"
                TDBGrid1.Bookmark = i
                
                GoTo err
            End If
            If RS.Fields(1).Value <> RS.Fields(2).Value Then nDet = Replace(CreateDetail(i), vbCrLf, "")
            
            
            PrintedCode = Rnd * 2 ^ 21
            
            If x(i, icKey) = "" Then
                a = "insert into t_InputStock~(NoBukti, Status, Tanggal, IdStock, n1, n2, Keterangan, PrintedCode, nDet) values(" & _
                    x(i, iNoNota) & _
                    "," & IIf(x(i, iOK) = 0, 0, 1) & _
                    "," & cD(fTanggal) & _
                    "," & IdStock & _
                    "," & cNum(x(i, in1)) & _
                    "," & cNum(x(i, in2)) & _
                    ",'" & esc(x(i, iKet)) & _
                    "'," & PrintedCode & _
                    ",'" & nDet & _
                    "')"
                If ExecMe(a) <= 0 Then GoTo err
                If CLng(x(i, iOK)) <> 0 Then
                    a = "update m_Stock~ set JumlahBox=JumlahBox+" & cNum(x(i, in1)) & ", JumlahKG=JumlahKG+" & cNum(x(i, in2)) & " where IdStock=" & IdStock
                    If ExecMe(a) = 0 Then GoTo err
                End If
            Else
                If CLng(x(i, icOK)) <> 0 Then
                    a = "update m_Stock~ set JumlahBox=JumlahBox-" & cNum(x(i, iD1)) & ", JumlahKG=JumlahKG-" & cNum(x(i, id2)) & " where IdStock=" & x(i, icIdStock)
                    If ExecMe(a) = 0 Then GoTo err
                End If
                If CLng(x(i, iOK)) <> 0 Then
                    a = "update m_Stock~ set JumlahBox=JumlahBox+" & cNum(x(i, in1)) & ", JumlahKG=JumlahKG+" & cNum(x(i, in2)) & " where IdStock=" & IdStock
                    If ExecMe(a) = 0 Then GoTo err
                End If
                a = "update t_InputStock~ set NoBukti=" & x(i, iNoNota) & _
                    ", Status=" & IIf(x(i, iOK) = 0, 0, 1) & _
                    ", Tanggal=" & cD(fTanggal) & _
                    ", IdStock=" & IdStock & _
                    ", n1=" & cNum(x(i, in1)) & _
                    ", n2=" & cNum(x(i, in2)) & _
                    ", nDet='" & nDet & _
                    "', Keterangan='" & esc(x(i, iKet)) & _
                    "', PrintedCode=" & PrintedCode & " where TransID=" & x(i, icKey)
                If ExecMe(a) = 0 Then GoTo err
            End If
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    GetResult cD(fTanggal)
    fPrint.Enabled = True
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fTanggal_Validate(Cancel As Boolean)
    GetResult cD(fTanggal)
End Sub

Private Sub TDBDropDown1_DropDownClose()
On Error Resume Next
    col1(iWarna).Value = TDBDropDown1.Columns("Warna").Value
    col1(iKode).Value = TDBDropDown1.Columns("Kode Barang").Value
    col1(iTube).Value = TDBDropDown1.Columns("Tube").Value
    col1(iSatB).Value = TDBDropDown1.Columns("Sat Besar").Value
    col1(iSatK).Value = TDBDropDown1.Columns("Sat Kecil").Value
End Sub

Private Sub TDBDropDown1_Paint()
On Error Resume Next
    TDBGrid1.SelLength = Len(TDBGrid1.Text)
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    If iPrinted = ColIndex Then Exit Sub
    col1(iUpdated).Value = 1
    fPrint.Enabled = False
End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If InStr(col1(iKet).Value, "RETUR DARI") <> 0 Then Cancel = True
End Sub



Private Sub TDBGrid1_BeforeRowColChange(Cancel As Integer)
    If TDBGrid1.AddNewMode = dbgNoAddNew Then
        TDBGrid1.Tag = "OK"
    End If
End Sub

Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
    Response = False
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, iUpdated) = 1 Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Sub SetOtherRowData(ByVal tIdStock As Long)
    a = "select Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil from m_Stock~ where IdStock=" & tIdStock
    query a
    If RS.RecordCount > 0 Then
        col1(iJenis).Value = RS.Fields("Jenis").Value
        col1(iKode).Value = RS.Fields("KodeBarang").Value
        col1(iWarna).Value = RS.Fields("Warna").Value
        col1(iNoWarna).Value = RS.Fields("NoWarna").Value
        col1(iTube).Value = RS.Fields("Tube").Value
        col1(iGrade).Value = RS.Fields("Grade").Value
        col1(iSatB).Value = RS.Fields("SatBesar").Value
        col1(iSatK).Value = RS.Fields("SatKecil").Value
        col1(iUpdated).Value = 1
    End If
    TDBGrid1.SetFocus
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid2, KeyCode
End Sub

Private Sub TDBGrid2_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid2, KeyAscii
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err
    If LastCol = -1 Then Exit Sub
    If LastRow = TDBGrid1.Row Then Exit Sub
    fKet = col1(iJenis).Value & "  " & col1(iKode).Value & "  " & col1(iWarna).Value & "  " & col1(iNoWarna).Value & "  " & col1(iTube).Value & " GRADE " & col1(iGrade).Value
    'TDBGrid2.Update
    If TDBGrid1.Tag = "OK" Then
        TDBGrid1.Tag = ""
        If Not IsNull(LastRow) Then
            AddToCollection x2, TDBGrid1.RowBookmark(LastRow)
        End If
    End If
    GetFromCollection x2, TDBGrid1.Bookmark
    TDBGrid2.Rebind
err:
End Sub

Private Sub AddToCollection(Item As Variant, Key As String)
    On Error GoTo err
        Dim a As New XArrayDB
        a.ReDim 0, 9, 0, 9
        For i = 0 To Item.UpperBound(1)
            For j = 0 To Item.UpperBound(2)
                a(i, j) = Item(i, j)
            Next
        Next
        cArr.Add a, Key
    Exit Sub
err:
    cArr.Remove Key
    cArr.Add a, Key
End Sub

Private Sub GetFromCollection(Item As Variant, Key As String)
    Item.Clear
    
    If CheckCollection(Key) Then
        For i = 0 To Item.UpperBound(1)
            For j = 0 To Item.UpperBound(2)
                Item(i, j) = cArr(Key).Value(i, j)
            Next
        Next
    Else
        If IsEmpty(x(Key, icKey)) Then Exit Sub
        s = "select nDet from t_InputStock~ where TransID=" & x(Key, icKey)
        query s
        If RS.RecordCount = 0 Then Exit Sub
        If IsNull(RS.Fields(0).Value) Then Exit Sub
        s2 = Split(RS.Fields(0).Value, "__")
        p = 0
        q = 0
        For i = 0 To UBound(s2)
           Item(q, p) = s2(i)
           p = p + 1
           If p = 10 Then
                p = 0
                q = q + 1
           End If
           
        Next
    End If
    
End Sub

Private Function CheckCollection(Key As String) As Boolean
    On Error GoTo err
    Dim a As Variant
    CheckCollection = True
    Set a = cArr(Key)
    Exit Function
err:
CheckCollection = False
End Function

Private Sub TDBGrid1_Validate(Cancel As Boolean)
    TDBGrid1.Update
End Sub
Private Sub TDBGrid2_Validate(Cancel As Boolean)
    TDBGrid2.Update
End Sub
Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    Dim a As Double
    Dim a2 As Integer
    a = 0
    a2 = 0
    For i = 0 To x2.UpperBound(1)
        For j = 0 To x2.UpperBound(2)
            If i = TDBGrid2.Bookmark Then tVal = TDBGrid2.Columns(j).Value Else tVal = x2(i, j)
            If Not IsEmpty(tVal) And tVal <> "" And tVal > 0 Then
                a2 = a2 + 1
                a = a + tVal
            End If
        Next
        
    Next
    TDBGrid1.Columns(in1).Value = a2
    TDBGrid1.Columns(in2).Value = a
    col1(iUpdated).Value = 1
    fPrint.Enabled = False
End Sub

Private Sub TDBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If TDBGrid1.AddNewMode <> dbgNoAddNew Then Cancel = True
End Sub

VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormLaporanKeuangan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LAPORAN KEUANGAN"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox fBulan 
      Height          =   315
      ItemData        =   "FormLaporanKeuangan.frx":0000
      Left            =   960
      List            =   "FormLaporanKeuangan.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton fPosting 
      Caption         =   "&POST"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton fCopy 
      Caption         =   "&COPY"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2990
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No Account"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Ket Acc"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Debet"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Kredit"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=4022"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3942"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2011"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1931"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2011"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1931"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
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
      Caption         =   "BUKU BESAR"
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
   Begin UsrText.IText fTahun 
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   600
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4048
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No Account"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Ket Acc"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nilai"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NormalDK"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2593"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2514"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2990"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2910"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2858"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2778"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(3)._MinWidth=128"
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
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
   Begin VB.CommandButton fBukuBesar 
      Caption         =   "&BUKU BESAR"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton fNeraca 
      Caption         =   "&NERACA"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton fRL 
      Caption         =   "&RL"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label fLastPosted 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   2700
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tahun"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "FormLaporanKeuangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim y As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim col2 As TrueOleDBGrid80.Columns
Dim mTipe As Byte
Dim iNoAcc As Integer
Dim iKetAcc As Integer
Dim iNilai As Integer
Dim iNormalDK As Integer
Dim akhir As Long
Dim tipe() As Variant

Private Sub fAkhir_Validate(Cancel As Boolean)
    If cD(fAkhir) > cD(pServerDate) Then fAkhir = pServerDate
End Sub

Private Sub fCopy_Click()
    CopyGrid TDBGrid1
End Sub

Private Sub fNeraca_Click()
    mTipe = 0
    DoQuery
End Sub


Private Sub Form_Load()
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    Set TDBGrid2.Array = y
    Set col2 = TDBGrid2.Columns
    TDBGrid1.AllowUpdate = False
    col1("Nilai").Tag = "Decimal"
    TDBGridLoad TDBGrid1
    a = "select Tipe from m_Tipe"
    query a
    tipe = RS.GetRows
    fBulan.ListIndex = Mid(pServerDate, 4, 2) - 1
    fTahun = "20" & Right(pServerDate, 2)
    fAkhir = pServerDate
    a = "select Tanggal from AutoUpdate where Nama='POSTED'"
    query a
    fLastPosted = "Last Posted: " & cTanggal(RS.Fields(0).Value)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid2.Height = ScaleHeight - TDBGrid2.Top - 100
    TDBGrid2.Width = ScaleWidth - TDBGrid2.Left - 100
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left - 100
End Sub

Private Sub fPosting_Click()
On Error GoTo err
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    CN.BeginTrans
    th = Right(fTahun, 2) & zerofill(fBulan.ListIndex + 1, 2)
    awal = th & "01"
    akhir = th & cari_max(fBulan.ListIndex + 1, Right(fTahun, 2))
    a = "select Tipe, Level1, Level2, Level3, Level4, Level5, Child, m_ChartAccount.NoAccount, m_ChartAccount.Deskripsi, m_ChartAccount.NormalDK from t_LaporanSetting left join m_ChartAccount on t_LaporanSetting.NoAccount=m_ChartAccount.NoAccount order by Level1, Level2, Level3, Level4, Level5"
    rs1.Open a, CN, adOpenKeyset, adLockReadOnly
    If rs1.RecordCount = 0 Then Exit Sub
    'Ambil dari Setting yang Ada
    For i = 0 To rs1.RecordCount - 1
        MFilter = ""
        If rs1.Fields("Level2").Value = 0 Then
            MFilter = " where Level1=" & rs1.Fields("Level1").Value
        ElseIf rs1.Fields("Level3").Value = 0 Then
            MFilter = " where Level1=" & rs1.Fields("Level1").Value & " and Level2=" & rs1.Fields("Level2").Value
        ElseIf rs1.Fields("Level4").Value = 0 Then
            MFilter = " where Level1=" & rs1.Fields("Level1").Value & " and Level2=" & rs1.Fields("Level2").Value & " and Level3=" & rs1.Fields("Level3").Value
        ElseIf rs1.Fields("Level5").Value = 0 Then
            MFilter = " where Level1=" & rs1.Fields("Level1").Value & " and Level2=" & rs1.Fields("Level2").Value & " and Level3=" & rs1.Fields("Level3").Value & " and Level4=" & rs1.Fields("Level4").Value
        Else
            MFilter = " where Level1=" & rs1.Fields("Level1").Value & " and Level2=" & rs1.Fields("Level2").Value & " and Level3=" & rs1.Fields("Level3").Value & " and Level4=" & rs1.Fields("Level4").Value & " and Level5=" & rs1.Fields("Level5").Value
        End If
        a = "select SQLSummary from m_SettingAccounting left join m_ChartAccount on m_SettingAccounting.NoAcc=m_ChartAccount.NoAccount" & MFilter
        If rs2.State = 1 Then rs2.Close
        rs2.Open a, CN, adOpenKeyset, adLockReadOnly
        Dim a1 As String
        For j = 0 To rs2.RecordCount - 1
            a1 = rs2.Fields("SQLSummary").Value
            a1 = Replace(a1, "@Akhir", akhir)
            a1 = Replace(a1, "@Awal", awal)
            a1 = Replace(a1, "!", " between " & awal & " and " & akhir)
            TipeRepeat a1
            query a1
            a = "update t_LaporanSetting set n" & fBulan.ListIndex + 1 & "=" & cNum(RS.Fields(0).Value) & " where NoAccount='" & rs1.Fields("NoAccount").Value & "'"
            CN.Execute a
            rs2.MoveNext
        Next
        a1 = "select sum(iif(NormalDK=DK, Nilai, -Nilai)) from (select Tanggal, 'D' as DK, DebetAcc as NoAcc, Nilai from t_Account union all select Tanggal, 'K', KreditAcc, Nilai from t_Account union all {select Tanggal, 'D', DebetAcc, Nilai from t_STTPembayaran~} union all select Tanggal, 'K', DebetAcc, Nilai from t_PembelianSTTPembayaran) as Temp left join m_ChartAccount on m_ChartAccount.NoAccount=Temp.NoAcc " & MFilter & " and Tanggal!"
        'a = "select sum(n) from (select iif(NormalDK='D',Nilai,-Nilai) as n from t_Account left join m_ChartAccount on t_Account.DebetAcc=m_ChartAccount.NoAccount " & MFilter & " and Tanggal! union all select iif(NormalDK='K', Nilai, -Nilai) from t_Account left join m_ChartAccount on t_Account.KreditAcc=m_ChartAccount.NoAccount " & MFilter & " and Tanggal!)"
        a1 = Replace(a1, "!", " between " & awal & " and " & akhir)
        TipeRepeat a1
        query a1
        If Not IsNull(RS.Fields(0).Value) Then
            a = "update t_LaporanSetting set n" & fBulan.ListIndex + 1 & "=n" & fBulan.ListIndex + 1 & "+" & cNum(RS.Fields(0).Value) & " where NoAccount='" & rs1.Fields("NoAccount").Value & "'"
            CN.Execute a
        Else
            MsgBox "Data ada Pembayaran Belum Ada No Accountnya!!!"
            GoTo err
        End If
        rs1.MoveNext
    Next
    a = "update AutoUpdate set Tanggal=" & cD(pServerDate) & " where Nama='POSTED'"
    CN.Execute a
    CN.CommitTrans
    MsgBox "SUKSES"
    fLastPosted = "Last Posted: " & cTanggal(akhir)
    TDBGridClear TDBGrid1
    Exit Sub
err:
    CN.RollbackTrans
End Sub
Sub TipeRepeat(a As String)
    b = InStr(1, a, "{")
    If b > 0 Then
        c = InStr(b, a, "}")
        d = ""
        Dim kal As String
        kal = Mid(a, b + 1, c - b - 1)
        For k = 0 To UBound(tipe, 2)
            d = d & " union all " & Replace(kal, "~", tipe(0, k))
        Next
        a = Left(a, b - 1) & Mid(d, 12) & Mid(a, c + 1)
    End If
End Sub

Private Sub fRL_Click()
    mTipe = 1
    DoQuery
End Sub

Sub DoQuery()
Dim rs1 As New ADODB.Recordset
    iNoAcc = 0
    iKetAcc = 1
    iNilai = 2
    iNormalDK = 3
    a = "select Level1, Level2, Level3, Level4, Level5, n" & fBulan.ListIndex + 1 & " as n, m_ChartAccount.NoAccount, m_ChartAccount.Deskripsi, NormalDK from t_LaporanSetting left join m_ChartAccount on t_LaporanSetting.NoAccount=m_ChartAccount.NoAccount where Tipe=" & mTipe & " order by Level1, Level2, Level3, Level4, Level5"
    query a
    x.ReDim 0, RS.RecordCount - 1, 0, col1.Count - 1
    For i = 0 To RS.RecordCount - 1
        x(i, iNoAcc) = RS.Fields("NoAccount").Value
        x(i, iKetAcc) = RS.Fields("Deskripsi").Value
        x(i, iNormalDK) = RS.Fields("NormalDK").Value
        x(i, iNilai) = RS.Fields("n").Value
        RS.MoveNext
    Next
    TDBGrid1.Rebind
End Sub


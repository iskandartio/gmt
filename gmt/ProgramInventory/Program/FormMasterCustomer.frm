VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMasterCustomer 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CUSTOMER"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Tag             =   "21"
   WindowState     =   2  'Maximized
   Begin VB.TextBox fPenerima 
      Height          =   285
      Left            =   5880
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton fTambah 
      Caption         =   "&TAMBAH"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton fTambahPenerima 
      Caption         =   "TAMBAH"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton fUpdatePenerima 
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin UsrText.IText fFind 
      Height          =   270
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
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
      Columns(1).Caption=   "Active"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kode"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nama"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Alamat"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Alamat Pendek"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Limit"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Kota"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Telepon"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Fax"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Contact"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Tempo"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Mata Uang"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1032"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=953"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1111"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1032"
      Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=4180"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=4101"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(31)=   "Column(7).Width=1799"
      Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=1720"
      Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(35)=   "Column(8).Width=1588"
      Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=1508"
      Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(39)=   "Column(9).Width=1667"
      Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=1588"
      Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(43)=   "Column(10).Width=1693"
      Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=1614"
      Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(47)=   "Column(11).Width=1085"
      Splits(0)._ColumnProps(48)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(11)._WidthInPix=1005"
      Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(51)=   "Column(12).Width=1588"
      Splits(0)._ColumnProps(52)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(12)._WidthInPix=1508"
      Splits(0)._ColumnProps(54)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(55)=   "Column(12)._MinWidth=-3"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=61,.parent=2,.namedParent=63"
      _StyleDefs(17)  =   "FilterBarStyle:id=64,.parent=1,.namedParent=66"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=62,.parent=61"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=65,.parent=64"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=74,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=82,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=24,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=12,.alignment=3"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=13,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=40,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=37,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=38,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=39,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=78,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=86,.parent=11,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=56,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=53,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=54,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=55,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=44,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=41,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=42,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=43,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=48,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=45,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=46,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=47,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=52,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=49,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=50,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=51,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=60,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=57,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=58,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=59,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=70,.parent=11"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=67,.parent=12"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=68,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=69,.parent=15"
      _StyleDefs(82)  =   "Named:id=29:Normal"
      _StyleDefs(83)  =   ":id=29,.parent=0"
      _StyleDefs(84)  =   "Named:id=30:Heading"
      _StyleDefs(85)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=30,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=31:Footing"
      _StyleDefs(88)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=32:Selected"
      _StyleDefs(90)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=33:Caption"
      _StyleDefs(92)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(93)  =   "Named:id=34:HighlightRow"
      _StyleDefs(94)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(95)  =   "Named:id=35:EvenRow"
      _StyleDefs(96)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=36:OddRow"
      _StyleDefs(98)  =   ":id=36,.parent=29"
      _StyleDefs(99)  =   "Named:id=63:RecordSelector"
      _StyleDefs(100) =   ":id=63,.parent=30"
      _StyleDefs(101) =   "Named:id=66:FilterBar"
      _StyleDefs(102) =   ":id=66,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2990
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
      Columns(1).Caption=   "NoUrut"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nama"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Alamat"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Alamat Pendek"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Telepon"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Fax"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Contact"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1032"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=4180"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=4101"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1588"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1508"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=1667"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=1588"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=1693"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=1614"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=61,.parent=2,.namedParent=63"
      _StyleDefs(17)  =   "FilterBarStyle:id=64,.parent=1,.namedParent=66"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=62,.parent=61"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=65,.parent=64"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=3"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13,.alignment=3"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=74,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=40,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=37,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=38,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=39,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=78,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=44,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=41,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=42,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=43,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=48,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=45,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=46,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=47,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=52,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=49,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=50,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=51,.parent=15"
      _StyleDefs(62)  =   "Named:id=29:Normal"
      _StyleDefs(63)  =   ":id=29,.parent=0"
      _StyleDefs(64)  =   "Named:id=30:Heading"
      _StyleDefs(65)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   ":id=30,.wraptext=-1"
      _StyleDefs(67)  =   "Named:id=31:Footing"
      _StyleDefs(68)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   "Named:id=32:Selected"
      _StyleDefs(70)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=33:Caption"
      _StyleDefs(72)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(73)  =   "Named:id=34:HighlightRow"
      _StyleDefs(74)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(75)  =   "Named:id=35:EvenRow"
      _StyleDefs(76)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(77)  =   "Named:id=36:OddRow"
      _StyleDefs(78)  =   ":id=36,.parent=29"
      _StyleDefs(79)  =   "Named:id=63:RecordSelector"
      _StyleDefs(80)  =   ":id=63,.parent=30"
      _StyleDefs(81)  =   "Named:id=66:FilterBar"
      _StyleDefs(82)  =   ":id=66,.parent=29"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Nama Penerima"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PENERIMA"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
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
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Find"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FormMasterCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim y As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim col2 As TrueOleDBGrid80.Columns

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub fFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        fFind.Cancel = True
        Dim cCol As Integer
        cCol = TDBGrid1.Col
        n = TDBGrid1.Bookmark
        m = n + 1
        If m = x.UpperBound(1) + 1 Then m = 0
        Do While m <> n
            If InStr(1, x(m, cCol), fFind, vbTextCompare) <> 0 Then Exit Do
            m = m + 1
            If m = x.UpperBound(1) + 1 Then m = 0
        Loop
        If InStr(1, x(m, cCol), fFind, vbTextCompare) = 0 Then MsgBox "Tidak Ketemu"
        TDBGrid1.Bookmark = m
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Caption = Caption & "---" & pTipe
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    Set TDBGrid2.Array = y
    Set col2 = TDBGrid2.Columns
    col1("updated").Visible = False
    col2("updated").Visible = False
    col2("NoUrut").Visible = False
    col1("Active").Tag = "OK?"
    TDBGridLoad TDBGrid1
    TDBGrid1.FetchRowStyle = True
    TDBGrid2.FetchRowStyle = True
    DoQuery
    GetDetail
End Sub

Private Sub DoQuery()
On Error GoTo err
    a = "select 0, IsActive, Kode, Nama, Alamat, AlamatPendek, Limit, Kota, Telepon, Fax, ContactPerson, WaktuPembayaran, CusCurrency from m_Customer order by Nama"
    query a
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
err:
End Sub

Private Sub GetDetail()
On Error GoTo err
    y.ReDim 0, 0, 0, TDBGrid2.Columns.Count - 1
    y.DeleteRows 0
    If col1("Kode").Value = "" Then GoTo err
    a = "select 0, NoUrut, Nama, Alamat, AlamatPendek, Telepon, Fax, Contact from m_Penerima where Kode=" & col1("Kode").Value & " order by NoUrut"
    query a
    If RS.RecordCount > 0 Then y.LoadRows RS.GetRows
err:
    TDBGrid2.Rebind
End Sub

Private Sub Form_Resize()
On Error Resume Next
    fKet.Width = ScaleWidth - fKet.Left - 150
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid2.Width = TDBGrid1.Width
    TDBGrid2.Height = ScaleHeight - TDBGrid2.Top - 200
End Sub

Private Sub fPenerima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        fFind = ""
        a = "select Kode from m_penerima where Nama like '" & esc(fPenerima) & "'"
        query a
        If RS.RecordCount = 0 Then Exit Sub
        TDBGrid1.Col = col1("Kode").ColIndex
        fFind = RS.Fields(0).Value
        For i = 0 To x.UpperBound(1)
            If StrComp(fFind, x(i, col1("Kode").ColIndex), vbTextCompare) = 0 Then
                TDBGrid1.Bookmark = i
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub fTambah_Click()
On Error GoTo err
    TDBGrid1.Update
    x.AppendRows
    x(x.UpperBound(1), TDBGrid1.Columns("Tempo").ColIndex) = "30"
    x(x.UpperBound(1), TDBGrid1.Columns("Mata Uang").ColIndex) = "RP"
    TDBGrid1.Rebind
    TDBGrid1.MoveLast
    TDBGrid1.SetFocus
    GetDetail
err:
End Sub

Private Sub fTambahPenerima_Click()
On Error GoTo err
    TDBGrid2.Update
    If y.UpperBound(1) = -1 Then
        y.AppendRows
        y(y.UpperBound(1), col2("updated").ColIndex) = "1"
        y(y.UpperBound(1), col2("Nama").ColIndex) = col1("Nama").Value
        y(y.UpperBound(1), col2("Alamat").ColIndex) = col1("Alamat").Value
        y(y.UpperBound(1), col2("Alamat Pendek").ColIndex) = col1("Alamat Pendek").Value
        y(y.UpperBound(1), col2("Telepon").ColIndex) = col1("Telepon").Value
        y(y.UpperBound(1), col2("Fax").ColIndex) = col1("Fax").Value
        y(y.UpperBound(1), col2("Contact").ColIndex) = col1("Contact").Value
    Else
        y.AppendRows
    End If
    TDBGrid2.Rebind
    TDBGrid2.MoveLast
    TDBGrid2.SetFocus
err:
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    col1("updated").Value = 1
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    x.QuickSort 0, x.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_NUMBER
    TDBGrid1.Rebind
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        fFind_KeyDown 13, 0
    End If
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If LastRow <> TDBGrid1.Bookmark Then
        GetDetail
    End If
    fKet = TDBGrid1.Columns(TDBGrid1.Col).Value
End Sub

Private Sub TDBGrid2_AfterColEdit(ByVal ColIndex As Integer)
    col2("updated").Value = 1
End Sub

Private Sub TDBGrid2_BeforeDelete(Cancel As Integer)
On Error GoTo err
    If MsgBox("Yakin mau Hapus?", vbYesNo) = vbNo Then GoTo err
    BeginTransaction
    a = "delete from m_penerima where Kode=" & TDBGrid1.Columns("Kode").Value & " and NoUrut=" & y(TDBGrid2.Bookmark, col2("NoUrut").ColIndex)
    ExecMe a
    a = "update m_Penerima set NoUrut=NoUrut-1 where Kode=" & TDBGrid1.Columns("Kode").Value & " and NoUrut>" & y(TDBGrid2.Bookmark, col2("NoUrut").ColIndex)
    ExecMe a
    CommitTransaction
    Exit Sub
err:
    RollBackTransaction
    Cancel = True
End Sub

Private Sub fUpdatePenerima_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid2.Update
    For i = 0 To y.UpperBound(1)
        If y(i, col2("updated").ColIndex) <> "0" Then
            If y(i, col2("NoUrut").ColIndex) <> "" Then
                a = "update m_Penerima set Kode=" & col1("Kode").Value & _
                    ", Nama='" & y(i, col2("Nama").ColIndex) & _
                    "', Alamat='" & y(i, col2("Alamat").ColIndex) & _
                    "', AlamatPendek='" & y(i, col2("Alamat Pendek").ColIndex) & _
                    "', Telepon='" & y(i, col2("Telepon").ColIndex) & _
                    "', Fax='" & y(i, col2("Fax").ColIndex) & _
                    "', Contact='" & y(i, col2("Contact").ColIndex) & _
                    "' where Kode=" & col1("Kode").Value & " and NoUrut=" & i
                If ExecMe(a) = 0 Then GoTo err
            Else
                a = "insert into m_Penerima(Kode, Nama, Alamat, AlamatPendek, Telepon, Fax, Contact, NoUrut) values(" & col1("Kode").Value & _
                    ",'" & y(i, col2("Nama").ColIndex) & _
                    "','" & y(i, col2("Alamat").ColIndex) & _
                    "','" & y(i, col2("Alamat Pendek").ColIndex) & _
                    "','" & y(i, col2("Telepon").ColIndex) & _
                    "','" & y(i, col2("Fax").ColIndex) & _
                    "','" & y(i, col2("Contact").ColIndex) & _
                    "'," & i & ")"
                If ExecMe(a) = 0 Then GoTo err
                y(i, col2("NoUrut").ColIndex) = i
            End If
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    GetDetail
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If CLng(x(i, col1("updated").ColIndex)) <> 0 Then
            
            If x(i, col1("Kode").ColIndex) <> "" Then
                a = "select top 1 NoSPP from t_SPP~ where Kode=" & x(i, col1("Kode").ColIndex)
                query a
                If RS.RecordCount > 0 Then
                    b = MsgBox("Kode Customer " & x(i, col1("Kode").ColIndex) & " sudah ada di Penjualan, Teruskan?", vbYesNo)
                    If b = vbNo Then GoTo err
                End If
                a = "update m_Customer set Kode=" & x(i, col1("Kode").ColIndex) & _
                    ", Nama='" & x(i, col1("Nama").ColIndex) & _
                    "', IsActive=" & IIf(x(i, col1("Active").ColIndex) = 0, 0, 1) & _
                    ", Alamat='" & x(i, col1("Alamat").ColIndex) & _
                    "', AlamatPendek='" & x(i, col1("Alamat Pendek").ColIndex) & _
                    "', Limit='" & cNum(x(i, col1("Limit").ColIndex)) & _
                    "', Kota='" & x(i, col1("Kota").ColIndex) & _
                    "', Telepon='" & x(i, col1("Telepon").ColIndex) & _
                    "', Fax='" & x(i, col1("Fax").ColIndex) & _
                    "', ContactPerson='" & x(i, col1("Contact").ColIndex) & _
                    "', WaktuPembayaran='" & x(i, col1("Tempo").ColIndex) & _
                    "', CusCurrency='" & x(i, col1("Mata Uang").ColIndex) & _
                    "' where Kode=" & x(i, col1("Kode").ColIndex)
            Else
                a = "select max(Kode)+1 from m_Customer"
                query a
                Kode = IIf(IsNull(RS.Fields(0).Value), 1, RS.Fields(0).Value)
                a = "insert into m_Customer(Kode, IsActive, Nama, Alamat, AlamatPendek, Limit, Kota, Telepon, Fax, ContactPerson, WaktuPembayaran, CusCurrency) values(" & _
                    Kode & _
                    ",1" & _
                    ",'" & x(i, col1("Nama").ColIndex) & _
                    "','" & x(i, col1("Alamat").ColIndex) & _
                    "','" & x(i, col1("Alamat Pendek").ColIndex) & _
                    "','" & cNum(x(i, col1("Limit").ColIndex)) & _
                    "','" & x(i, col1("Kota").ColIndex) & _
                    "','" & x(i, col1("Telepon").ColIndex) & _
                    "','" & x(i, col1("Fax").ColIndex) & _
                    "','" & x(i, col1("Contact").ColIndex) & _
                    "','" & x(i, col1("Tempo").ColIndex) & _
                    "','" & x(i, col1("Mata Uang").ColIndex) & "')"
                x(i, col1("Kode").ColIndex) = Kode
            End If
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

Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If y(Bookmark, col2("updated").ColIndex) = "1" Then RowStyle.BackColor = vbYellow
End Sub
Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("updated").ColIndex) = "1" Then RowStyle.BackColor = vbYellow
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    TDBGridKeyDown TDBGrid2, KeyCode
End Sub

Private Sub TDBGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ActiveControl.Name <> "TDBGrid2" Then
        TDBGrid2.SetFocus
    End If
End Sub
Private Sub TDBGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ActiveControl.Name <> "TDBGrid1" Then
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If ActiveControl = TDBGrid2 Then
        fKet = TDBGrid2.Columns(TDBGrid2.Col).Value
    End If
End Sub


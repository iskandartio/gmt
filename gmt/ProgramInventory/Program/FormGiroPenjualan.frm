VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormGiroPenjualan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "GIRO PENJUALAN"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   11145
   Tag             =   "5"
   WindowState     =   2  'Maximized
   Begin VB.ComboBox fGiroOK 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FormGiroPenjualan.frx":0000
      Left            =   7800
      List            =   "FormGiroPenjualan.frx":000D
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton fGO 
      Caption         =   "&SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton fPost 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin UsrText.IText fJatuhTempo 
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      Text            =   "__/__/__"
      DataType        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8493
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "updated"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "No Giro"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Bank"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tanggal Tempo"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Edit Mask"
      Columns(3).EditMask=   "##/##/##"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tanggal Setor"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Edit Mask"
      Columns(4).EditMask=   "##/##/##"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Bank Setor"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Tanggal Cair"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "Edit Mask"
      Columns(6).EditMask=   "##/##/##"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Nilai Giro"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "No STT"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Customer"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Keterangan"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=714"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=635"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=529"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0)._MinWidth=29557"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1931"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1852"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8724"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(1)._MinWidth=29557"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1984"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1905"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8724"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(2)._MinWidth=115"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2090"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2011"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=532"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(3)._MinWidth=73868652"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2037"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1958"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=532"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(4)._MinWidth=-1"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=532"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(5)._MinWidth=-1"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1693"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=16"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(6)._MinWidth=-1"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2566"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2487"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=8724"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=1455"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1376"
      Splits(0)._ColumnProps(52)=   "Column(8)._ColStyle=532"
      Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(54)=   "Column(9).Width=2990"
      Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=2910"
      Splits(0)._ColumnProps(57)=   "Column(9)._ColStyle=532"
      Splits(0)._ColumnProps(58)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(59)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(10)._ColStyle=532"
      Splits(0)._ColumnProps(63)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
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
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Times New Roman"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=69,.parent=2,.namedParent=71"
      _StyleDefs(19)  =   "FilterBarStyle:id=72,.parent=1,.namedParent=74"
      _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(23)  =   ":id=12,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(24)  =   ":id=12,.fontname=Times New Roman"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=70,.parent=69"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=73,.parent=72"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=82,.parent=11,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=24,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13,.alignment=3"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=40,.parent=11,.alignment=3,.locked=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=37,.parent=12,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=38,.parent=13,.alignment=3"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=39,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=56,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=53,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=55,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=68,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=65,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=66,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=67,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=60,.parent=11,.alignment=0,.locked=0"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=57,.parent=12,.alignment=3"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=58,.parent=13,.alignment=3"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=59,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=44,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=41,.parent=12,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=42,.parent=13,.alignment=3"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=43,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=64,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=61,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=62,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=63,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=48,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=45,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=46,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=47,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=52,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=49,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=50,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=51,.parent=15"
      _StyleDefs(78)  =   "Named:id=29:Normal"
      _StyleDefs(79)  =   ":id=29,.parent=0,.valignment=2"
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
      _StyleDefs(95)  =   "Named:id=71:RecordSelector"
      _StyleDefs(96)  =   ":id=71,.parent=30"
      _StyleDefs(97)  =   "Named:id=74:FilterBar"
      _StyleDefs(98)  =   ":id=74,.parent=29"
   End
   Begin UsrText.IText fJatuhTempo2 
      Height          =   330
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      Text            =   "__/__/__"
      DataType        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "GIRO CUSTOMER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Tempo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "FormGiroPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim StatusSeb() As Integer

Private Sub fGiroOK_Click()
    fGO_Click

End Sub


Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Function MyFilter() As String
    If cD(fJatuhTempo) <> "A" Then MyFilter = " and TanggalGiro between " & cD(fJatuhTempo) & " and " & cD(fJatuhTempo2.Text)
    If fGiroOK.Text = "NOT OK" Then
        MyFilter = MyFilter & " and TanggalCair=991231"
    ElseIf fGiroOK.Text = "OK" Then
        MyFilter = MyFilter & " and TanggalCair<991231"
    End If
End Function

Private Sub fGiroOK_Validate(Cancel As Boolean)
    fGO_Click
    
End Sub

Private Sub fGO_Click()
On Error GoTo err
    
    a = "select 0, NoGiro, NamaBank, TanggalGiro, TanggalSetor, BankSetor, TanggalCair, Nilai, NoSTT, KetCustomer, Keterangan  from t_STTPembayaran" & pTipe & " where CaraBayar='GIRO'" & MyFilter() & " order by TanggalGiro"
    query a
    x.ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then
        x.LoadRows RS.GetRows
        ReDim StatusSeb(x.UpperBound(1))
        For i = 0 To x.UpperBound(1)
            x(i, 3) = cTanggal2(x(i, 3))
            x(i, 4) = cTanggal2(x(i, 4))
            x(i, 6) = cTanggal2(x(i, 6))
            If x(i, 0) = 0 Then x(i, 0) = 0 Else x(i, 0) = -1
            StatusSeb(i) = x(i, 0)
        Next
    End If
err:
    TDBGrid1.Rebind
End Sub

Private Sub fJatuhTempo_KeyDown(KeyCode As Integer, Shift As Integer)
    fJatuhTempo.Cancel = True
    If KeyCode = 13 Then fGO_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    TDBGrid1.Columns("updated").value = "1"
End Sub

Private Sub fGiroOK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fGO_Click
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    TDBGrid1.Columns("Nilai Giro").NumberFormat = "FormatText Event"
    TDBGrid1.Columns("Nilai Giro").Alignment = dbgRight
    For i = 0 To TDBGrid1.Columns.count - 1
        TDBGrid1.Columns(i).Locked = True
    Next
    TDBGrid1.Columns("Tanggal Cair").Locked = False
    TDBGrid1.Columns("Tanggal Setor").Locked = False
    TDBGrid1.Columns("Bank Setor").Locked = False
    Set TDBGrid1.Array = x
    fGiroOK.ListIndex = 0
    fJatuhTempo2.Text = pServerDate
    fJatuhTempo.Text = cTanggal((cD(fJatuhTempo2.Text) \ 100) * 100 + 1)
    fGO_Click
End Sub

Private Sub TDBGrid1_GotFocus()
On Error Resume Next
    If x.UpperBound(1) = -1 Then TDBGrid1.Row = 0
End Sub

Private Sub fPost_Click()
'On Error GoTo err
    If Not cekValid("EDIT", Tag) Then Exit Sub
    TDBGrid1.Update
    BeginTransaction
    For i = 0 To x.UpperBound(1)
        If x(i, 0) = "1" Then
            a = "update t_STTPembayaran~ set BankSetor='" & esc(x(i, 5)) & "', TanggalGiro=" & cD(x(i, 3)) & ", TanggalSetor=" & cD(x(i, 4)) & ", TanggalCair=" & cD(x(i, 6)) & " where NoSTT='" & esc(x(i, 8)) & "' and NoGiro='" & esc(x(i, 1)) & "'"
            b = ExecMe(a)
            a = "update t_STTPelunasan~ set TanggalPelunasan=" & cD(x(i, 6)) & " where NoSTT='" & esc(x(i, 8)) & "' and NoGiro='" & esc(x(i, 1)) & "'"
            b = ExecMe(a)
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    fGO_Click
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, value As Variant, Bookmark As Variant)
    a = TDBGrid1.Columns(ColIndex).Caption
    If a = "Tanggal Tempo" Then
        value = cTanggal(value)
    ElseIf a = "Nilai Giro" Then
        value = cDecimal(value)
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    If KeyCode = 13 Then
        Dim cCol As Integer
        Dim cRow As Integer
        KeyCode = 0
        With TDBGrid1
        a = .Columns(.Col).Caption
        cCol = .Col
        cRow = .Row
        Do
            If cCol = TDBGrid1.Columns.count - 1 Then
                cRow = cRow + 1
                cCol = 0
            Else
                cCol = cCol + 1
            End If
            If Not .Columns(cCol).Locked And .Columns(cCol).Visible Then Exit Do
        Loop
        TDBGrid1.Row = cRow
        TDBGrid1.Col = cCol
        End With
    ElseIf KeyCode = 117 Then
        TDBGrid1.Font.Size = TDBGrid1.Font.Size + 2
        FontSize = TDBGrid1.Font.Size
        TDBGrid1.RowHeight = TextHeight("I") + 50
        KeyCode = 0
    ElseIf KeyCode = 116 Then
        KeyCode = 0
        If TDBGrid1.Font.Size <= 6 Then Exit Sub
        TDBGrid1.Font.Size = TDBGrid1.Font.Size - 2
        FontSize = TDBGrid1.Font.Size
        TDBGrid1.RowHeight = TextHeight("I") + 50
    End If
err:
End Sub

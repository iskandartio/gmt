VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{BD09B73E-A5EF-4CAB-A002-921A8335B40E}#1.0#0"; "usrtruecombo.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormPelunasanPenjualan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PELUNASAN PENJUALAN"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Tag             =   "4"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   3480
      TabIndex        =   41
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   3120
      TabIndex        =   40
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   2760
      TabIndex        =   39
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   2400
      TabIndex        =   38
      Top             =   120
      Width           =   375
   End
   Begin UsrText.IText fQuick 
      Height          =   270
      Left            =   1080
      TabIndex        =   35
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
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
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
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
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   10800
      TabIndex        =   28
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   11655
      Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
         Height          =   2775
         Left            =   2520
         TabIndex        =   44
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4895
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No Acc"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Ket Acc"
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
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=64"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=33"
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
      Begin VB.CommandButton fKalkulasi 
         Caption         =   "&KALKULASI"
         Height          =   375
         Left            =   10080
         TabIndex        =   31
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox fKeterangan 
         Height          =   285
         Left            =   5280
         TabIndex        =   29
         Top             =   6480
         Width           =   3495
      End
      Begin VB.TextBox fPembulatanPenjualan 
         Height          =   285
         Left            =   3240
         TabIndex        =   26
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox fKredit 
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox fDebet 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox fTotalPembayaran 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox fTotalPelunasan 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox fTotalPotongan 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5880
         Width           =   1815
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2355
         _LayoutType     =   4
         _RowHeight      =   14
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Cara Bayar"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Bank"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No Giro"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tanggal Lunas"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "Edit Mask"
         Columns(3).EditMask=   "##/##/##"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nilai"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nilai RP"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "No Acc"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Ket Acc"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Keterangan"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "IdDP"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=532"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2540"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2461"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=532"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2487"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2408"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=532"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1984"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1905"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=532"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2514"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2434"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=532"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=532"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2223"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2143"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=532"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3545"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3466"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=532"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=4868"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=4789"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=532"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=532"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(51)=   "Column(9)._MinWidth=82263952"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483631
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=65,.parent=2,.namedParent=67"
         _StyleDefs(19)  =   "FilterBarStyle:id=68,.parent=1,.namedParent=70"
         _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.bgcolor=&HE8FDFF&,.fgcolor=&H2878A&"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=13,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=15,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=66,.parent=65"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=69,.parent=68"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=40,.parent=11"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=44,.parent=11"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=41,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=42,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=43,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=48,.parent=11"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=45,.parent=12"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=46,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=47,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=56,.parent=11"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=53,.parent=12"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=54,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=55,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=64,.parent=11"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=61,.parent=12"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=62,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=63,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=74,.parent=11"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=12"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=52,.parent=11"
         _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=49,.parent=12"
         _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=50,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=51,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=60,.parent=11"
         _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=57,.parent=12"
         _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=58,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=59,.parent=15"
         _StyleDefs(72)  =   "Named:id=29:Normal"
         _StyleDefs(73)  =   ":id=29,.parent=0,.valignment=2"
         _StyleDefs(74)  =   "Named:id=30:Heading"
         _StyleDefs(75)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=30,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=31:Footing"
         _StyleDefs(78)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=32:Selected"
         _StyleDefs(80)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=33:Caption"
         _StyleDefs(82)  =   ":id=33,.parent=30,.alignment=2"
         _StyleDefs(83)  =   "Named:id=34:HighlightRow"
         _StyleDefs(84)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(85)  =   "Named:id=35:EvenRow"
         _StyleDefs(86)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=36:OddRow"
         _StyleDefs(88)  =   ":id=36,.parent=29"
         _StyleDefs(89)  =   "Named:id=67:RecordSelector"
         _StyleDefs(90)  =   ":id=67,.parent=30"
         _StyleDefs(91)  =   "Named:id=70:FilterBar"
         _StyleDefs(92)  =   ":id=70,.parent=29"
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2778
         _LayoutType     =   4
         _RowHeight      =   14
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No SJ"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Tanggal"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nilai Faktur"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Lunas"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Sisa"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Bayar"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Bayar RP"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Keterangan"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "MataUang"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2884"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2805"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8724"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1667"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1588"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8724"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2778"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2699"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8724"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2275"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2196"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8724"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2566"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2487"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8724"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2328"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2249"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=532"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=532"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=532"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=532"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483631
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=81,.parent=2,.namedParent=83"
         _StyleDefs(19)  =   "FilterBarStyle:id=84,.parent=1,.namedParent=86"
         _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1,.bgcolor=&H80000005&"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.bgcolor=&HE8FDFF&,.fgcolor=&H828C&"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=13,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=15,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=82,.parent=81"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=85,.parent=84"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=3,.locked=-1"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=2"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13,.alignment=3"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=76,.parent=11,.alignment=3,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=73,.parent=12,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=74,.parent=13,.alignment=3"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=75,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=40,.parent=11,.alignment=3,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13,.alignment=3"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=44,.parent=11,.alignment=3,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=41,.parent=12,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=42,.parent=13,.alignment=3"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=43,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=60,.parent=11,.alignment=3,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=57,.parent=12,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=58,.parent=13,.alignment=3"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=59,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=64,.parent=11,.bgcolor=&HFFFDDF&"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=61,.parent=12"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=62,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=63,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=28,.parent=11"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=12"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=68,.parent=11"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=65,.parent=12"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=66,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=67,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=80,.parent=11"
         _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=77,.parent=12"
         _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=78,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=79,.parent=15"
         _StyleDefs(68)  =   "Named:id=29:Normal"
         _StyleDefs(69)  =   ":id=29,.parent=0,.valignment=2"
         _StyleDefs(70)  =   "Named:id=30:Heading"
         _StyleDefs(71)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=30,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=31:Footing"
         _StyleDefs(74)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=32:Selected"
         _StyleDefs(76)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=33:Caption"
         _StyleDefs(78)  =   ":id=33,.parent=30,.alignment=2"
         _StyleDefs(79)  =   "Named:id=34:HighlightRow"
         _StyleDefs(80)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(81)  =   "Named:id=35:EvenRow"
         _StyleDefs(82)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=36:OddRow"
         _StyleDefs(84)  =   ":id=36,.parent=29"
         _StyleDefs(85)  =   "Named:id=83:RecordSelector"
         _StyleDefs(86)  =   ":id=83,.parent=30"
         _StyleDefs(87)  =   "Named:id=86:FilterBar"
         _StyleDefs(88)  =   ":id=86,.parent=29"
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid3 
         Height          =   1095
         Left            =   120
         TabIndex        =   32
         Top             =   4680
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   1931
         _LayoutType     =   4
         _RowHeight      =   13
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No Bukti"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Tanggal"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nilai Retur"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Sdh Potong"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Sisa"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nilai"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Nilai RP"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Keterangan"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2672"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2593"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8724"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=8"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1879"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1799"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8724"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=76875508"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=532"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=82522080"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2302"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2223"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=532"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(3)._MinWidth=82522080"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=532"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(4)._MinWidth=82522080"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2223"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2143"
         Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=532"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(5)._MinWidth=82522080"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2778"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2699"
         Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=532"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(42)=   "Column(7).Width=2990"
         Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2910"
         Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=532"
         Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483631
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=61,.parent=2,.namedParent=63"
         _StyleDefs(19)  =   "FilterBarStyle:id=64,.parent=1,.namedParent=66"
         _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.bgcolor=&HE1FFFE&,.fgcolor=&H828C&"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=13,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=15,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=62,.parent=61"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=65,.parent=64"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=3,.locked=-1"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=2"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13,.alignment=3"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=60,.parent=11,.alignment=3,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=57,.parent=12,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=58,.parent=13,.alignment=3"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=59,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=70,.parent=11"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=74,.parent=11"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=56,.parent=11"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=53,.parent=12"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=54,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=55,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=40,.parent=11,.alignment=3,.locked=0"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=37,.parent=12,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=38,.parent=13,.alignment=3"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=39,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=28,.parent=11"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=12"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=44,.parent=11"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=41,.parent=12"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=42,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=43,.parent=15"
         _StyleDefs(64)  =   "Named:id=29:Normal"
         _StyleDefs(65)  =   ":id=29,.parent=0,.valignment=2"
         _StyleDefs(66)  =   "Named:id=30:Heading"
         _StyleDefs(67)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=30,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=31:Footing"
         _StyleDefs(70)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=32:Selected"
         _StyleDefs(72)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=33:Caption"
         _StyleDefs(74)  =   ":id=33,.parent=30,.alignment=2"
         _StyleDefs(75)  =   "Named:id=34:HighlightRow"
         _StyleDefs(76)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(77)  =   "Named:id=35:EvenRow"
         _StyleDefs(78)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=36:OddRow"
         _StyleDefs(80)  =   ":id=36,.parent=29"
         _StyleDefs(81)  =   "Named:id=63:RecordSelector"
         _StyleDefs(82)  =   ":id=63,.parent=30"
         _StyleDefs(83)  =   "Named:id=66:FilterBar"
         _StyleDefs(84)  =   ":id=66,.parent=29"
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "KETERANGAN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   30
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BEBAN PEMBULATAN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "KREDIT"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DEBET"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   6240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PEMBAYARAN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PEMBAYARAN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PELUNASAN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PELUNASAN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL POTONGAN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "RETUR DAN POTONGAN HARGA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   2775
      End
   End
   Begin VB.CommandButton fList 
      Caption         =   "&LIST"
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton fNew 
      Caption         =   "&NEW"
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   3720
      TabIndex        =   1
      Top             =   600
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
   Begin UsrText.IText fNoSTT 
      Height          =   270
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
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
   Begin UsrText.IText fMataUang 
      Height          =   270
      Left            =   6120
      TabIndex        =   36
      Top             =   600
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
      Enabled         =   0   'False
   End
   Begin UsrText.IText fKurs 
      Height          =   270
      Left            =   6480
      TabIndex        =   42
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "KURS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   43
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MATA UANG"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   37
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Quick &Find"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PELUNASAN PENJUALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NO PELUNASAN"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "FormPelunasanPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x(2) As New XArrayDB
Dim y As New XArrayDB
Dim LCustomer As Boolean
Dim col1 As TrueOleDBGrid80.Columns
Dim col2 As TrueOleDBGrid80.Columns
Dim col3 As TrueOleDBGrid80.Columns
Dim LNoAccount As Boolean
Dim z As New XArrayDB
Dim mChange As Boolean

Private Sub fFirst_Click()
    ShortNo 1
End Sub

Private Sub fNoSTT_LostFocus()
    b = zerofill(Left(fNoSTT, 5), 5) & Right(fTanggal, 3)
    fNoSTT = b
End Sub

Private Sub fPrev_Click()
On Error Resume Next
    a = Left(fQuick, 5) - 1
    If a < 1 Then a = 1
    ShortNo a
End Sub

Private Sub fNext_Click()
On Error Resume Next
    a = Left(fQuick, 5) + 1
    ShortNo a
End Sub

Private Sub fLast_Click()
On Error Resume Next
    a = Right(fQuick, 3)
    If Left(a, 1) = "/" Then
        b = Right(a, 2) * 10000
    Else
        b = Right(pAddNo, 2) * 10000
    End If
    a = "select max(NoSTT) from t_STT~ where Tanggal>=" & b & " and Tanggal <= " & b + 10000
    query a
    If Not IsNull(RS.Fields(0).value) Then
        a = Left(RS.Fields(0).value, 5)
        ShortNo a
    End If
End Sub

Private Sub ShortNo(ByVal tNo As Long)
    a = Right(fQuick, 3)
    If Left(a, 1) = "/" Then
        b = zerofill(tNo, 5) & a
    Else
        b = zerofill(tNo, 5) & pAddNo
    End If
    fQuick = b
    GetResult fQuick
End Sub

Private Sub fQuick_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then
        ShortNo Left(fQuick, 5)
        fQuick.Cancel = True
        fQuick.FocusSelect
    End If
End Sub

Private Sub fCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not LCustomer Then
        a = "select distinct Nama, Kota, MataUang, m_customer.Kode from t_SPP" & pTipe & " left join m_customer on t_SPP" & pTipe & ".Kode=m_customer.Kode where Total>Pelunasan order by Nama"
        query a
Dim rs1() As Variant
        rs1 = RS.GetRows
        fCustomer.SetDB rs1
        fCustomer.SetType "String"
        LCustomer = True
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub


Private Sub fCustomer_LostFocus()
On Error GoTo err
    If Not fCustomer.Enabled Then Exit Sub
    If fCustomer.GetData("Kode") = "" Then Exit Sub
    fMataUang = fCustomer.GetData("Mata Uang")
    a = "select 'DP','','',Tanggal, NilaiDP-Terpakai as Nilai, 0, '', '', NoSC+' '+Keterangan, IdDP from t_DP where KodeCustomer=" & fCustomer.GetData("Kode") & " and NilaiDP>Terpakai"
    query a
    x(0).ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x(0).DeleteRows 0
    If RS.RecordCount > 0 Then x(0).LoadRows RS.GetRows
    For i = 0 To x(0).UpperBound(1)
        a = x(0)(i, col1("Tanggal Lunas").ColIndex)
        x(0)(i, col1("Tanggal Lunas").ColIndex) = cTanggal2(a)
    Next
    TDBGrid1.Rebind
    a = "select NoSJ, TanggalSJ, Total, Pelunasan, Total-Pelunasan, 0, '', '', MataUang from t_SPP" & pTipe & " where abs(Total-Pelunasan)*100>100 and t_SPP~.Kode=" & fCustomer.GetData("Kode") & " and MataUang='" & fMataUang & "' order by TanggalSJ, NoSJ"
    query a
    x(1).ReDim 0, 0, 0, TDBGrid2.Columns.count - 1
    x(1).DeleteRows 0
    If RS.RecordCount > 0 Then x(1).LoadRows RS.GetRows
    TDBGrid2.Rebind
    a = "select NoNR, TanggalNR, Total, Potong, Total-Potong, 0, '','' from t_NR~ where Total-Potong>0 and t_NR~.KodeCustomer=" & fCustomer.GetData("Kode") & " and MataUang='" & fMataUang & "'"
    query a
    x(2).ReDim 0, 0, 0, TDBGrid3.Columns.count - 1
    x(2).DeleteRows 0
    If RS.RecordCount > 0 Then x(2).LoadRows RS.GetRows
    TDBGrid3.Rebind
    col1("Nilai RP").Visible = fMataUang = "USD"
    Exit Sub
err:
End Sub

Private Sub fTanggal_LostFocus()
    fNoSTT_LostFocus
End Sub

Private Sub TDBDropDown1_DropDownClose()
    col1("No Acc").value = TDBDropDown1.Columns("No Acc").value
    col1("Ket Acc").value = TDBDropDown1.Columns("Ket Acc").value
End Sub

Private Sub TDBDropDown1_DropDownOpen()
    If Not LNoAccount Then
        a = "select NoAccount, Deskripsi from m_ChartAccount where Child=0 order by NoAccount"
        query a
        z.ReDim 0, 0, 0, TDBDropDown1.Columns.count - 1
        z.DeleteRows 0
        If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
        TDBDropDown1.Rebind
        LNoAccount = True
    End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    mChange = True
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_Validate(Cancel As Boolean)
    HitungTotal 0
End Sub

Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    mChange = True
End Sub

Private Sub TDBGrid2_DblClick()
    If Not TDBGrid2.AllowUpdate Then Exit Sub
    If col2(TDBGrid2.Col).Caption = "Bayar" Then
        col2("Bayar").value = col2("Sisa").value
        mChange = True
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If fMataUang = "RP" Then n = col1("Nilai").value Else n = col1("Nilai RP").value
    fKurs.Tag = n / col1("Nilai").value
    fKurs = Round(fKurs.Tag)
End Sub

Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If CDbl(x(1)(Bookmark, col2("Sisa").ColIndex)) = CDbl(x(1)(Bookmark, col2("Bayar").ColIndex)) Then
        RowStyle.ForeColor = vbBlue
    End If
End Sub

Private Sub GoEvent(ByVal tVal As String)
    v = tVal = "NEW"
    fNoSTT.Enabled = v
    fTanggal.Enabled = v
    TDBGrid1.AllowUpdate = v
    TDBGrid2.AllowUpdate = v
    TDBGrid3.AllowUpdate = v
    fTotalPembayaran.Enabled = v
    fTotalPelunasan.Enabled = v
    fTotalPotongan.Enabled = v
    fDebet.Enabled = v
    fKredit.Enabled = v
    fPembulatanPenjualan.Enabled = v
    fSave.Enabled = v
    fKalkulasi.Enabled = v
    fDelete.Enabled = Not v
    fPrint.Enabled = Not v
    fCustomer.Enabled = v
    fNew.Enabled = Not v
End Sub

Private Sub fKalkulasi_Click()
On Error Resume Next
    TDBGrid2.Update
    Nilai = fDebet.Tag
    For i = 0 To x(1).UpperBound(1)
        x(1)(i, col2("Bayar").ColIndex) = 0
        If Nilai >= CDbl(x(1)(i, col2("Sisa").ColIndex)) - 2000 Then
            x(1)(i, col2("Bayar").ColIndex) = x(1)(i, col2("Sisa").ColIndex)
        ElseIf Nilai > 0 Then
            x(1)(i, col2("Bayar").ColIndex) = Nilai
        End If
        Nilai = Nilai - x(1)(i, col2("Bayar").ColIndex)
    Next
    TDBGrid2.Rebind
    mChange = True
    HitungTotal 1
End Sub

Private Sub HitungTotal(ByVal tVal As Integer)
On Error Resume Next
    If Not mChange Then Exit Sub
    If tVal < 1 Then
        a = 0
        TDBGrid1.Update
        For i = 0 To x(0).UpperBound(1)
            a = a + x(0)(i, col1("Nilai").ColIndex)
        Next
        fTotalPembayaran.Tag = a
        fTotalPembayaran = cDecimal(a)
    End If
    If tVal < 2 Then
        a = 0
        TDBGrid2.Update
        For i = 0 To x(1).UpperBound(1)
            a = a + x(1)(i, col2("Bayar").ColIndex)
        Next
        fTotalPelunasan.Tag = a
        fTotalPelunasan = cDecimal(a)
    End If
    If tVal < 3 Then
        a = 0
        TDBGrid3.Update
        For i = 0 To x(2).UpperBound(1)
            a = a + x(2)(i, col3("Nilai").ColIndex)
        Next
        fTotalPotongan.Tag = a
        fTotalPotongan = cDecimal(a)
    End If
    fDebet.Tag = fTotalPembayaran.Tag
    fDebet = fTotalPembayaran
    fKredit.Tag = fTotalPelunasan.Tag - fTotalPotongan.Tag
    fKredit = cDecimal(fKredit.Tag)
    fPembulatanPenjualan.Tag = fTotalPembayaran.Tag - fTotalPelunasan.Tag + fTotalPotongan.Tag
    fPembulatanPenjualan = cDecimal(fPembulatanPenjualan.Tag)
    mChange = False
End Sub

Private Sub fList_Click()
    FormList.LoadMe _
        "STT Belum Cair@STT Lunas", "select NoSTT,Tanggal,Debet,Kredit,Pembulatan,Nama from t_STT~ left join m_customer on t_STT" & pTipe & ".KodeCustomer=m_customer.Kode where GiroNum>0@" & _
        "select NoSTT,Tanggal,Debet,Kredit,Pembulatan,Nama from t_STT~ left join m_customer on t_STT~.KodeCustomer=m_customer.Kode where GiroNum=0", _
        "No STT@TANGGAL@CUSTOMER@MATA UANG", "NoSTT@TANGGAL@Nama@MataUang", "1000@1000@2000@700", "String@Date@StringLike@String", _
        "NO STT@TANGGAL@DEBET@KREDIT@PEMBULATAN@CUSTOMER", _
        "1000@1000@1500@1500@1000@2000", _
        "String@Date@Decimal@Decimal@Decimal@String", Me, " order by left(Tanggal,1) Desc, NoSTT Desc"
    FormList.Show , Me
End Sub

Sub GetResult(ByVal tNo As String)
On Error Resume Next
    fKurs = ""
    fKurs.Tag = ""
    a = "select top 1 NoSTT,Tanggal,Nama, MataUang from t_STT" & pTipe & " left join m_customer on m_customer.Kode=t_STT" & pTipe & ".KodeCustomer where NoSTT='" & esc(tNo) & "'"
    query a
    If RS.RecordCount <= 0 Then
        MsgBox "No Data"
        ClearScreen
        Exit Sub
    End If
    GoEvent "SEE"
    fNoSTT = RS.Fields("NoSTT").value
    fQuick = fNoSTT
    fTanggal = cTanggal(RS.Fields("Tanggal").value)
    fCustomer = RS.Fields("Nama").value
    fCustomer.FindIndex
    fMataUang = RS.Fields("MataUang").value
    col1("Nilai RP").Visible = fMataUang <> "RP"
    col2("Bayar RP").Visible = fMataUang <> "RP"
    col3("Nilai RP").Visible = fMataUang <> "RP"
    a = "select CaraBayar,NamaBank,NoGiro, TanggalGiro, Nilai, NilaiRP, DebetAcc, KetAcc, Keterangan, IdDP from t_STTPembayaran" & pTipe & " where NoSTT='" & esc(tNo) & "' order by TanggalGiro"
    query a
    x(0).ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x(0).DeleteRows 0
    If RS.RecordCount > 0 Then
        x(0).LoadRows RS.GetRows
        For i = 0 To RS.RecordCount - 1
            If x(0)(i, col1("Tanggal Lunas").ColIndex) <> 0 Then
                Tgl = x(0)(i, col1("Tanggal Lunas").ColIndex)
                x(0)(i, col1("Tanggal Lunas").ColIndex) = cTanggal2(Tgl)
            Else
                x(0)(i, col1("Tanggal Lunas").ColIndex) = ""
            End If
        Next
    End If
    TDBGrid1.Rebind
    a = "select NoFaktur, TanggalFaktur, NilaiFaktur, Lunas, Sisa, sum(Nilai), sum(BayarRP), Keterangan, MataUang from t_STTPelunasan" & pTipe & " where NoSTT='" & esc(tNo) & "' group by NoFaktur, TanggalFaktur, NilaiFaktur, Lunas, Sisa, Keterangan, MataUang order by TanggalFaktur, NoFaktur"
    query a
    x(1).ReDim 0, 0, 0, TDBGrid2.Columns.count - 1
    x(1).DeleteRows 0
    If RS.RecordCount > 0 Then x(1).LoadRows RS.GetRows
    TDBGrid2.Rebind
    a = "select NoBukti,TanggalBukti,NilaiRetur, Terpotong, Sisa, Nilai,NilaiRP, Keterangan from t_STTPotongan" & pTipe & " where NoSTT='" & esc(tNo) & "'"
    query a
    x(2).ReDim 0, 0, 0, TDBGrid3.Columns.count - 1
    x(2).DeleteRows 0
    If RS.RecordCount > 0 Then x(2).LoadRows RS.GetRows
    TDBGrid3.Rebind
    mChange = True
    HitungTotal 0
End Sub

Private Sub ClearScreen()
On Error Resume Next
    x(0).ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x(0).DeleteRows 0
    TDBGrid1.Rebind
    x(1).ReDim 0, 0, 0, TDBGrid2.Columns.count - 1
    x(1).DeleteRows 0
    TDBGrid2.Rebind
    x(2).ReDim 0, 0, 0, TDBGrid3.Columns.count - 1
    x(2).DeleteRows 0
    TDBGrid3.Rebind
    fTotalPelunasan = 0
    fTotalPembayaran = 0
    fTotalPotongan = 0
    fDebet = 0
    fKredit = 0
    fPembulatanPenjualan = 0
    fCustomer = ""
    fNoSTT = ""
    fTanggal = "__/__/__"
    fMataUang = "RP"
    fKurs.Tag = ""
    fKurs = ""
    col1("Nilai RP").Visible = False
    col2("Bayar RP").Visible = False
    col3("Nilai RP").Visible = False
End Sub

Private Sub fNew_Click()
On Error Resume Next
    ClearScreen
    a = "select max(NoSTT) from t_STT" & pTipe & " where Tanggal>" & pAddNoLong
    query a
    If Not IsNull(RS.Fields(0).value) Then
        b = Left(RS.Fields(0).value, 5) + 1
    Else
        b = 1
    End If
    fNoSTT = zerofill(b, 5) & "/" & Right(pServerDate, 2)
    fQuick = fNoSTT
    fTanggal = pServerDate
    GoEvent "NEW"
    fNoSTT.SetFocus
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    LCustomer = False
    Set TDBDropDown1.Array = z
    Set TDBGrid1.Array = x(0)
    Set col1 = TDBGrid1.Columns
    col1("Ket Acc").Locked = True
    fNew_Click
    Dim vitem As New TrueOleDBGrid80.ValueItem
    With col1("Cara Bayar").ValueItems
        .Presentation = dbgComboBox
        a = "select Cara from m_CaraBayar"
        query a
        For i = 0 To RS.RecordCount - 1
            vitem.value = RS.Fields(0).value
            .Add vitem
            RS.MoveNext
        Next
    End With
    With col1("Nama Bank").ValueItems
        .Presentation = dbgComboBox
        a = "select NamaBank from m_Bank order by NamaBank"
        query a
        For i = 0 To RS.RecordCount - 1
            vitem.value = RS.Fields(0).value
            .Add vitem
            RS.MoveNext
        Next
    End With
    col1("No Acc").DropDown = TDBDropDown1
    col1("No Acc").AutoDropDown = True
    col1("Tanggal Lunas").Alignment = dbgRight
    col1("Nilai").NumberFormat = "Standard"
    col1("Nilai").Alignment = dbgRight
    col1("Nilai RP").NumberFormat = "Standard"
    col1("Nilai RP").Alignment = dbgRight
    col1("Nilai RP").Visible = False
    col1("IdDP").Visible = False
    
    Set TDBGrid2.Array = x(1)
    Set col2 = TDBGrid2.Columns
    Set TDBGrid3.Array = x(2)
    col2("Nilai Faktur").NumberFormat = "Standard"
    col2("Nilai Faktur").Alignment = dbgRight
    col2("Lunas").NumberFormat = "Standard"
    col2("Lunas").Alignment = dbgRight
    col2("Sisa").NumberFormat = "Standard"
    col2("Sisa").Alignment = dbgRight
    col2("Bayar").NumberFormat = "Standard"
    col2("Bayar").Alignment = dbgRight
    col2("Bayar RP").NumberFormat = "Standard"
    col2("Bayar RP").Alignment = dbgRight
    col2("Bayar RP").Locked = True
    col2("Bayar RP").Visible = False
    col2("MataUang").Visible = False
    TDBGrid2.FetchRowStyle = True
    
    Set col3 = TDBGrid3.Columns
    col3("Tanggal").NumberFormat = "FormatText Event"
    TDBGridSetTag TDBGrid3, "Nilai Retur@Sdh Potong@Sisa@Nilai@Nilai RP", "Decimal"
    TDBGridLoad TDBGrid3
    TDBGrid3.FetchRowStyle = True
    
    fCustomer.SetHeader "NAMA CUSTOMER@KOTA@MATA UANG@*KODE"
    fCustomer.SetWidth "3000@1500@700@1000"
    fCustomer.SetType "String@String@String@Integer"
    
End Sub

Private Sub fPrint_Click()
    'FormReport.LoadMe "STTPenjualan" & pTipe & ".rpt", fNoSTT
End Sub

Private Sub fSave_Click()
'On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    TDBGrid2.Update
    TDBGrid3.Update
    HitungTotal 0
    fCustomer_KeyDown 0, 0
    TDBGrid1_RowColChange 0, 0
'Hitung Tanggal Pelunasan Per Surat Jalan
Dim Tgl() As Long
Dim NoGiro() As String
Dim TglIndex() As Long
Dim tNilaiBayar() As Double

    ReDim TglIndex(x(0).UpperBound(1))
    ReDim Tgl(x(0).UpperBound(1))
    ReDim tNilaiBayar(x(0).UpperBound(1))
    ReDim NoGiro(x(0).UpperBound(1))
    ReDim carabayar(x(0).UpperBound(1))
    'Urutin Tanggal Bayar Dulu
    For i = 0 To x(0).UpperBound(1)
        If cD(x(0)(i, col1("Tanggal Lunas").ColIndex)) = "A" Then x(0)(i, col1("Tanggal Lunas").ColIndex) = Replace(fTanggal, "/", "")
        a = cD(x(0)(i, col1("Tanggal Lunas").ColIndex))
        If a = 0 Then a = cD(pServerDate)
        Tgl(i) = a
    Next
    For i = 0 To x(0).UpperBound(1)
        m = i
        For j = i + 1 To x(0).UpperBound(1)
            If Tgl(j) < Tgl(i) Then m = j
        Next
        
        TglIndex(i) = m
        TglIndex(m) = i
        t = Tgl(i)
        Tgl(i) = Tgl(m)
        Tgl(m) = t
    Next
    For i = 0 To UBound(TglIndex)
        tNilaiBayar(i) = x(0)(TglIndex(i), col1("Nilai").ColIndex)
        NoGiro(i) = x(0)(TglIndex(i), col1("No Giro").ColIndex)
        carabayar(i) = x(0)(TglIndex(i), col1("Cara Bayar").ColIndex)
    Next
    r = 0
    For i = 0 To x(2).UpperBound(1)
       r = r + x(2)(i, col3("Nilai").ColIndex)
    Next
    tNilaiBayar(0) = tNilaiBayar(0) + r - CDec(fPembulatanPenjualan.Text)
Dim GiroNum As Byte
Dim Nilai As Double
Dim nilaiRP As Double
    For i = 0 To x(0).UpperBound(1)
        If cD(x(0)(i, col1("Tanggal Lunas").ColIndex)) <= cD(fTanggal) Then
            MsgBox "Salah Input Tanggal No Giro: " & x(0)(i, col1("No Giro").ColIndex)
            GoTo err
        End If
        If x(0)(i, col1("Cara Bayar").ColIndex) <> "" Then
            nilaiRP = nilaiRP + IIf(fMataUang = "RP", x(0)(i, col1("Nilai").ColIndex), x(0)(i, col1("Nilai RP").ColIndex))
            Nilai = Nilai + x(0)(i, col1("Nilai").ColIndex)
            If x(0)(i, col1("Cara Bayar").ColIndex) = "GIRO" Then
                If x(0)(i, col1("No Giro").ColIndex) = "" Then
                    MsgBox "No Giro Harus Diisi"
                    GoTo err
                Else
                    GiroNum = GiroNum + 1
                End If
            End If
            IdDP = IIf(x(0)(i, col1("Cara Bayar").ColIndex) <> "DP" Or IsNull(x(0)(i, col1("IdDP").ColIndex)) Or x(0)(i, col1("IdDP").ColIndex) = "", -1, x(0)(i, col1("IdDP").ColIndex))
            
            a = "insert into t_STTPembayaran" & pTipe & "(NoSTT,Tanggal,CaraBayar,NamaBank,NoGiro,TanggalGiro, Nilai,NilaiRP,DebetAcc, KetAcc, Keterangan,Status,KetCustomer,MataUang,IdDP) values('" & _
                fNoSTT & _
                "'," & cD(fTanggal) & _
                ",'" & x(0)(i, col1("Cara Bayar").ColIndex) & _
                "','" & x(0)(i, col1("Nama Bank").ColIndex) & _
                "','" & x(0)(i, col1("No Giro").ColIndex) & _
                "'," & IIf(cD(x(0)(i, col1("Tanggal Lunas").ColIndex)) = 0 Or (cD(x(0)(i, col1("Tanggal Lunas").ColIndex)) < cD(fTanggal) And x(0)(i, col1("Cara Bayar").ColIndex) = "GIRO"), cD(fTanggal), cD(x(0)(i, col1("Tanggal Lunas").ColIndex))) & _
                "," & cNum(x(0)(i, col1("Nilai").ColIndex)) & _
                "," & cNum(nilaiRP) & _
                ",'" & x(0)(i, col1("No Acc").ColIndex) & _
                "','" & x(0)(i, col1("Ket Acc").ColIndex) & _
                "','" & x(0)(i, col1("Keterangan").ColIndex) & _
                "'," & IIf(x(0)(i, col1("Cara Bayar").ColIndex) = "GIRO", 0, 1) & _
                ",'" & fCustomer & _
                "','" & fMataUang & "'," & IdDP & ")"
            If ExecMe(a) = 0 Then GoTo err
            If cD(x(0)(i, col1("Tanggal Lunas").ColIndex)) = 0 Then
                s = "update t_STTPembayaran~ set TanggalCair=TanggalGiro, TanggalSetor=TanggalGiro where NoSTT='" & fNoSTT & "'"
            End If
            If x(0)(i, col1("Cara Bayar").ColIndex) = "DP" Then
                a = "update t_DP set Terpakai=Terpakai+" & cNum(x(0)(i, col1("Nilai").ColIndex)) & " where IdDP=" & x(0)(i, col1("IdDP").ColIndex)
                If ExecMe(a) = 0 Then GoTo err
            End If
        End If
        a = IIf(cD(x(0)(i, col1("Tanggal Lunas").ColIndex)) = 0 Or (cD(x(0)(i, col1("Tanggal Lunas").ColIndex)) < cD(fTanggal) And x(0)(i, col1("Cara Bayar").ColIndex) = "GIRO"), cD(fTanggal), cD(x(0)(i, col1("Tanggal Lunas").ColIndex)))
    Next
    TDBGrid2.Update
    Dim kurs As Double
    kurs = nilaiRP / Nilai
    j = 0
    
    For i = 0 To x(1).UpperBound(1)
        If x(1)(i, col2("Bayar").ColIndex) <> "" And x(1)(i, col2("Bayar").ColIndex) <> "0" Then
            Dim BayarRP As Double
            Dim Bayar As Double
            BayarRP = x(1)(i, col2("Bayar").ColIndex) * kurs
            Bayar = x(1)(i, col2("Bayar").ColIndex)
            tNilaiBayar(j) = tNilaiBayar(j) - Bayar
            a = "update t_SPP" & pTipe & " set WaktuUpdate=now, NoSTT='" & fNoSTT.Text & "', TanggalLunas=" & Tgl(j) & ", Pelunasan=Pelunasan+" & cNum(x(1)(i, col2("Bayar").ColIndex)) & " where NoSJ='" & esc(x(1)(i, col2("No SJ").ColIndex)) & "'"
            ExecMe a
            If tNilaiBayar(j) < 0 Then
                a = "insert into t_STTPelunasan" & pTipe & "(NoSTT,Tanggal,NoFaktur,TanggalFaktur,KodeCustomer, NilaiFaktur, Lunas, Sisa, Nilai, BayarRP, Keterangan, MataUang, NoGiro, TanggalGiro, TanggalPelunasan) values('" & _
                fNoSTT & _
                "'," & cD(fTanggal) & _
                ",'" & x(1)(i, col2("No SJ").ColIndex) & _
                "'," & x(1)(i, col2("Tanggal").ColIndex) & _
                "," & fCustomer.GetData("Kode") & _
                "," & cNum(x(1)(i, col2("Nilai Faktur").ColIndex)) & _
                "," & cNum(x(1)(i, col2("Lunas").ColIndex)) & _
                "," & cNum(x(1)(i, col2("Sisa").ColIndex)) & _
                "," & cNum(tNilaiBayar(j) + Bayar) & _
                "," & cNum((tNilaiBayar(j) + Bayar) * kurs) & _
                ",'" & x(1)(i, col2("Keterangan").ColIndex) & _
                "','" & fMataUang.Text & _
                "','" & NoGiro(j) & _
                "'," & Tgl(j) & _
                "," & IIf(carabayar(j) = "GIRO", 999999, Tgl(j)) & ")"
                If ExecMe(a) = 0 Then GoTo err
                Bayar = -tNilaiBayar(j)
                j = j + 1
                If j <= UBound(tNilaiBayar) Then tNilaiBayar(j) = tNilaiBayar(j) + tNilaiBayar(j - 1)
            End If
            If j <= UBound(tNilaiBayar) Then
                a = "insert into t_STTPelunasan" & pTipe & "(NoSTT,Tanggal,NoFaktur,TanggalFaktur,KodeCustomer, NilaiFaktur, Lunas, Sisa, Nilai, BayarRP, Keterangan, MataUang, NoGiro, TanggalGiro, TanggalPelunasan) values('" & _
                    fNoSTT & _
                    "'," & cD(fTanggal) & _
                    ",'" & x(1)(i, col2("No SJ").ColIndex) & _
                    "'," & x(1)(i, col2("Tanggal").ColIndex) & _
                    "," & fCustomer.GetData("Kode") & _
                    "," & cNum(x(1)(i, col2("Nilai Faktur").ColIndex)) & _
                    "," & cNum(x(1)(i, col2("Lunas").ColIndex)) & _
                    "," & cNum(x(1)(i, col2("Sisa").ColIndex)) & _
                    "," & cNum(Bayar) & _
                    "," & cNum(Bayar * kurs) & _
                    ",'" & x(1)(i, col2("Keterangan").ColIndex) & _
                    "','" & fMataUang.Text & _
                    "','" & NoGiro(j) & _
                    "'," & Tgl(j) & _
                    "," & IIf(carabayar(j) = "GIRO", 999999, Tgl(j)) & ")"
                If ExecMe(a) = 0 Then
                    GoTo err
                End If
            End If
        End If
    Next
    TDBGrid3.Update
    c = ""
    For i = 0 To x(2).UpperBound(1)
        If x(2)(i, col3("No Bukti").ColIndex) <> "" Then
            nilaiRP = x(2)(i, col3("Nilai").ColIndex) * kurs
            If nilaiRP > 0 Then
                a = "insert into t_STTPotongan" & pTipe & "(NoSTT,Tanggal,NoBukti,TanggalBukti,NilaiRetur, Terpotong, Sisa, Nilai,NilaiRP,Keterangan, MataUang, NamaCustomer) values('" & _
                    fNoSTT & _
                    "'," & cD(fTanggal) & _
                    ",'" & x(2)(i, col3("No Bukti").ColIndex) & _
                    "'," & x(2)(i, col3("Tanggal").ColIndex) & _
                    "," & cNum(x(2)(i, col3("Nilai Retur").ColIndex)) & _
                    "," & cNum(x(2)(i, col3("Sdh Potong").ColIndex)) & _
                    "," & cNum(x(2)(i, col3("Sisa").ColIndex)) & _
                    "," & cNum(x(2)(i, col3("Nilai").ColIndex)) & _
                    "," & cNum(nilaiRP) & _
                    ",'" & x(2)(i, col3("Keterangan").ColIndex) & _
                    "','" & fMataUang & _
                    "','" & fCustomer & "')"
                If ExecMe(a) = 0 Then GoTo err
                a = "update t_NR~ set Potong=Potong+" & cNum(x(2)(i, col3("Nilai").ColIndex)) & " where NoNR='" & x(2)(i, col3("No Bukti").ColIndex) & "'"
                If ExecMe(a) = 0 Then GoTo err
            End If
        End If
    Next
    a = "insert into t_STT" & pTipe & "(NoSTT,Tanggal,Debet,Kredit,Pembulatan,KodeCustomer,GiroNum,MataUang) values('" & _
        fNoSTT & _
        "'," & cD(fTanggal) & _
        "," & cNum(fDebet.Tag) & _
        "," & cNum(fKredit.Tag) & _
        "," & cNum(fPembulatanPenjualan.Tag) & _
        "," & fCustomer.GetData("Kode") & _
        "," & GiroNum & _
        ",'" & fMataUang & "')"
    If ExecMe(a) = 0 Then GoTo err
    CommitTransaction
    MsgBox "SUKSES"
    GetResult fNoSTT
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fDelete_Click()
'On Error GoTo err
    BeginTransaction
    For i = 0 To x(0).UpperBound(1)
        If x(0)(i, col1("Cara Bayar").ColIndex) = "DP" Then
            If x(0)(i, col1("IdDP").ColIndex) > -1 Then
                a = "update t_DP set Terpakai=Terpakai-" & cNum(x(0)(i, col1("Nilai").ColIndex)) & " where IdDP=" & x(0)(i, col1("IdDP").ColIndex)
                If ExecMe(a) = 0 Then GoTo err
            End If
        End If
    Next
    For i = 0 To x(1).UpperBound(1)
        a = "update t_SPP" & pTipe & " set WaktuUpdate=now, TanggalLunas=0,NoSTT=0, StatusLunas=0 , Pelunasan=Pelunasan-" & cNum(x(1)(i, col2("Bayar").ColIndex)) & " where NoSJ='" & esc(x(1)(i, col2("No SJ").ColIndex)) & "'"
        ExecMe a
    Next
    For i = 0 To x(2).UpperBound(1)
        a = "update t_NR~ set Potong=Potong-" & cNum(x(2)(i, col3("Nilai").ColIndex)) & " where NoNR='" & esc(x(2)(i, col3("No Bukti").ColIndex)) & "'"
        ExecMe a
    Next
    a = "delete from t_STTPembayaran" & pTipe & " where NoSTT='" & esc(fNoSTT) & "'"
    ExecMe a
    a = "delete from t_STTPelunasan" & pTipe & " where NoSTT='" & esc(fNoSTT) & "'"
    ExecMe a
    a = "delete from t_STTPotongan" & pTipe & " where NoSTT='" & esc(fNoSTT) & "'"
    ExecMe a
    a = "delete from t_STT" & pTipe & " where NoSTT='" & esc(fNoSTT) & "'"
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    GoEvent "NEW"
    LCustomer = False
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBGrid2_FormatText(ByVal ColIndex As Integer, value As Variant, Bookmark As Variant)
On Error Resume Next
    a = col2(ColIndex).Caption
    If a = "Tanggal" Then
        value = cTanggal(value)
    End If
End Sub

Private Sub TDBGrid2_Validate(Cancel As Boolean)
    HitungTotal 1
End Sub

Private Sub TDBGrid3_AfterColUpdate(ByVal ColIndex As Integer)
    mChange = True
End Sub

Private Sub TDBGrid3_DblClick()
On Error Resume Next
    If Not TDBGrid3.AllowUpdate Then Exit Sub
    If col3(TDBGrid3.Col).Caption = "Nilai" Then
        col3("Nilai").value = col3("Sisa").value
        mChange = True
    End If
End Sub

Private Sub TDBGrid3_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(2)(Bookmark, col3("Nilai").ColIndex) = x(2)(Bookmark, col3("Sisa").ColIndex) Then
        RowStyle.ForeColor = vbBlue
    End If
End Sub

Private Sub TDBGrid3_FormatText(ByVal ColIndex As Integer, value As Variant, Bookmark As Variant)
On Error Resume Next
    a = col3(ColIndex).Caption
    If a = "Tanggal" Then
        value = cTanggal(value)
    End If
End Sub

Private Sub TDBGrid3_Validate(Cancel As Boolean)
    HitungTotal 2
End Sub


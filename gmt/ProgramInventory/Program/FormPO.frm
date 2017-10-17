VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{F6D22ACD-8630-4FE1-97C4-D56AB4AD4DEA}#1.0#0"; "usrtruecombo.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormPO 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PURCHASE ORDER"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Tag             =   "11"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   9840
      TabIndex        =   61
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   10200
      TabIndex        =   60
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   10560
      TabIndex        =   59
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   10920
      TabIndex        =   58
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   5160
      TabIndex        =   49
      Top             =   1200
      Width           =   1815
      Begin UsrText.IText fPPN 
         Height          =   270
         Left            =   120
         TabIndex        =   50
         Top             =   960
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
         Locked          =   -1  'True
      End
      Begin VB.CheckBox fPPNYesNo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "PPN"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin UsrText.IText fTotal 
         Height          =   270
         Left            =   120
         TabIndex        =   52
         Top             =   360
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
         Locked          =   -1  'True
      End
      Begin UsrText.IText fGrandTotal 
         Height          =   270
         Left            =   120
         TabIndex        =   53
         Top             =   1560
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
         Locked          =   -1  'True
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.CommandButton fHistorisPembelian 
      Caption         =   "&H"
      Height          =   375
      Left            =   7200
      TabIndex        =   48
      Top             =   120
      Width           =   375
   End
   Begin UsrTrueCombo.ITrueCombo fSupplier 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VB.CommandButton fGotoPR 
      Caption         =   "&GO PR"
      Height          =   375
      Left            =   6360
      TabIndex        =   44
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   5400
      TabIndex        =   42
      Top             =   120
      Width           =   855
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   40
      Top             =   3840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Printed"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "No PO"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nama Barang"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "QTY"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Harga"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "PPN"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Total"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1138"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1058"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1561"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1482"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2302"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1164"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1085"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1720"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1640"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=900"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=820"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2196"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2117"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(62)  =   "Named:id=33:Normal"
      _StyleDefs(63)  =   ":id=33,.parent=0"
      _StyleDefs(64)  =   "Named:id=34:Heading"
      _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   ":id=34,.wraptext=-1"
      _StyleDefs(67)  =   "Named:id=35:Footing"
      _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   "Named:id=36:Selected"
      _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=37:Caption"
      _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(73)  =   "Named:id=38:HighlightRow"
      _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=39:EvenRow"
      _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(77)  =   "Named:id=40:OddRow"
      _StyleDefs(78)  =   ":id=40,.parent=33"
      _StyleDefs(79)  =   "Named:id=41:RecordSelector"
      _StyleDefs(80)  =   ":id=41,.parent=34"
      _StyleDefs(81)  =   "Named:id=42:FilterBar"
      _StyleDefs(82)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   4320
      TabIndex        =   37
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   3480
      TabIndex        =   36
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton fList 
      Caption         =   "&LIST"
      Height          =   375
      Left            =   2760
      TabIndex        =   35
      Top             =   120
      Width           =   615
   End
   Begin UsrText.IText fNamaBarang 
      Height          =   270
      Left            =   7680
      TabIndex        =   6
      Top             =   2640
      Width           =   3375
      _ExtentX        =   5953
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
   Begin UsrText.IText fNoPO 
      Height          =   270
      Left            =   7680
      TabIndex        =   1
      Top             =   840
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
   Begin UsrText.IText fTanggalPO 
      Height          =   270
      Left            =   9840
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
   Begin UsrText.IText fNoPR 
      Height          =   270
      Left            =   7680
      TabIndex        =   3
      Top             =   1440
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
      Enabled         =   0   'False
   End
   Begin UsrText.IText fTanggalPR 
      Height          =   270
      Left            =   9840
      TabIndex        =   4
      Top             =   1440
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
      Enabled         =   0   'False
      MaxLength       =   10
   End
   Begin UsrText.IText fAlamatSupplier 
      Height          =   270
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
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
   Begin UsrText.IText fTelepon 
      Height          =   270
      Left            =   120
      TabIndex        =   24
      Top             =   2040
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
      Enabled         =   0   'False
   End
   Begin UsrText.IText fFax 
      Height          =   270
      Left            =   120
      TabIndex        =   26
      Top             =   2640
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
      Enabled         =   0   'False
   End
   Begin UsrText.IText fContactPerson 
      Height          =   270
      Left            =   120
      TabIndex        =   28
      Top             =   3240
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
      Enabled         =   0   'False
   End
   Begin UsrText.IText fQTY 
      Height          =   270
      Left            =   7680
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   476
      DataType        =   2
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
   Begin UsrText.IText fSatuan 
      Height          =   270
      Left            =   9240
      TabIndex        =   8
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
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
   Begin UsrText.IText fHarga 
      Height          =   270
      Left            =   7680
      TabIndex        =   10
      Top             =   3840
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
   Begin UsrText.IText fMataUang 
      Height          =   270
      Left            =   9240
      TabIndex        =   11
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
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
   Begin UsrText.IText fTempoPembayaran 
      Height          =   270
      Left            =   10320
      TabIndex        =   12
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
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
      Left            =   10320
      TabIndex        =   9
      Top             =   3240
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
   Begin UsrText.IText fKetPO 
      Height          =   270
      Left            =   7680
      TabIndex        =   13
      Top             =   4440
      Width           =   3975
      _ExtentX        =   7011
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
   Begin UsrText.IText fKetPR 
      Height          =   270
      Left            =   7680
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
      _ExtentX        =   7011
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
   Begin UsrText.IText fQuick 
      Height          =   270
      Left            =   8520
      TabIndex        =   46
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
   Begin UsrText.IText fNoContract 
      Height          =   270
      Left            =   5160
      TabIndex        =   56
      Top             =   840
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
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "No Contract"
      Height          =   255
      Left            =   5160
      TabIndex        =   57
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label fStatusPO 
      BackColor       =   &H00404040&
      Caption         =   "BELUM DATANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1635
      Left            =   2280
      TabIndex        =   47
      Top             =   1860
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Quick &Find"
      Height          =   255
      Left            =   7680
      TabIndex        =   45
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "List PO yang Supplier, Tanggal, Mata Uang, Waktu Pembayaran dan PPN Yang Sama"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3600
      Width           =   6495
   End
   Begin VB.Label fStatus 
      BackColor       =   &H00404040&
      Caption         =   "BELUM DI PRINT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7680
      TabIndex        =   41
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan PR"
      Height          =   255
      Left            =   7680
      TabIndex        =   39
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan PO"
      Height          =   255
      Left            =   7680
      TabIndex        =   38
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Kirim"
      Height          =   255
      Left            =   10320
      TabIndex        =   34
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo Pembayaran"
      Height          =   255
      Left            =   10320
      TabIndex        =   33
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Mata Uang"
      Height          =   255
      Left            =   9240
      TabIndex        =   32
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      Height          =   255
      Left            =   7680
      TabIndex        =   31
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan"
      Height          =   255
      Left            =   9240
      TabIndex        =   30
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "QTY"
      Height          =   255
      Left            =   7680
      TabIndex        =   29
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal PR"
      Height          =   255
      Left            =   9840
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NO PR"
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal PO"
      Height          =   255
      Left            =   9840
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NO PO"
      Height          =   255
      Left            =   7680
      TabIndex        =   17
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE ORDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FormPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LSupplier As Boolean
Dim x As New XArrayDB
Dim fIdStock As Long
Dim mKoma As Boolean
Dim mStatusPO As Byte

Sub GetResult(ByVal tNo As String, Optional ByVal tRefreshTable As Boolean = True)
On Error Resume Next
    fSupplier_KeyDown 0, 0
    If Len(tNo) < 9 Then
        b = InStr(tNo, "/")
        If b = 0 Then
            tNo = zerofill(tNo, 5) & pAddNo
        Else
            tNo = zerofill(Left(tNo, b - 1), 5) & "/" & zerofill(Mid(tNo, b + 1), 2)
        End If
        fQuick = tNo
    End If
    If Len(tNo) < 9 Then
        a = "select m_StockBeli.IdStock, NoPO, TanggalPO, NoPR, TanggalPR, t_PR.NamaSupplier, Alamat, Telepon, Fax, ContactPerson, " & _
            "NamaBarang, QTYOrder, Satuan, TanggalKirim, Harga, t_PR.MataUang, " & _
            "t_PR.WaktuPembayaran, t_PR.PPNYesNo, KetPR, t_PR.Printed, StatusPO, NoBTB from (t_PR left join m_stockBeli on m_stockBeli.IdStock=t_PR.IdStock) left join m_supplier on m_supplier.Nama=m_stockBeli.SuppSuggestion where NoPO='" & esc(tNo) & "'"
    Else
        a = "select m_StockBeli.IdStock, NoPO, TanggalPO, NoPR, TanggalPR, t_PR.NamaSupplier, Alamat, Telepon, Fax, ContactPerson, " & _
            "NamaBarang, QTYOrder, Satuan, TanggalKirim, Harga, t_PR.MataUang, " & _
            "t_PR.WaktuPembayaran, t_PR.PPNYesNo, KetPR, t_PR.Printed, StatusPO, NoBTB from (t_PR left join m_stockBeli on m_stockBeli.IdStock=t_PR.IdStock) left join m_supplier on m_supplier.Nama=m_stockBeli.SuppSuggestion where NoPR='" & esc(tNo) & "'"
    End If
    query a
    If RS.RecordCount < 1 Then
        MsgBox "No Data"
        Exit Sub
    End If
    If RS.Fields("StatusPO").Value = 0 Then
        LoadMe tNo
        Exit Sub
    End If
    fIdStock = RS.Fields("IdStock").Value
    fNoPO = RS.Fields("NoPO").Value
    fQuick = fNoPO
    fTanggalPO = cTanggal(RS.Fields("TanggalPO").Value)
    fNoPR = RS.Fields("NoPR").Value
    fTanggalPR = cTanggal(RS.Fields("TanggalPR").Value)
    fSupplier = RS.Fields("NamaSupplier").Value
    fAlamatSupplier = RS.Fields("Alamat").Value
    fTelepon = RS.Fields("Telepon").Value
    fFax = RS.Fields("Fax").Value
    fContactPerson = RS.Fields("ContactPerson").Value
    fNamaBarang = RS.Fields("NamaBarang").Value
    fQTY = cDecimal(RS.Fields("QTYOrder").Value)
    fSatuan = RS.Fields("Satuan").Value
    fTanggalKirim = cTanggal(RS.Fields("TanggalKirim").Value)
    fHarga = cDecimal(RS.Fields("Harga").Value)
    fMataUang = RS.Fields("MataUang").Value
    fTempoPembayaran = RS.Fields("WaktuPembayaran").Value
    fTotal = cDecimal(CDec(fQTY) * CDec(fHarga))
    fPPNYesNo.Value = IIf(IsNull(RS.Fields("PPNYesNo").Value), 0, RS.Fields("PPNYesNo").Value)
    fPPN = cDecimal(IIf(fPPNYesNo.Value = 0, 0, 0.1 * CDec(fTotal)))
    fGrandTotal = cDecimal(CDec(fTotal) + CDec(fPPN))
    fKetPR = RS.Fields("KetPR").Value
    fStatus = IIf(RS.Fields("Printed").Value = 0, "BELUM DIPRINT", "SUDAH DIPRINT")
    mStatusPO = RS.Fields("StatusPO").Value
    fStatusPO = IIf(mStatusPO = 1, "BELUM DATANG", "SUDAH DATANG " & Mid(RS.Fields("NoBTB").Value, 2))
    If Not tRefreshTable Then Exit Sub
    a = "select Printed-1, Printed*-1, NoPO, NamaBarang, QTYOrder, Harga, t_PR.PPNYesNo*-1, Total from t_PR left join m_StockBeli on t_PR.IdStock=m_StockBeli.IdStock " & _
        "where NamaSupplier='" & esc(fSupplier) & "' and TanggalPO=" & cD(fTanggalPO) & _
        " and t_PR.MataUang='" & esc(fMataUang) & "' and t_PR.WaktuPembayaran='" & esc(fTempoPembayaran) & "'" & _
        " and t_PR.PPNYesNo=" & fPPNYesNo & " order by NoPO"
    query a
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    Set TDBGrid1.Array = x
    TDBGrid1.Rebind
    v = fStatusPO = "BELUM DATANG"
    fSave.Enabled = v
    fDelete.Enabled = v
    If pUpdateHargaSC = 1 Then fSave.Enabled = True
End Sub

Sub LoadMe(ByVal tNo As String)
On Error Resume Next
    Show
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
    fSupplier_KeyDown 0, 0
    a = "select m_StockBeli.IdStock,NoPR, TanggalPR, SuppSuggestion, Alamat, Telepon, Fax, ContactPerson, " & _
        "NamaBarang, QTYOrder, Satuan, TanggalKirim, HargaTerakhir, m_stockBeli.MataUang, " & _
        "m_supplier.WaktuPembayaran, m_stockBeli.PPNYesNo, KetPR from (t_PR left join m_stockBeli on m_stockBeli.IdStock=t_PR.IdStock) left join m_supplier on m_supplier.Nama=m_stockBeli.SuppSuggestion where NoPR='" & esc(tNo) & "' and statusPO=0"
    query a
    fIdStock = RS.Fields("IdStock").Value
    fNoPR = RS.Fields("NoPR").Value
    fTanggalPR = cTanggal(RS.Fields("TanggalPR").Value)
    fSupplier = RS.Fields("SuppSuggestion").Value & ""
    fAlamatSupplier = RS.Fields("Alamat").Value & ""
    fTelepon = RS.Fields("Telepon").Value & ""
    fFax = RS.Fields("Fax").Value & ""
    fContactPerson = RS.Fields("ContactPerson").Value & ""
    fNamaBarang = RS.Fields("NamaBarang").Value
    fQTY = cDecimal(RS.Fields("QTYOrder").Value)
    fSatuan = RS.Fields("Satuan").Value
    fTanggalKirim = cTanggal(RS.Fields("TanggalKirim").Value)
    fHarga = cDecimal(RS.Fields("HargaTerakhir").Value)
    fMataUang = RS.Fields("MataUang").Value
    fTempoPembayaran = RS.Fields("WaktuPembayaran").Value
    fTotal = cDecimal(CDec(fQTY) * CDec(fHarga))
    fPPNYesNo.Value = IIf(IsNull(RS.Fields("PPNYesNo").Value), 0, RS.Fields("PPNYesNo").Value)
    fPPN = IIf(fPPNYesNo.Value = 0, 0, 0.1 * CDec(fTotal))
    fGrandTotal = cDecimal(CDec(fTotal) + CDec(fPPN))
    fKetPR = RS.Fields("KetPR").Value
    a = "select max(NoPO) from t_PR where TanggalPO>" & pAddNoLong
    query a
    If IsNull(RS.Fields(0).Value) Then
        fNoPO = "00001/" & Right(pServerDate, 2)
    Else
        fNoPO = zerofill(Left(RS.Fields(0).Value, 5) + 1, 5) & "/" & Right(pServerDate, 2)
    End If
    fTanggalPO = pServerDate
End Sub

Private Sub fDelete_Click()
    a = "update t_PR set NoContract='', statusPO=0, Printed=0 where NoPR='" & esc(fNoPR) & "'"
    ExecMe a
    MsgBox "SUKSES"
End Sub

Private Sub fFirst_Click()
On Error Resume Next
    b = InStr(fQuick, "/")
    If b = 0 Then fQuick = fQuick & "/" & Right(pServerDate, 2)
    a = "select min(NoPO) from t_PR where TanggalPO>" & Right(fQuick, 2) * 10000
    query a
    fQuick = RS.Fields(0).Value
    fQuick_KeyDown 13, 0
End Sub

Private Sub fLast_Click()
On Error Resume Next
    b = InStr(fQuick, "/")
    If b = 0 Then fQuick = fQuick & "/" & Right(pServerDate, 2)
    a = "select max(NoPO) from t_PR where TanggalPO<" & Right(fQuick, 2) * 10000 + 10000
    query a
    fQuick = RS.Fields(0).Value
    fQuick_KeyDown 13, 0
End Sub
Private Sub fGotoPR_Click()
    FormPR.Show
End Sub

Private Sub fHarga_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 110 Then mKoma = True
End Sub

Private Sub fHarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And mKoma Then
        KeyAscii = 44
        mKoma = False
    End If
End Sub

Private Sub fHarga_Validate(Cancel As Boolean)
    HitungTotal
End Sub

Private Sub fHistorisPembelian_Click()
    FormHistoryPembelian.LoadMe fSupplier, fNamaBarang, Me
End Sub

Private Sub fList_Click()
On Error Resume Next
    FormList.LoadMe "PR@PO@Pembelian", _
        "select NoPR, NoPO, TanggalPO, SuppSuggestion, NamaBarang, t_PR.QTYOrder, Satuan, m_stockBeli.HargaTerakhir, m_StockBeli.PPNYesNo, (1+m_StockBeli.PPNYesNo/10)*t_PR.QTYOrder*HargaTerakhir from t_PR left join m_StockBeli on t_PR.IdStock=m_stockBeli.IdStock where StatusPO=0 and Closed=0@" & _
        "select NoPR, NoPO, TanggalPO, NamaSupplier, NamaBarang, t_PR.QTYOrder, Satuan, Harga, t_PR.PPNYesNo, Total from t_PR left join m_StockBeli on t_PR.IdStock=m_stockBeli.IdStock where StatusPO=1 and Closed=0@" & _
        "select NoPR, NoPO, TanggalPO, NamaSupplier, NamaBarang, t_PR.QTYOrder, Satuan, Harga, t_PR.PPNYesNo, Total from t_PR left join m_StockBeli on t_PR.IdStock=m_stockBeli.IdStock where StatusPO>1 and Closed=0", _
        "Tanggal PO@Nama Supplier@NamaBarang", "TanggalPO@NamaSupplier@NamaBarang", "900@2500@2500", "Date@StringLike@StringLike", _
        "No PR@No PO@Tanggal PO@Nama Supplier@Nama Barang@QTY@Satuan@Harga@PPN@Total", _
        "1500@1000@900@1500@1500@1000@1000@1000@400@1000", _
        "String@String@Date@String@String@Decimal@String@Decimal@YesNo@Decimal", Me, " order by NoPO Desc"
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LSupplier = False
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    LSupplier = False
    fSupplier.ZOrder 0
    fSupplier.SetHeader "Nama@Alamat@Telepon@Fax@Contact Person@*Tempo@*MataUang"
    fSupplier.SetWidth "2000@2000@1500@1000@1500@1000@1000"
    fSupplier.SetType "String@String@String@String@String@String@String"
With TDBGrid1
    .HeadingStyle.Alignment = dbgCenter
    .Columns("").ValueItems.Presentation = dbgCheckBox
    .Columns("Printed").ValueItems.Presentation = dbgCheckBox
    .Columns("PPN").ValueItems.Presentation = dbgCheckBox
    
    .Columns("").Alignment = dbgCenter
    .Columns("Printed").Alignment = dbgCenter
    .Columns("PPN").Alignment = dbgCenter
    .Columns("QTY").Alignment = dbgRight
    .Columns("Harga").Alignment = dbgRight
    .Columns("Total").Alignment = dbgRight
    
    .Columns("Harga").NumberFormat = "Standard"
    .Columns("Total").NumberFormat = "Standard"
    
    For i = 0 To .Columns.Count - 1
        .Columns(i).Locked = True
    Next
    .Columns(0).Locked = False
End With
End Sub

Private Sub fPPNYesNo_Click()
    HitungTotal
End Sub

Private Sub fPrev_Click()
On Error Resume Next
    a = fQuick
    b = InStr(a, "/")
    If b = 0 Then
        a = zerofill(a, 5) & "/" & Right(pServerDate, 2)
    Else
        a = zerofill(Left(a, b - 1) - 1, 5) & "/" & zerofill(Mid(a, b + 1), 2)
    End If
    a = "select NoPO from t_PR where NoPO='" & esc(a) & "'"
    query a
    fQuick = RS.Fields(0).Value
    fQuick_KeyDown 13, 0
End Sub


Private Sub fNext_Click()
On Error Resume Next
    a = fQuick
    b = InStr(a, "/")
    If b = 0 Then
        a = zerofill(a, 5) & "/" & Right(pServerDate, 2)
    Else
        a = zerofill(Left(a, b - 1) + 1, 5) & "/" & zerofill(Mid(a, b + 1), 2)
    End If
    a = "select NoPO from t_PR where NoPO='" & esc(a) & "'"
    query a
    fQuick = RS.Fields(0).Value
    fQuick_KeyDown 13, 0
End Sub
Private Sub fPrint_Click()
With TDBGrid1
    .Update
    c = ""
    For i = 0 To x.UpperBound(1)
        If x(i, 0) <> 0 Then
            c = c & "','" & x(i, .Columns("No PO").ColIndex)
        End If
    Next
End With
    If c = "" Then Exit Sub
    c = " where NoPO in ('" & Mid(c, 4) & "')"
    FormPreview.LoadMe Me, "PO", "", _
    "SELECT t_PR.NamaSupplier, Alamat, Kota, Telepon, Fax, TanggalPO, t_PR.MataUang, t_PR.WaktuPembayaran, t_PR.NoPO, NamaBarang, QTYOrder, Satuan, Harga, QTYOrder*Harga, t_PR.PPNYesNo FROM  (t_PR left join m_StockBeli ON m_StockBeli.IdStock=t_PR.IdStock) left JOIN m_Supplier ON t_PR.NamaSupplier=m_Supplier.Nama " & c & " ORDER BY t_PR.NoPO", _
    "update t_PR set Printed=1" & c
End Sub

Private Sub fQTY_Validate(Cancel As Boolean)
    HitungTotal
End Sub

Private Sub fQuick_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        fQuick.Cancel = True
        GetResult fQuick
        fQuick.FocusSelect
    End If
End Sub

Private Sub fSave_Click()
On Error GoTo err
    BeginTransaction
    a = "update t_PR set " & _
        "NoPO='" & fNoPO & _
        "', TanggalPO=" & cD(fTanggalPO) & _
        ", NamaSupplier='" & fSupplier & _
        "', MataUang='" & fMataUang & _
        "', WaktuPembayaran='" & fTempoPembayaran & _
        "', QTYOrder=" & cNum(fQTY) & _
        ", Harga=" & cNum(fHarga) & _
        ", Total=" & cNum(fGrandTotal) & _
        ", KetPO='" & fKetPO & _
        "', NoContract='" & fNoContract & _
        "', StatusPO=1" & _
        ", PPNYesNo=" & fPPNYesNo.Value & _
        ", TanggalKirim=" & cD(fTanggalKirim) & " where NoPR='" & esc(fNoPR) & "'"
    If ExecMe(a) < 1 Then GoTo err
    If mStatusPO > 1 Then
        a = "update t_BTBDetail set Harga=" & cNum(fHarga) & " where NoPODetail='" & esc(fNoPO) & "'"
        If ExecMe(a) = 0 Then GoTo err
    End If
    a = "update m_Supplier set MataUang='" & fMataUang & "', WaktuPembayaran='" & fTempoPembayaran & "' where Nama='" & esc(fSupplier) & "'"
    If ExecMe(a) < 1 Then GoTo err
    a = "update m_StockBeli set HargaTerakhir=" & cNum(fHarga) & ", TanggalTerakhir=" & cD(fTanggalPO) & " where IdStock=" & fIdStock
    If ExecMe(a) < 1 Then GoTo err
    CommitTransaction
    MsgBox "SUKSES"
    GetResult fNoPO
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not LSupplier Then
        a = "select Nama, Alamat, Telepon, Fax, ContactPerson, WaktuPembayaran, MataUang from m_supplier order by Nama"
        query a
        Dim rs1() As Variant
        rs1 = RS.GetRows
        fSupplier.SetDB rs1
        fSupplier.SetType "String"
        LSupplier = True
    End If
    If KeyCode = 116 Then
        LSupplier = False
    End If
End Sub

Private Sub HitungTotal()
On Error Resume Next
    fTotal = cDecimal(CDec(fQTY) * CDec(fHarga))
    If fPPNYesNo.Value = 0 Then
        fPPN = 0
    Else
        fPPN = cDecimal(0.1 * CDec(fTotal))
    End If
    fGrandTotal = cDecimal(CDec(fTotal) + CDec(fPPN))
    fHarga = cDecimal(fHarga)
End Sub

Private Sub fSupplier_Validate(Cancel As Boolean)
On Error Resume Next
    If fSupplier.ListIndex = -1 Then Cancel = True
    fSupplier = fSupplier.GetData("Nama")
    fAlamatSupplier = fSupplier.GetData("Alamat")
    fTelepon = fSupplier.GetData("Telepon")
    fFax = fSupplier.GetData("Fax")
    fContactPerson = fSupplier.GetData("ContactPerson")
    fTempoPembayaran = fSupplier.GetData("Tempo")
    fMataUang = fSupplier.GetData("MataUang")
End Sub

Private Sub fTempoPembayaran_KeyPress(KeyAscii As Integer)
    LSupplier = False
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If LastRow <> TDBGrid1.Bookmark Then
        GetResult TDBGrid1.Columns("No PO").Value, False
    End If
End Sub


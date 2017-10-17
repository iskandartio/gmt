VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{F6D22ACD-8630-4FE1-97C4-D56AB4AD4DEA}#1.0#0"; "usrtruecombo.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormPR 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PURCHASE REQUISITION"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Tag             =   "10"
   WindowState     =   2  'Maximized
   Begin UsrTrueCombo.ITrueCombo fKodeDepartemen 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
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
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   7920
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton fDaftar 
      Caption         =   "&DAFTAR BARANG"
      Height          =   375
      Left            =   8520
      TabIndex        =   22
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton fStockBaru 
      Caption         =   "STOCK BARU"
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton fBuatPO 
      Caption         =   "P&O"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   480
      Width           =   855
   End
   Begin UsrText.IText IText 
      Height          =   270
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Tag             =   "Closed"
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   476
      Text            =   "0"
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
   Begin UsrText.IText IText 
      Height          =   270
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Tag             =   "Akhir"
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
   Begin UsrText.IText IText 
      Height          =   270
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Tag             =   "Awal"
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
   Begin UsrText.IText IText 
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "NamaBarang"
      Top             =   1200
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
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   2655
      Left            =   3960
      TabIndex        =   12
      Top             =   2400
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Satuan"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "IdStock"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4419"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4339"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1852"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1773"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1164"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1085"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=41,.parent=2,.namedParent=43"
      _StyleDefs(17)  =   "FilterBarStyle:id=44,.parent=1,.namedParent=46"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=42,.parent=41"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=45,.parent=44"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=40,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
      _StyleDefs(42)  =   "Named:id=29:Normal"
      _StyleDefs(43)  =   ":id=29,.parent=0"
      _StyleDefs(44)  =   "Named:id=30:Heading"
      _StyleDefs(45)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=30,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=31:Footing"
      _StyleDefs(48)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=32:Selected"
      _StyleDefs(50)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=33:Caption"
      _StyleDefs(52)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(53)  =   "Named:id=34:HighlightRow"
      _StyleDefs(54)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(55)  =   "Named:id=35:EvenRow"
      _StyleDefs(56)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=36:OddRow"
      _StyleDefs(58)  =   ":id=36,.parent=29"
      _StyleDefs(59)  =   "Named:id=43:RecordSelector"
      _StyleDefs(60)  =   ":id=43,.parent=30"
      _StyleDefs(61)  =   "Named:id=46:FilterBar"
      _StyleDefs(62)  =   ":id=46,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2655
      Left            =   720
      TabIndex        =   11
      Top             =   2400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode Jenis"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Jenis"
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
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=75543200"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2037"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1958"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=74384360"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Named:id=29:Normal"
      _StyleDefs(39)  =   ":id=29,.parent=0"
      _StyleDefs(40)  =   "Named:id=30:Heading"
      _StyleDefs(41)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=30,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=31:Footing"
      _StyleDefs(44)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=32:Selected"
      _StyleDefs(46)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=33:Caption"
      _StyleDefs(48)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(49)  =   "Named:id=34:HighlightRow"
      _StyleDefs(50)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(51)  =   "Named:id=35:EvenRow"
      _StyleDefs(52)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=36:OddRow"
      _StyleDefs(54)  =   ":id=36,.parent=29"
      _StyleDefs(55)  =   "Named:id=39:RecordSelector"
      _StyleDefs(56)  =   ":id=39,.parent=30"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=29"
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox fDepartemen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "nopr2"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "No PR"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tanggal"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tgl Kirim"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "IdStock"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Kode Jenis"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "NamaJenis"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Nama Barang"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Satuan"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Order"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Datang"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Sisa"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Ket. PR"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "No Contract"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Closed"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "PO"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "updated"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1138"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1058"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1455"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1376"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1667"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1588"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1508"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1535"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1455"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2037"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1958"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=1402"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1323"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=1455"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1376"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=1376"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=1296"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=1191"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=1111"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=1773"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=1693"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=1482"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=1402"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=1958"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=1879"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=85,.parent=2,.namedParent=87"
      _StyleDefs(17)  =   "FilterBarStyle:id=88,.parent=1,.namedParent=90"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=86,.parent=85"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=89,.parent=88"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=102,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=24,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=68,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=65,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=66,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=67,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=84,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=81,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=82,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=83,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=72,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=69,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=70,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=71,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=28,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=94,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=91,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=92,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=93,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=40,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=37,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=38,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=39,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=64,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=61,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=62,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=63,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=44,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=41,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=42,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=43,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=48,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=45,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=46,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=47,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=52,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=49,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=50,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=51,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=80,.parent=11"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=77,.parent=12"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=78,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=79,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=60,.parent=11"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=57,.parent=12"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=58,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=59,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=56,.parent=11"
      _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=53,.parent=12"
      _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=54,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=55,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=76,.parent=11"
      _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=73,.parent=12"
      _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=74,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=75,.parent=15"
      _StyleDefs(94)  =   "Splits(0).Columns(16).Style:id=98,.parent=11"
      _StyleDefs(95)  =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=12"
      _StyleDefs(96)  =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=15"
      _StyleDefs(98)  =   "Named:id=29:Normal"
      _StyleDefs(99)  =   ":id=29,.parent=0"
      _StyleDefs(100) =   "Named:id=30:Heading"
      _StyleDefs(101) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(102) =   ":id=30,.wraptext=-1"
      _StyleDefs(103) =   "Named:id=31:Footing"
      _StyleDefs(104) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(105) =   "Named:id=32:Selected"
      _StyleDefs(106) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(107) =   "Named:id=33:Caption"
      _StyleDefs(108) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(109) =   "Named:id=34:HighlightRow"
      _StyleDefs(110) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(111) =   "Named:id=35:EvenRow"
      _StyleDefs(112) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(113) =   "Named:id=36:OddRow"
      _StyleDefs(114) =   ":id=36,.parent=29"
      _StyleDefs(115) =   "Named:id=87:RecordSelector"
      _StyleDefs(116) =   ":id=87,.parent=30"
      _StyleDefs(117) =   "Named:id=90:FilterBar"
      _StyleDefs(118) =   ":id=90,.parent=29"
   End
   Begin UsrText.IText IText 
      Height          =   270
      Index           =   4
      Left            =   5040
      TabIndex        =   6
      Tag             =   "PO"
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
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
   Begin UsrText.IText IText 
      Height          =   270
      Index           =   5
      Left            =   6240
      TabIndex        =   20
      Tag             =   "NoContract"
      Top             =   1200
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "No Contract"
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.Label fKetNama 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   8775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PO"
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Closed"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Departemen"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE REQUISITION"
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
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "FormPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim LDepartemen As Boolean
Dim z1 As New XArrayDB
Dim z2 As New XArrayDB
Dim LJenis As Boolean
Dim col1 As TrueOleDBGrid80.Columns

Private Sub fBuatPO_Click()
'    If col1("PO").Value <> "" And Not IsNull(col1("PO").Value) Then Exit Sub
    nopr = zerofill(col1("No PR").Value, 5) & "/" & zerofill(fKodeDepartemen, 2) & Mid(col1("Tanggal").Text, 3)
    FormPO.GetResult nopr
    FormPO.Show
End Sub

Private Sub fDaftar_Click()
    FormStockBeli.LoadMe Me
End Sub

Sub SetOtherRowData(ByVal tIdStock As Long)
    If col1("PO").Value <> "" Then Exit Sub
    a = "select top 1 KodeJenis, Jenis, NamaBarang, Satuan, IdStock from m_StockBeli where IdStock=" & tIdStock
    query a
    If RS.RecordCount = 0 Then Exit Sub
    col1("Kode Jenis").Value = RS.Fields("KodeJenis").Value
    col1("NamaJenis").Value = RS.Fields("Jenis").Value
    col1("Nama Barang").Value = RS.Fields("NamaBarang").Value
    col1("Satuan").Value = RS.Fields("Satuan").Value
    col1("IdStock").Value = RS.Fields("IdStock").Value
    col1("updated").Value = "1"
End Sub

Private Sub fKodeDepartemen_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If Not LDepartemen Then
        a = "select KdDept,Departemen, Bos from m_departemen order by KdDept"
        query a
Dim rs1() As Variant
        rs1 = RS.GetRows
        fKodeDepartemen.SetDB rs1
        fKodeDepartemen.SetType "String"
        LDepartemen = True
    End If
End Sub

Private Sub fKodeDepartemen_KeyUp(KeyCode As Integer, Shift As Integer)
    fKodeDepartemen.Cancel
    For i = 0 To fKodeDepartemen.ListCount - 1
        If fKodeDepartemen.zz(i, "Kode") = fKodeDepartemen Then
            fKodeDepartemen.SetListIndex i
        End If
    Next
End Sub

Private Sub fKodeDepartemen_Validate(Cancel As Boolean)
On Error Resume Next
    If fKodeDepartemen.ListIndex = -1 Then
        fDepartemen = ""
        x.ReDim 0, 0, 0, col1.Count - 1
        x.DeleteRows 0
        TDBGrid1.Rebind
    Else
        fDepartemen = fKodeDepartemen.GetData("Departemen")
    End If
    GetDetail
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LDepartemen = False
End Sub

Private Sub Form_Load()
On Error Resume Next
    Caption = Caption & "---" & pTipe
    fKodeDepartemen.SetHeader "Kode@Departemen@*Bos"
    fKodeDepartemen.SetWidth "1000@2000@1500"
    fKodeDepartemen.SetType "String@String@String"
    Set col1 = TDBGrid1.Columns
    TDBGrid1.Array = x
    TDBDropDown1.Array = z1
    TDBDropDown2.Array = z2
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    col1("Tanggal").Tag = "Date"
    col1("Tgl Kirim").Tag = "Date"
    col1("Datang").Tag = "Decimal"
    col1("Order").Tag = "Decimal"
    col1("Sisa").Tag = "Decimal"
    col1("Closed").Tag = "OK?"
    TDBGridLoad TDBGrid1
    col1("Kode Jenis").DropDown = TDBDropDown1
    col1("Nama Barang").DropDown = TDBDropDown2
    col1("Kode Jenis").AutoDropDown = True
    col1("Nama Barang").AutoDropDown = True
    TDBGridSetLock TDBGrid1, "Datang@Sisa@PO@No Contract", True
    TDBGridSetVisible TDBGrid1, "NamaJenis@updated@nopr2", False
    
    TDBGrid1.FetchRowStyle = True
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
End Sub

Private Function MyFilter() As String
    For i = 0 To IText.Count - 1
        IText(i) = Trim(IText(i))
    Next
    If IText(0) <> "" Then MyFilter = MyFilter & " and NamaBarang like '%" & IText(0) & "%'"
    If cD(IText(1)) <> "A" Then MyFilter = MyFilter & " and TanggalPR>=" & cD(IText(1))
    If cD(IText(2)) <> "A" Then MyFilter = MyFilter & " and TanggalPR<=" & cD(IText(2))
    If IText(3) <> "" Then MyFilter = MyFilter & " and Closed=" & IText(3)
    If IText(4) <> "" Then MyFilter = MyFilter & " and StatusPO=" & IText(4)
    If IText(5) <> "" Then MyFilter = MyFilter & " and NoContract like '%" & IText(5) & "%'"
End Function

Private Sub GetDetail()
On Error Resume Next
    If fKodeDepartemen = "" Then Exit Sub
    a = "select nopr,left(NoPR,5), TanggalPR, TanggalKirim, t_PR.IdStock, KodeJenis, Jenis, NamaBarang, Satuan, QTYOrder, Datang, QTYOrder-Datang, KetPR, NoContract, Closed*-1, NoPO, 0  from t_PR left join m_StockBeli on t_PR.IdStock=m_stockBeli.IdStock where t_PR.Dept='" & esc(fDepartemen) & "'" & MyFilter & " order by TanggalPR,NoPR"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 And fKodeDepartemen.ListIndex <> -1 Then x.LoadRows RS.GetRows
    For i = 0 To x.UpperBound(1)
        x(i, col1("No PR").ColIndex) = CLng(x(i, col1("No PR").ColIndex))
        x(i, col1("Tanggal").ColIndex) = cTanggal2(x(i, col1("Tanggal").ColIndex))
        x(i, col1("Tgl Kirim").ColIndex) = cTanggal2(x(i, col1("Tgl Kirim").ColIndex))
    Next
    TDBGrid1.Rebind
    TDBGrid1.MoveLast
    TDBGrid1.SetFocus
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left * 2
    fKetNama.Width = TDBGrid1.Width
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fPrint_Click()
    'FormReport.LoadMe "PR.rpt", cD(col1("Tanggal").Value) & "@" & fKodeDepartemen.GetData("Departemen")
End Sub

Private Sub fSave_Click()
    UpdateDB
End Sub

Private Sub fStockBaru_Click()
On Error Resume Next
    a = "select max(IdStock) from m_StockBeli"
    query a
    b = 1
    If Not IsNull(RS.Fields(0).Value) Then b = RS.Fields(0).Value + 1
    a = "insert into m_StockBeli(IdStock, KodeJenis, Jenis, Satuan, NamaBarang) values(" & _
        b & _
        "," & col1("Kode Jenis").Value & _
        ",'" & col1("NamaJenis").Value & _
        "','" & col1("Satuan").Value & _
        "','" & esc(col1("Nama Barang").Value) & "')"
    If ExecMe(a) = 0 Then
        MsgBox "Stock Sudah Ada"
        Exit Sub
    End If
    col1("IdStock").Value = b
End Sub

Private Sub IText_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        GetDetail
    ElseIf KeyCode = 112 Then
        If IText(Index).Tag = "NoContract" Then
            If IText(3) = "0" Then
                HelpMe "No Contract Not Closed", Me
            ElseIf IText(3) = "1" Then
                HelpMe "No Contract Closed", Me
            Else
                HelpMe "No Contract All", Me
            End If
        ElseIf IText(Index).Tag = "Closed" Then
            HelpMe "Closed", Me
        End If
        IText(Index).Cancel = True
    End If
End Sub

Sub FormHelpKeyDown(ByVal tStr As String)
    ActiveControl.Text = tStr
    IText_KeyDown 0, 13, 0
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo err
    col1("updated").Value = 1
    a = col1(ColIndex).Caption
    If a = "IdStock" Then
        SetOtherRowData col1("IdStock").Value
    ElseIf a = "Order" Then
        If col1("Datang").Value = "" Then col1("Datang").Value = 0
        col1("Sisa").Value = col1("Order").Value - col1("Datang").Value
    End If
err:
End Sub

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)
    If col1("PO").Value <> "" Or Not IsNull(col1("PO").Value) Then
        Cancel = True
    Else
        If MsgBox("Yakin mau Hapus?", vbYesNo) = vbNo Then Exit Sub
        nopr = zerofill(col1("No PR").Value, 5) & "/" & zerofill(fKodeDepartemen, 2) & Mid(col1("Tanggal").Text, 3)
        a = "delete from t_PR where NoPR='" & esc(nopr) & "'"
        ExecMe a
    End If
End Sub

Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
    Response = False
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("Sisa").ColIndex) <= 0 Then
        RowStyle.ForeColor = vbBlue
    End If
    If x(Bookmark, col1("updated").ColIndex) > 0 Then
        RowStyle.BackColor = vbYellow
    End If
End Sub


Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo err
    TDBGrid1.Columns("No PR").Tag = "Decimal"
    TDBGridKeyPress TDBGrid1, KeyAscii
    'If col1("Order").ColIndex = TDBGrid1.Col Then
    '    If KeyAscii = 46 Then
    '       KeyAscii = 44
    '    ElseIf KeyAscii = 8 Then
    '        Exit Sub
    '    ElseIf KeyAscii < 44 Or KeyAscii > 57 Then
    '        KeyAscii = 0
    '    End If
    'End If
err:
End Sub

Private Sub TDBDropDown2_DropDownClose()
On Error GoTo err
    If TDBDropDown2.Bookmark > -1 Then
        col1("Nama Barang").Value = TDBDropDown2.Columns("Nama").Value
        col1("IdStock").Value = TDBDropDown2.Columns("IdStock").Value
        col1("Satuan").Value = TDBDropDown2.Columns("Satuan").Value
    Else
        'col1("Nama Barang").Value = ""
        col1("IdStock").Value = ""
        col1("Satuan").Value = ""
    End If
err:
End Sub
Private Sub TDBDropDown1_DropDownClose()
On Error GoTo err
    If TDBDropDown1.Bookmark > -1 Then
        col1("NamaJenis").Value = TDBDropDown1.Columns("Jenis").Value
    Else
        col1("NamaJenis").Value = ""
    End If
err:
End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
On Error Resume Next
    If fKodeDepartemen.ListIndex = -1 Then
        Cancel = True
        Exit Sub
    End If
    If col1(ColIndex).Caption <> "Closed" And (IsNull(col1("PO").Value) Or col1("PO").Value <> "") Then
        Cancel = True
        Exit Sub
    End If
    a = col1(ColIndex).Caption
    If a = "Kode Jenis" Then
        col1("Nama Barang").Value = ""
        If Not LJenis Then
            a = "select distinct KodeJenis, Jenis from m_stockBeli order by KodeJenis"
            query a
            z1.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
            z1.DeleteRows 0
            If RS.RecordCount > 0 Then z1.LoadRows RS.GetRows
            TDBDropDown1.Rebind
            LJenis = True
        End If
    ElseIf a = "Nama Barang" Then
        If col1("Kode Jenis").Value = "" Then Exit Sub
        a = "select NamaBarang, Satuan, IdStock from m_StockBeli where KodeJenis=" & col1("Kode Jenis").Value & " order by NamaBarang"
        query a
        z2.ReDim 0, 0, 0, TDBDropDown2.Columns.Count - 1
        z2.DeleteRows 0
        If RS.RecordCount > 0 Then z2.LoadRows RS.GetRows
        TDBDropDown2.Rebind
    End If
End Sub

Sub UpdateDB()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) <> 0 Then
            If cD(x(i, col1("Tanggal").ColIndex)) <> "A" Then
                If x(i, col1("IdStock").ColIndex) = "" Then
                    MsgBox "Stock Tidak ada"
                    GoTo err
                End If
                nopr = zerofill(x(i, col1("No PR").ColIndex), 5) & "/" & _
                    zerofill(fKodeDepartemen, 2) & "/" & Mid(x(i, col1("Tanggal").ColIndex), 3, 2) & "/" & Mid(x(i, col1("Tanggal").ColIndex), 5, 2)
                If x(i, col1("updated").ColIndex) = 1 Then
                    If x(i, col1("nopr2").ColIndex) <> "" Then
                        a = "update t_PR set NoPR='" & nopr & _
                        "', TanggalPR=" & cD(x(i, col1("Tanggal").ColIndex)) & _
                        ", TanggalKirim=" & cD(x(i, col1("Tgl Kirim").ColIndex)) & _
                        ", IdStock=" & x(i, col1("IdStock").ColIndex) & _
                        ", QTYOrder=" & cNum(x(i, col1("Order").ColIndex)) & _
                        ", KetPR='" & x(i, col1("Ket. PR").ColIndex) & _
                        "', Closed=" & IIf(x(i, col1("Closed").ColIndex) = -1, 1, 0) & _
                        " where NoPR='" & esc(x(i, col1("nopr2").ColIndex)) & "'"
                    Else
                        a = "insert into t_PR(NoPR,TanggalPR,TanggalKirim,IdStock,QTYOrder, KetPR, Dept) values(" & _
                        "'" & nopr & _
                        "'," & cD(x(i, col1("Tanggal").ColIndex)) & _
                        "," & cD(x(i, col1("Tgl Kirim").ColIndex)) & _
                        "," & x(i, col1("IdStock").ColIndex) & _
                        "," & cNum(x(i, col1("Order").ColIndex)) & _
                        ",'" & x(i, col1("Ket. PR").ColIndex) & _
                        "','" & fDepartemen & "')"
                    End If
                    If ExecMe(a) = 0 Then
                        TDBGrid1.Bookmark = i
                        TDBGrid1.SetFocus
                        GoTo err
                    End If
                End If
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
    DoEvents
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If LastCol = -1 Then Exit Sub
    If IsNull(TDBGrid1.Bookmark) Then TDBGrid1.Row = 0
    fKetNama = Replace(col1("Nama Barang").Value, "&", "&&")
    If IsNull(TDBGrid1.Row) Then TDBGrid1.Row = 0
    If ActiveControl.Name <> "TDBGrid1" Then
        TDBGrid1.SetFocus
    End If
End Sub


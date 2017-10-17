VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Object = "{F6D22ACD-8630-4FE1-97C4-D56AB4AD4DEA}#1.0#0"; "UsrTrueCombo.ocx"
Begin VB.Form FormJurnalUmum 
   BackColor       =   &H00FFC0C0&
   Caption         =   "JURNAL UMUM"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Tag             =   "30"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Filter"
      Height          =   795
      Left            =   6060
      TabIndex        =   12
      Top             =   0
      Width           =   3555
      Begin VB.TextBox fFilterKet 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Tag             =   "Jenis"
         Top             =   420
         Width           =   1875
      End
      Begin VB.TextBox fFilterAccount 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Tag             =   "Jenis"
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No Account"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   975
      End
   End
   Begin UsrText.IText fKetNoAccount 
      Height          =   270
      Left            =   2520
      TabIndex        =   10
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
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
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2775
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4895
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=74216168"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=7237481"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=72,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   2160
      TabIndex        =   1
      Top             =   120
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
   Begin UsrText.IText fAwal 
      Height          =   270
      Left            =   1200
      TabIndex        =   0
      Top             =   120
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4683
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
      Columns(2).Caption=   "Tanggal"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No Account"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Ket Acc"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Debet"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Kredit"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Keterangan"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Pengupdate"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=49"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1429"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1349"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=49"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(3)._MinWidth=49"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=3598"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3519"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(4)._MinWidth=2359295"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2302"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=2328"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=2249"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(34)=   "Column(7).Width=4895"
      Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=4815"
      Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(38)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(8).Order=9"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=104,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
   End
   Begin UsrTrueCombo.ITrueCombo fNoAccount 
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   9495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No Account"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "FormJurnalUmum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim z As New XArrayDB
Dim col2 As TrueOleDBGrid80.Columns
Dim col1 As TrueOleDBGrid80.Columns
Dim LNoAccount As Boolean

Private Sub fAkhir_Validate(Cancel As Boolean)
    If cD(fAkhir) = "A" Then fAkhir = pServerDate
    LoadData
End Sub

Private Sub fFilterAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        HelpMe "No Account", Me
    ElseIf KeyCode = 13 Then
        LoadData
    End If
End Sub

Sub FormHelpKeyDown(ByVal tVal As String)
    LoadData
End Sub

Private Sub fFilterKet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then LoadData
End Sub

Private Sub fNoAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not LNoAccount Then
        LNoAccount = True
        Dim rs1() As Variant
        a = "select NoAccount, Deskripsi from m_ChartAccount where Child=0 order by NoAccount"
        query a
        rs1 = RS.GetRows
        fNoAccount.SetDB rs1
        fNoAccount.SetType "String"
    End If
End Sub

Private Sub fNoAccount_LostFocus()
On Error GoTo err
    fNoAccount = fNoAccount.GetData("No Account")
    fKetNoAccount = fNoAccount.GetData("Ket Acc")
    Set TDBDropDown1.Array = z
    Set col2 = TDBDropDown1.Columns
    z.ReDim 0, fNoAccount.ListCount - 1, 0, col2.Count - 1
    k = 0
    For i = 0 To fNoAccount.ListCount - 1
        If fNoAccount.zz(i, "No Account") <> fNoAccount Then
            z(k, 0) = fNoAccount.zz(i, "No Account")
            z(k, 1) = fNoAccount.zz(i, "Ket Acc")
            k = k + 1
        End If
    Next
    z.ReDim 0, k - 1, 0, col2.Count - 1
    TDBDropDown1.Rebind
    LoadData
    Exit Sub
err:
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LNoAccount = False
End Sub

Private Sub LoadData()
On Error Resume Next
    If Not fNoAccount.Validate Then
        Exit Sub
    End If
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    a = "select ckey, DebetAcc, KetDebetAcc, KreditAcc, KetKreditAcc, Nilai, Tanggal, Keterangan, Pengupdate from t_Account where Tanggal>=" & cD(fAwal) & " and Tanggal<=" & cD(fAkhir) & " and (DebetAcc='" & esc(fNoAccount) & "' or KreditAcc='" & esc(fNoAccount) & "') and Keterangan like '%" & esc(fFilterKet) & "%'"
    If Not cekValid("SEE", Tag, True) Then a = "select 1 from AutoUpdate where 1=0"
    query a
    If RS.RecordCount > 0 Then x.ReDim 0, RS.RecordCount - 1, 0, col1.Count - 1
    k = 0
    For i = 0 To RS.RecordCount - 1
        If fNoAccount = RS.Fields("DebetAcc").Value Then
            NoAccount = RS.Fields("KreditAcc").Value
            KetAcc = RS.Fields("KetKreditAcc").Value
            debet = ""
            kredit = RS.Fields("Nilai").Value
        Else
            NoAccount = RS.Fields("DebetAcc").Value
            KetAcc = RS.Fields("KetDebetAcc").Value
            debet = RS.Fields("Nilai").Value
            kredit = ""
        End If
        If Trim(fFilterAccount) = "" Or NoAccount = fFilterAccount Then
            x(k, col1("updated").ColIndex) = 0
            x(k, col1("ckey").ColIndex) = RS.Fields("cKey").Value
            x(k, col1("Tanggal").ColIndex) = cTanggal2(RS.Fields("Tanggal").Value)
            x(k, col1("No Account").ColIndex) = NoAccount
            x(k, col1("Ket Acc").ColIndex) = KetAcc
            x(k, col1("Debet").ColIndex) = debet
            x(k, col1("Kredit").ColIndex) = kredit
            x(k, col1("Keterangan").ColIndex) = RS.Fields("Keterangan").Value
            x(k, col1("Pengupdate").ColIndex) = RS.Fields("Pengupdate").Value
            k = k + 1
        End If
        RS.MoveNext
    Next
    If k > 0 Then x.ReDim 0, k - 1, 0, col1.Count - 1
    TDBGrid1.Rebind
    TDBGrid1.SetFocus
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    Set TDBGrid1.Array = x
    TDBGrid1.Rebind
    Set col1 = TDBGrid1.Columns
    col1("Tanggal").Tag = "Date"
    col1("Debet").Tag = "Decimal"
    col1("Kredit").Tag = "Decimal"
    TDBGridLoad TDBGrid1
    TDBGrid1.FetchRowStyle = True
    TDBGridSetVisible TDBGrid1, "updated@ckey", False
    col1("No Account").DropDown = TDBDropDown1
    col1("No Account").AutoDropDown = True
    TDBGridSetLock TDBGrid1, "Ket Acc@Pengupdate", True
    fNoAccount.SetHeader "No Account@Ket Acc"
    fNoAccount.SetWidth "1500@2500"
    fNoAccount.SetType "String@String"
    fNoAccount.ZOrder 0
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
    fAwal = pServerDate
    fAkhir = fAwal
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    fKet.Width = TDBGrid1.Width
End Sub

Private Sub fSave_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) = 1 Then
            If x(i, col1("Ket Acc").ColIndex) = "" Then
                MsgBox "No Account Harus Diisi"
                GoTo err
            End If
            If x(i, col1("Debet").ColIndex) = "" Then
                DebetAcc = fNoAccount
                ketdebetacc = fKetNoAccount
                kreditacc = x(i, col1("No Account").ColIndex)
                ketkreditacc = x(i, col1("Ket Acc").ColIndex)
                Nilai = x(i, col1("Kredit").ColIndex)
            Else
                DebetAcc = x(i, col1("No Account").ColIndex)
                ketdebetacc = x(i, col1("Ket Acc").ColIndex)
                kreditacc = fNoAccount
                ketkreditacc = fKetNoAccount
                Nilai = x(i, col1("Debet").ColIndex)
            End If
            If x(i, col1("cKey").ColIndex) = "" Then
                a = "insert into t_Account(DebetAcc, KetDebetAcc, KreditAcc, KetKreditAcc, Nilai, Tanggal, Keterangan, Pengupdate) values('" & _
                    DebetAcc & _
                    "','" & ketdebetacc & _
                    "','" & kreditacc & _
                    "','" & ketkreditacc & _
                    "'," & cNum(Nilai) & _
                    "," & cD(x(i, col1("Tanggal").ColIndex)) & _
                    ",'" & esc(x(i, col1("Keterangan").ColIndex)) & _
                    "','" & esc(pUsr) & "')"
            Else
                a = "update t_Account set " & _
                    "DebetAcc='" & DebetAcc & _
                    "', KetDebetAcc='" & ketdebetacc & _
                    "', KreditAcc='" & kreditacc & _
                    "', KetKreditAcc='" & ketkreditacc & _
                    "', Nilai=" & cNum(Nilai) & _
                    ", Tanggal=" & cD(x(i, col1("Tanggal").ColIndex)) & _
                    ", Keterangan='" & esc(x(i, col1("Keterangan").ColIndex)) & _
                    "', Pengupdate='" & esc(pUsr) & _
                    "' where cKey=" & x(i, col1("cKey").ColIndex)
            End If
            If ExecMe(a) = 0 Then GoTo err
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoEvents
    LoadData
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBDropDown1_DropDownClose()
On Error Resume Next
    col1("No Account").Value = col2("No Account").Value
    col1("Ket Acc").Value = col2("Ket Acc").Value
End Sub

Private Sub TDBDropDown1_Paint()
    TDBGrid1.SelLength = Len(TDBGrid1.Text)
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    col1("updated").Value = 1
End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If TDBGrid1.AddNewMode <> dbgNoAddNew Then
        If Not cekValid("NEW", Tag) Then GoTo err
    End If
    If col1("cKey").Value <> "" Then
        If Not cekValid("EDIT", Tag) Then
            GoTo err
        ElseIf pUpdateHargaSC = 0 And pUsr <> col1("Pengupdate").Value Then
            MsgBox "Tidak Berhak!!!"
            GoTo err
        End If
    End If
    a = col1(ColIndex).Caption
    If a = "Debet" Then
        If col1("Kredit").Value <> "" Then
            GoTo err
        End If
    ElseIf a = "Kredit" Then
        If col1("Debet").Value <> "" Then
            GoTo err
        End If
    End If
    Exit Sub
err:
    Cancel = True
End Sub

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)
    If Not cekValid("DELETE", Tag) Then
        GoTo err
    ElseIf Not pUpdateHargaSC And pUsr <> col1("Pengupdate").Value Then
        MsgBox "Tidak Berhak!!!"
        GoTo err
    End If
    b = MsgBox("Yakin Mau Hapus?", vbYesNo)
    If b = vbNo Then GoTo err
    a = "delete from t_Account where cKey=" & col1("cKey").Value
    If ExecMe(a) = 0 Then
        Cancel = True
        MsgBox "GAGAL"
    End If
    Exit Sub
err:
    Cancel = True
End Sub

Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
    Response = False
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("updated").ColIndex) = 1 Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    fKet = col1(TDBGrid1.Col).Text
End Sub


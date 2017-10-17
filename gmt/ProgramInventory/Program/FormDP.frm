VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Begin VB.Form FormDP 
   BackColor       =   &H00FFC0C0&
   Caption         =   "DP"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Tag             =   "43"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "COPY KE CLIPBOARD"
      Height          =   375
      Left            =   5220
      TabIndex        =   16
      Top             =   780
      Width           =   1995
   End
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   4200
      TabIndex        =   12
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
   Begin UsrText.IText fAwal 
      Height          =   270
      Left            =   3240
      TabIndex        =   11
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
   Begin UsrText.IText fFCustomer 
      Height          =   270
      Left            =   720
      TabIndex        =   10
      Top             =   840
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
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton fSemua 
      Caption         =   "SEMUA"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton fBelumTerpakai 
      Caption         =   "BELUM TERPAKAI"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   2295
      Left            =   6720
      TabIndex        =   4
      Top             =   1920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4048
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
      Columns(1).Caption=   "Tanggal SC"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Curr"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3387"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3307"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=128"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1799"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1720"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=74974144"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1138"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1058"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=74975264"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=240,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   3735
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6588
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama Customer"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Kode Customer"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=7488"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=7408"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2064"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1984"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1)._MinWidth=76170420"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=188,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5953
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
      Columns(2).Caption=   "Kode Customer"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tanggal"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Customer"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "No SC"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Mata Uang"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Nilai DP"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Terpakai"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Keterangan"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1164"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1085"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1746"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1667"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=4339"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=4260"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=3122"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3043"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1640"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1561"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2381"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2302"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2302"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2223"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=4392"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=4313"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
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
      DeadAreaBackColor=   13160660
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=36,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=78,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=82,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(70)  =   "Named:id=33:Normal"
      _StyleDefs(71)  =   ":id=33,.parent=0"
      _StyleDefs(72)  =   "Named:id=34:Heading"
      _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=34,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=35:Footing"
      _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=36:Selected"
      _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=37:Caption"
      _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(81)  =   "Named:id=38:HighlightRow"
      _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=39:EvenRow"
      _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(85)  =   "Named:id=40:OddRow"
      _StyleDefs(86)  =   ":id=40,.parent=33"
      _StyleDefs(87)  =   "Named:id=41:RecordSelector"
      _StyleDefs(88)  =   ":id=41,.parent=34"
      _StyleDefs(89)  =   "Named:id=42:FilterBar"
      _StyleDefs(90)  =   ":id=42,.parent=33"
   End
   Begin UsrText.IText fFind 
      Height          =   270
      Left            =   5160
      TabIndex        =   7
      Top             =   240
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Find"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Click Update untuk Update Data DP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "FormDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid80.Columns
Dim x As New XArrayDB
Dim y1 As New XArrayDB
Dim y2 As New XArrayDB
Dim coly1 As TrueOleDBGrid80.Columns
Dim coly2 As TrueOleDBGrid80.Columns
Dim LCustomer As Boolean

Private Function MyFilter() As String
    MyFilter = ""
    If Trim(fFCustomer) <> "" Then MyFilter = " and Nama like '%" & fFCustomer & "%'"
    If cD(fAwal) <> "A" Then MyFilter = MyFilter & " and Tanggal>=" & cD(fAwal)
    If cD(fAkhir) <> "A" Then MyFilter = MyFilter & " and Tanggal<=" & cD(fAkhir)
End Function

Private Sub cmdCopy_Click()
    CopyGrid TDBGrid1
End Sub

Private Sub fAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If fBelumTerpakai.FontBold = True Then DoQuery Else DoQuery2
    End If
End Sub

Private Sub fBelumTerpakai_Click()
    DoQuery
End Sub

Private Sub fDelete_Click()
On Error Resume Next
    If Not cekValid("DELETE", Tag) Then Exit Sub
    b = MsgBox("Yakin Mau Hapus", vbYesNo)
    If b = vbNo Then Exit Sub
    a = "delete from t_DP where IdDP=" & col1("cKey").Value
    ExecMe a
    TDBGrid1.Delete
End Sub

Private Sub fFCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If fBelumTerpakai.FontBold = True Then DoQuery Else DoQuery2
    End If
End Sub

Private Sub fFind_KeyDown(KeyCode As Integer, Shift As Integer)
    FindKeyDown KeyCode, TDBGrid1, x, fFind
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Sub DoQuery()
    fBelumTerpakai.FontBold = True
    fSemua.FontBold = False
    If Not cekValid("SEE", Tag) Then
        x.ReDim 0, 0, 0, col1.Count - 1
        x.DeleteRows 0
        TDBGrid1.Rebind
        Exit Sub
    End If
    a = "select 0, IdDP, KodeCustomer, Tanggal, Nama, NoSC, t_DP.MataUang, NilaiDP, Terpakai, Keterangan from t_DP left join m_Customer on t_DP.KodeCustomer=m_Customer.Kode where NilaiDP>Terpakai " & MyFilter & " order by Tanggal Desc, Nama"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    For i = 0 To x.UpperBound(1)
        x(i, col1("Tanggal").ColIndex) = cTanggal2(x(i, col1("Tanggal").ColIndex))
    Next
    TDBGrid1.Rebind
End Sub

Sub DoQuery2()
    fBelumTerpakai.FontBold = False
    fSemua.FontBold = True
    a = "select 0, IdDP, KodeCustomer, Tanggal, Nama, NoSC, t_DP.MataUang, NilaiDP, Terpakai, Keterangan from t_DP left join m_Customer on t_DP.KodeCustomer=m_Customer.Kode where 1=1 " & MyFilter & " order by Tanggal Desc, Nama"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    For i = 0 To x.UpperBound(1)
        x(i, col1("Tanggal").ColIndex) = cTanggal2(x(i, col1("Tanggal").ColIndex))
    Next
    TDBGrid1.Rebind
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    DoQuery
    col1("updated").Visible = False
    col1("cKey").Visible = False
    col1("Kode Customer").Visible = False
    TDBGrid1.FetchRowStyle = True
    col1("Nilai DP").NumberFormat = "FormatText Event"
    col1("Terpakai").NumberFormat = "FormatText Event"
    col1("Nilai DP").Tag = "Decimal"
    col1("Tanggal").NumberFormat = "Edit Mask"
    col1("Tanggal").EditMask = "##/##/##"
    col1("Nilai DP").Alignment = dbgRight
    col1("Terpakai").Alignment = dbgRight
    TDBGrid1.AllowAddNew = True
    Set coly1 = TDBDropDown1.Columns
    Set coly2 = TDBDropDown2.Columns
    Set TDBDropDown1.Array = y1
    Set TDBDropDown2.Array = y2
    col1("Nama Customer").DropDown = TDBDropDown1
    col1("No SC").DropDown = TDBDropDown2
    col1("Nama Customer").AutoDropDown = True
    col1("No SC").AutoDropDown = True
    col1("Terpakai").Locked = True
    fAkhir = pServerDate
    fAwal = add_tanggal(fAkhir, -89)
End Sub

Private Sub fSemua_Click()
    DoQuery2
End Sub

Private Sub TDBDropDown1_DropDownClose()
On Error Resume Next
    If TDBDropDown1.Bookmark < 0 Or TDBDropDown1.Bookmark = "" Then
        col1("Kode Customer").Value = ""
        col1("Nama Customer").Value = ""
    Else
        col1("Kode Customer").Value = coly1("Kode Customer").Value
    End If
End Sub

Private Sub TDBDropDown1_Paint()
    TDBGrid1.SelLength = Len(TDBGrid1.Text)
End Sub

Private Sub TDBDropDown2_DropDownClose()
    If TDBDropDown2.Bookmark < 0 Or TDBDropDown2.Bookmark = "" Then
        col1("No SC").Value = ""
        col1("Mata Uang").Value = ""
    Else
        col1("Mata Uang").Value = coly2("Curr").Value
    End If
End Sub

Private Sub TDBDropDown2_Paint()
    TDBGrid1.SelLength = Len(TDBGrid1.Text)
End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
On Error Resume Next
    If TDBGrid1.AddNewMode <> dbgNoAddNew Then
        If Not cekValid("NEW", Tag) Then GoTo err
    End If
    If col1("cKey").Value <> "" Then
        If Not cekValid("EDIT", Tag) Then GoTo err
    End If
    If col1("Terpakai").Value > 0 And col1("Terpakai").Value <> "" Then GoTo err
    a = col1(ColIndex).Caption
    If a = "Nama Customer" Then
        col1("No SC").Value = ""
        If Not LCustomer Then
            a = "select Nama, Kode from m_Customer where IsActive=1 order by Nama"
            query a
            y1.ReDim 0, 0, 0, coly1.Count - 1
            y1.DeleteRows 0
            If RS.RecordCount > 0 Then y1.LoadRows RS.GetRows
            TDBDropDown1.Rebind
            LCustomer = True
        End If
    ElseIf a = "No SC" Then
        a = "select NoSC, TanggalSC, MataUang from t_SC~ where Status=0 and Disetujui=1 and Kode=" & col1("Kode Customer").Value & " order by NoSC"
        query a
        y2.ReDim 0, 0, 0, coly2.Count - 1
        y2.DeleteRows 0
        If RS.RecordCount > 0 Then y2.LoadRows RS.GetRows
        For i = 0 To y2.UpperBound(1)
            y2(i, coly2("Tanggal SC").ColIndex) = cTanggal(y2(i, coly2("Tanggal SC").ColIndex))
        Next
        TDBDropDown2.Rebind
    End If
    Exit Sub
err:
    Cancel = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) = "1" Then
            If x(i, col1("cKey").ColIndex) = "" Then
                a = "insert into t_DP(KodeCustomer, NilaiDP, Tanggal, NoSC, Keterangan, Terpakai, MataUang) values('" & _
                    x(i, col1("Kode Customer").ColIndex) & _
                    "','" & cNum(x(i, col1("Nilai DP").ColIndex)) & _
                    "'," & cD(x(i, col1("Tanggal").ColIndex)) & _
                    ",'" & esc(x(i, col1("No SC").ColIndex)) & _
                    "','" & esc(x(i, col1("Keterangan").ColIndex)) & _
                    "'," & cNum(x(i, col1("Terpakai").ColIndex)) & _
                    ",'" & esc(x(i, col1("Mata Uang").ColIndex)) & "')"
                If ExecMe(a) = 0 Then GoTo err
            Else
                a = "update t_DP set KodeCustomer=" & x(i, col1("Kode Customer").ColIndex) & _
                    ", NilaiDP=" & cNum(x(i, col1("Nilai DP").ColIndex)) & _
                    ", Tanggal=" & cD(x(i, col1("Tanggal").ColIndex)) & _
                    ", NoSC='" & esc(x(i, col1("No SC").ColIndex)) & _
                    "', Keterangan='" & esc(x(i, col1("Keterangan").ColIndex)) & _
                    "', MataUang='" & esc(x(i, col1("Mata Uang").ColIndex)) & _
                    "' where IdDP=" & x(i, col1("cKey").ColIndex)
                If ExecMe(a) = 0 Then GoTo err
            End If
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

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
    col1("updated").Value = "1"
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("updated").ColIndex) = "1" Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
On Error Resume Next
    a = col1(ColIndex).Caption
    If a = "Nilai DP" Or a = "Terpakai" Then
        Value = cDecimal(Value)
    End If
End Sub


Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        FindKeyDown 13, TDBGrid1, x, fFind
    End If
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub


VERSION 5.00
Object = "{E2D3646A-2684-4DDE-BE47-3323E01328EE}#1.0#0"; "UsrText.ocx"
Object = "{EDF3EDD5-ECAB-11DA-99A6-000B6A30ACAC}#1.0#0"; "UsrTrueCombo.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FormPassword 
   Caption         =   "SETTING PASSWORD"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "36"
   WindowState     =   2  'Maximized
   Begin VB.CheckBox fUpdateHargaSC 
      Caption         =   "Spesial Update"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   840
      Width           =   1935
   End
   Begin UsrTrueCombo.ITrueCombo fDepartemen 
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.ComboBox fHeader 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   3975
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4471
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tag"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   582
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5133"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5054"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1244"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1164"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1)._MinWidth=66529052"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1984"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1905"
      Splits(0)._ColumnProps(15)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=66528684"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=3"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13,.alignment=3"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=40,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=37,.parent=12,.alignment=3"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=38,.parent=13,.alignment=3"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=39,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12,.alignment=3"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
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
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin UsrText.IText fUsr 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
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
   Begin UsrText.IText fPwd 
      Height          =   270
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   476
      DataType        =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
   End
   Begin UsrText.IText fNama 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Top             =   840
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3836
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Command"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Masuk"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   4
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "See"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   4
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Add"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Edit"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Delete"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   4
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Print"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Tag"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   582
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5133"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5054"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=-2147483640"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1270"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1191"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=218152112"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1005"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=926"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=1164"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1085"
      Splits(0)._ColumnProps(21)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=1164"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1085"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=1191"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1111"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=1270"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1191"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=8708"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=73,.parent=2,.namedParent=75"
      _StyleDefs(17)  =   "FilterBarStyle:id=76,.parent=1,.namedParent=78"
      _StyleDefs(18)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.alignment=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=74,.parent=73"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=77,.parent=76"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=3"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13,.alignment=3"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=40,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=37,.parent=12,.alignment=3"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=38,.parent=13,.alignment=3"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=39,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=72,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=69,.parent=12,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=70,.parent=13,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=71,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12,.alignment=3"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13,.alignment=3"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=44,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=41,.parent=12,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=42,.parent=13,.alignment=3"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=43,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=48,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=45,.parent=12,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=46,.parent=13,.alignment=3"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=47,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=52,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=49,.parent=12,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=50,.parent=13,.alignment=3"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=51,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=56,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=53,.parent=12,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=54,.parent=13,.alignment=3"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=55,.parent=15"
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
      _StyleDefs(79)  =   "Named:id=75:RecordSelector"
      _StyleDefs(80)  =   ":id=75,.parent=30"
      _StyleDefs(81)  =   "Named:id=78:FilterBar"
      _StyleDefs(82)  =   ":id=78,.parent=29"
   End
   Begin UsrText.IText fTipe 
      Height          =   270
      Left            =   3720
      TabIndex        =   4
      Top             =   480
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
   Begin VB.Label Label7 
      Caption         =   "Default Tipe"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Departemen"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "LIHAT REKAP OPERSIONAL"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Header"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Nama"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User ID"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FormPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim y As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim col2 As TrueOleDBGrid80.Columns
Dim Bag1 As Long
Dim Bag2 As Long
Dim BagSee1 As Long
Dim BagSee2 As Long
Dim BagAdd1 As Long
Dim BagEdit1 As Long
Dim BagDelete1 As Long
Dim BagPrint1 As Long
Dim BagAdd2 As Long
Dim BagEdit2 As Long
Dim BagDelete2 As Long
Dim BagPrint2 As Long
Dim Bag3 As Long
Dim Bag4 As Long

Dim CurVal1 As Long
Dim CurValSee1 As Long
Dim CurValAdd1 As Long
Dim CurValEdit1 As Long
Dim CurValDelete1 As Long
Dim CurValPrint1 As Long
Dim CurVal2 As Long
Dim CurValSee2 As Long
Dim CurValAdd2 As Long
Dim CurValEdit2 As Long
Dim CurValDelete2 As Long
Dim CurValPrint2 As Long
Dim mPass As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub fDelete_Click()
    a = "delete from m_users where usr='" & fUsr & "'"
    ExecMe a
    MsgBox "User Deleted"
End Sub
Private Sub fHeader_Validate(Cancel As Boolean)
    a = "select Command,0,0,0,0,0,0,Tag from m_DaftarMenu where Header='" & fHeader & "' order by Tag"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    Dim v As Long
    Dim vSee As Long
    Dim vAdd As Long
    Dim vEdit As Long
    Dim vDelete As Long
    Dim vPrint As Long
    CurVal1 = 0
    CurValSee1 = 0
    CurValAdd1 = 0
    CurValEdit1 = 0
    CurValDelete1 = 0
    CurValPrint1 = 0
    CurVal2 = 0
    CurValSee2 = 0
    CurValAdd2 = 0
    CurValEdit2 = 0
    CurValDelete2 = 0
    CurValPrint2 = 0
    For i = 0 To x.UpperBound(1)
        If x(i, col1("Tag").ColIndex) < 31 Then
            a = 2 ^ x(i, col1("Tag").ColIndex)
            v = a And Bag1
            vSee = a And BagSee1
            vAdd = a And BagAdd1
            vEdit = a And BagEdit1
            vDelete = a And BagDelete1
            vPrint = a And BagPrint1
            CurVal1 = CurVal1 + v
            CurValSee1 = CurValSee1 + vSee
            CurValAdd1 = CurValAdd1 + vAdd
            CurValEdit1 = CurValEdit1 + vEdit
            CurValDelete1 = CurValDelete1 + vDelete
            CurValPrint1 = CurValPrint1 + vPrint
        Else
            a = 2 ^ (x(i, col1("Tag").ColIndex) - 31)
            v = a And Bag2
            vSee = a And BagSee2
            vAdd = a And BagAdd2
            vEdit = a And BagEdit2
            vDelete = a And BagDelete2
            vPrint = a And BagPrint2
            CurVal2 = CurVal2 + v
            CurValSee2 = CurValSee2 + vSee
            CurValAdd2 = CurValAdd2 + vAdd
            CurValEdit2 = CurValEdit2 + vEdit
            CurValDelete2 = CurValDelete2 + vDelete
            CurValPrint2 = CurValPrint2 + vPrint
        End If
        x(i, col1("Masuk").ColIndex) = IIf(v = 0, 0, -1)
        x(i, col1("See").ColIndex) = IIf(vSee = 0, 0, -1)
        x(i, col1("Add").ColIndex) = IIf(vAdd = 0, 0, -1)
        x(i, col1("Edit").ColIndex) = IIf(vEdit = 0, 0, -1)
        x(i, col1("Delete").ColIndex) = IIf(vDelete = 0, 0, -1)
        x(i, col1("Print").ColIndex) = IIf(vPrint = 0, 0, -1)
    Next
    TDBGrid1.Rebind
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Load()
    fDepartemen.ZOrder 0
    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    Set col2 = TDBGrid2.Columns
    a = "select distinct Header from m_DaftarMenu order by Header"
    query a
    fHeader.Clear
    For i = 0 To RS.RecordCount - 1
        fHeader.List(i) = RS.Fields(0).Value
        RS.MoveNext
    Next
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    Set TDBGrid1.Array = x
    Set TDBGrid2.Array = y
    a = "select Lihat1, 0, Tag from m_Lihat order by Lihat1"
    query a
    y.ReDim 0, 0, 0, col2.Count - 1
    y.DeleteRows 0
    If RS.RecordCount > 0 Then y.LoadRows RS.GetRows
    TDBGrid2.Rebind
    fDepartemen.SetHeader "Departemen@*Kode"
    fDepartemen.SetWidth "2000@1000"
    fDepartemen.SetType "String@Integer"
    a = "select Departemen, KdDept from m_Departemen order by Departemen"
    query a
    Dim rs1() As Variant
    rs1 = RS.GetRows
    fDepartemen.SetDB rs1
End Sub

Private Sub fSave_Click()
On Error GoTo err
    BeginTransaction
    a = "delete from m_users where usr='" & fUsr & "'" 'Masuk Ke Database berdasarkan Tag-nya
    ExecMe a
    Dim Val1 As Long
    Dim ValSee1 As Long
    Dim ValAdd1 As Long
    Dim ValEdit1 As Long
    Dim ValDelete1 As Long
    Dim ValPrint1 As Long
    Dim Val2 As Long
    Dim ValSee2 As Long
    Dim ValAdd2 As Long
    Dim ValEdit2 As Long
    Dim ValDelete2 As Long
    Dim ValPrint2 As Long
    Val1 = Bag1 - CurVal1
    ValSee1 = BagSee1 - CurValSee1
    ValAdd1 = BagAdd1 - CurValAdd1
    ValEdit1 = BagEdit1 - CurValEdit1
    ValDelete1 = BagDelete1 - CurValDelete1
    ValPrint1 = BagPrint1 - CurValPrint1
    Val2 = Bag2 - CurVal2
    ValSee2 = BagSee2 - CurValSee2
    ValAdd2 = BagAdd2 - CurValAdd2
    ValEdit2 = BagEdit2 - CurValEdit2
    ValDelete2 = BagDelete2 - CurValDelete2
    ValPrint2 = BagPrint2 - CurValPrint2
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("Tag").ColIndex) < 31 Then
            a = 2 ^ x(i, col1("Tag").ColIndex)
            Val1 = Val1 + IIf(x(i, col1("Masuk").ColIndex) = 0, 0, a)
            ValSee1 = ValSee1 + IIf(x(i, col1("See").ColIndex) = 0, 0, a)
            ValAdd1 = ValAdd1 + IIf(x(i, col1("Add").ColIndex) = 0, 0, a)
            ValEdit1 = ValEdit1 + IIf(x(i, col1("Edit").ColIndex) = 0, 0, a)
            ValDelete1 = ValDelete1 + IIf(x(i, col1("Delete").ColIndex) = 0, 0, a)
            ValPrint1 = ValPrint1 + IIf(x(i, col1("Print").ColIndex) = 0, 0, a)
        Else
            a = 2 ^ (x(i, col1("Tag").ColIndex) - 31)
            Val2 = Val2 + IIf(x(i, col1("Masuk").ColIndex) = 0, 0, a)
            ValSee2 = ValSee2 + IIf(x(i, col1("See").ColIndex) = 0, 0, a)
            ValAdd2 = ValAdd2 + IIf(x(i, col1("Add").ColIndex) = 0, 0, a)
            ValEdit2 = ValEdit2 + IIf(x(i, col1("Edit").ColIndex) = 0, 0, a)
            ValDelete2 = ValDelete2 + IIf(x(i, col1("Delete").ColIndex) = 0, 0, a)
            ValPrint2 = ValPrint2 + IIf(x(i, col1("Print").ColIndex) = 0, 0, a)
        End If
    Next
    TDBGrid2.Update
    Bag3 = 0
    Bag4 = 0
    For i = 0 To y.UpperBound(1)
        a = Val(y(i, 2))
        b = IIf(y(i, 1) <> 0, 1, 0)
        If i < 31 Then
            Bag3 = Bag3 + b * 2 ^ a
        Else
            Bag4 = Bag4 + b * 2 ^ (a - 31)
        End If
    Next
    a = "insert into m_users(usr,pwd,nama,kddept,Tipe,bag1, bagSee1, bagAdd1, bagEdit1, bagDelete1, bagPrint1, bag2, bagSee2, bagAdd2, bagEdit2, bagDelete2, bagPrint2, bag3,bag4, UpdateHargaSC) values('" & _
        fUsr & _
        "','" & fPwd & _
        "','" & fNama & _
        "'," & fDepartemen.GetData("Kode") & _
        ",'" & fTipe & _
        "'," & Val1 & _
        "," & ValSee1 & _
        "," & ValAdd1 & _
        "," & ValEdit1 & _
        "," & ValDelete1 & _
        "," & ValPrint1 & _
        "," & Val2 & _
        "," & ValSee2 & _
        "," & ValAdd2 & _
        "," & ValEdit2 & _
        "," & ValDelete2 & _
        "," & ValPrint2 & _
        "," & Bag3 & _
        "," & Bag4 & _
        "," & fUpdateHargaSC & ")"
    If ExecMe(a) = 0 Then GoTo err
    CommitTransaction
    MsgBox "Password Saved"
    If pUsr = fUsr Then
        pBag1 = Val1
        pBagSee1 = ValSee1
        pBagAdd1 = ValAdd1
        pBagEdit1 = ValEdit1
        pBagDelete1 = ValDelete1
        pBagPrint1 = ValPrint1
        pBag2 = Val2
        pBagSee2 = ValSee2
        pBagAdd2 = ValAdd2
        pBagEdit2 = ValEdit2
        pBagDelete2 = ValDelete2
        pBagPrint2 = ValPrint2
        pBag3 = Bag3
        pBag4 = Bag4
        pUpdateHargaSC = fUpdateHargaSC
    End If
    fUsr_Validate False
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fUsr_Validate(Cancel As Boolean)
On Error Resume Next
    fPwd = ""
    fNama = ""
    Bag1 = 0
    Bag2 = 0
    BagSee1 = 0
    BagAdd1 = 0
    BagEdit1 = 0
    BagDelete1 = 0
    BagPrint1 = 0
    BagSee2 = 0
    BagAdd2 = 0
    BagEdit2 = 0
    BagDelete2 = 0
    BagPrint2 = 0
    Bag3 = 0
    Bag4 = 0
    a = "select top 1 usr,pwd,nama,kddept,tipe,bag1,bag2, bagSee1, bagAdd1, bagEdit1, bagDelete1, bagPrint1, bagSee2, bagAdd2, bagEdit2, bagDelete2, bagPrint2, bag3, bag4, UpdateHargaSC from m_users where usr='" & fUsr & "'"
    query a
    If RS.RecordCount > 0 Then
        fUsr = RS.Fields("usr").Value
        fPwd = RS.Fields("pwd").Value
        fNama = RS.Fields("nama").Value
        b = fDepartemen.ListCount - 1
        kddept = RS.Fields("kddept").Value
        For i = 0 To b
            If fDepartemen.zz(i, "Kode") = kddept Then
                fDepartemen = fDepartemen.zz(i, "Departemen")
                fDepartemen.SetListIndex i
                Exit For
            End If
        Next
        fTipe = RS.Fields("tipe").Value
        Bag1 = RS.Fields("bag1").Value
        Bag2 = RS.Fields("bag2").Value
        BagSee1 = RS.Fields("bagSee1").Value
        BagAdd1 = RS.Fields("bagAdd1").Value
        BagEdit1 = RS.Fields("bagEdit1").Value
        BagDelete1 = RS.Fields("bagDelete1").Value
        BagPrint1 = RS.Fields("bagPrint1").Value
        BagSee2 = RS.Fields("bagSee2").Value
        BagAdd2 = RS.Fields("bagAdd2").Value
        BagEdit2 = RS.Fields("bagEdit2").Value
        BagDelete2 = RS.Fields("bagDelete2").Value
        BagPrint2 = RS.Fields("bagPrint2").Value
        Bag3 = RS.Fields("bag3").Value
        Bag4 = RS.Fields("bag4").Value
    End If
    fUpdateHargaSC = RS.Fields("UpdateHargaSC").Value
    fHeader = ""
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
    For i = 0 To y.UpperBound(1)
        If i < 31 Then
            v = Bag3 And (2 ^ y(i, col2("Tag").ColIndex))
        Else
            v = Bag4 And (2 ^ (y(i, col2("Tag").ColIndex) - 31))
        End If
        y(i, 1) = IIf(v = 0, 0, -1)
    Next
    TDBGrid2.Rebind
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    a = TDBGrid1.Columns(ColIndex).Caption
    If a = "Masuk" Then
        v = IIf(TDBGrid1.Columns("Masuk").Value = 0, 0, -1)
        TDBGrid1.Columns("See").Value = v
        TDBGrid1.Columns("Add").Value = v
        TDBGrid1.Columns("Edit").Value = v
        TDBGrid1.Columns("Delete").Value = v
        TDBGrid1.Columns("Print").Value = v
    ElseIf TDBGrid1.Columns(ColIndex).Value = -1 Then
        TDBGrid1.Columns("Masuk").Value = -1
    End If
End Sub



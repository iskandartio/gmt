VERSION 5.00
Object = "{E2D3646A-2684-4DDE-BE47-3323E01328EE}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormJurnalPembayaran 
   BackColor       =   &H00FFC0C0&   Caption         =   "JURNAL PEMBAYARAN"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "31"
   WindowState     =   2  'Maximized
   Begin UsrText.IText fCaraBayar 
      Height          =   270
      Left            =   6360
      TabIndex        =   13
      Top             =   120
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
   Begin VB.CheckBox fBelumSet 
      BackColor       =   &H00FFC0C0&      Caption         =   "Belum Diset"
      Height          =   255
      Left            =   4020
      TabIndex        =   11
      Top             =   120
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2775
      Left            =   2280
      TabIndex        =   8
      Top             =   1800
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
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin UsrText.IText fAwal 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
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
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No STT"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Tanggal"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cara Bayar"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nama Bank"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "No Giro"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tanggal Lunas"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Nilai"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Supplier"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Kredit Acc"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Ket Acc"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "IdNo"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "updated"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=1769238625"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1323"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1244"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=163973724"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1614"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1535"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=164035200"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1693"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1614"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1826"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1746"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(4)._MinWidth=-1"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1746"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1667"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2275"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2196"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=3678"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=3598"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=1667"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=1588"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=3863"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=3784"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(45)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(48)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(49)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(11).Order=12"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=88,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=82,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
      _StyleDefs(78)  =   "Named:id=33:Normal"
      _StyleDefs(79)  =   ":id=33,.parent=0"
      _StyleDefs(80)  =   "Named:id=34:Heading"
      _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=34,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=35:Footing"
      _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=36:Selected"
      _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=37:Caption"
      _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(89)  =   "Named:id=38:HighlightRow"
      _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=39:EvenRow"
      _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=40:OddRow"
      _StyleDefs(94)  =   ":id=40,.parent=33"
      _StyleDefs(95)  =   "Named:id=41:RecordSelector"
      _StyleDefs(96)  =   ":id=41,.parent=34"
      _StyleDefs(97)  =   "Named:id=42:FilterBar"
      _StyleDefs(98)  =   ":id=42,.parent=33"
   End
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   2040
      TabIndex        =   4
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No STT"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Tanggal"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cara Bayar"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nama Bank"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "No Giro"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tanggal Lunas"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Nilai"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Customer"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Debet Acc"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Ket Acc"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "IdNo"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "updated"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=1769238625"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1323"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1244"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=163973724"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1614"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1535"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=164035200"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1693"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1614"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1826"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1746"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(4)._MinWidth=-1"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1746"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1667"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2434"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2355"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=3678"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=3598"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=1667"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=1588"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=4577"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=4498"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(45)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(48)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(49)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(11).Order=12"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=88,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=86,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=83,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=84,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=85,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(78)  =   "Named:id=33:Normal"
      _StyleDefs(79)  =   ":id=33,.parent=0"
      _StyleDefs(80)  =   "Named:id=34:Heading"
      _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=34,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=35:Footing"
      _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=36:Selected"
      _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=37:Caption"
      _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(89)  =   "Named:id=38:HighlightRow"
      _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=39:EvenRow"
      _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=40:OddRow"
      _StyleDefs(94)  =   ":id=40,.parent=33"
      _StyleDefs(95)  =   "Named:id=41:RecordSelector"
      _StyleDefs(96)  =   ":id=41,.parent=34"
      _StyleDefs(97)  =   "Named:id=42:FilterBar"
      _StyleDefs(98)  =   ":id=42,.parent=33"
   End
   Begin UsrText.IText fNoAccount 
      Height          =   270
      Left            =   8760
      TabIndex        =   14
      Top             =   120
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
      Caption         =   "No Account"
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cara Bayar"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label fKet2 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   10
      Top             =   4560
      Width           =   11655
   End
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   9
      Top             =   720
      Width           =   11655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PEMBELIAN"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "FormJurnalPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim y As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim col2 As TrueOleDBGrid80.Columns
Dim z As New XArrayDB
Dim col3 As TrueOleDBGrid80.Columns
Dim LAcc As Boolean

Private Sub fAkhir_Validate(Cancel As Boolean)
    DoQuery
End Sub
Private Sub DoQuery()
On Error Resume Next
    MyFilter = ""
    If fBelumSet Then MyFilter = " and isnull(DebetAcc)"
    If Trim(fCaraBayar) <> "" Then MyFilter = MyFilter & " and CaraBayar='" & fCaraBayar & "'"
    If Trim(fNoAccount) <> "" Then MyFilter = MyFilter & " and DebetAcc='" & fNoAccount & "'"
    a = "select NoSTT,Tanggal,CaraBayar,NamaBank,NoGiro,TanggalGiro,Nilai,NamaSupplier,DebetAcc,KetAcc,IdNo,0 from t_PembelianSTTPembayaran where TanggalGiro>=" & cD(fAwal) & " and TanggalGiro<=" & cD(fAkhir) & MyFilter & " order by TanggalGiro"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    a = "select NoSTT,Tanggal,CaraBayar,NamaBank,NoGiro,TanggalGiro,Nilai,KetCustomer,DebetAcc,KetAcc,IdNo,0 from t_STTPembayaran" & pTipe & " where TanggalGiro>=" & cD(fAwal) & " and TanggalGiro<=" & cD(fAkhir) & MyFilter & " order by TanggalGiro"
    query a
    y.ReDim 0, 0, 0, col2.Count - 1
    y.DeleteRows 0
    If RS.RecordCount > 0 Then y.LoadRows RS.GetRows
    TDBGrid2.Rebind
    If Not LAcc Then
        LAcc = True
        a = "select NoAccount, Deskripsi from m_ChartAccount where child=0 order by NoAccount"
        query a
        If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
        TDBDropDown1.Rebind
    End If
    TDBGrid1.Bookmark = 0
    TDBGrid2.Bookmark = 0
End Sub

Private Sub fBelumSet_Click()
    DoQuery
End Sub

Private Sub fCaraBayar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        HelpMe "Cara Bayar", Me
    ElseIf KeyCode = 13 Then
        DoQuery
    End If
End Sub

Sub FormHelpKeyDown(ByVal a As String)
    ActiveControl.Text = a
    DoQuery
End Sub

Private Sub fNoAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        HelpMe "No Account", Me
    ElseIf KeyCode = 13 Then
        DoQuery
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LAcc = False
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set TDBGrid1.Array = x
    Set TDBGrid2.Array = y
    Set TDBDropDown1.Array = z
    Set col1 = TDBGrid1.Columns
    Set col2 = TDBGrid2.Columns
    Set col3 = TDBDropDown1.Columns
    For i = 0 To col1.Count - 1
        col1(i).Locked = True
    Next
    For i = 0 To col2.Count - 1
        col2(i).Locked = True
    Next
    col1("Kredit Acc").DropDown = TDBDropDown1
    col2("Debet Acc").DropDown = TDBDropDown1
    col1("Tanggal").NumberFormat = "FormatText Event"
    col1("Tanggal Lunas").NumberFormat = "FormatText Event"
    col2("Tanggal").NumberFormat = "FormatText Event"
    col2("Tanggal Lunas").NumberFormat = "FormatText Event"
    TDBGridSetLock TDBGrid1, "Kredit Acc", False
    TDBGridSetLock TDBGrid2, "Debet Acc", False
    col1("Nilai").Tag = "Decimal"
    col2("Nilai").Tag = "Decimal"
    TDBGridSetVisible TDBGrid1, "updated@IdNo@No STT@Tanggal@Cara Bayar", False
    TDBGridSetVisible TDBGrid2, "updated@IdNo@No STT@Tanggal@Cara Bayar", False
    TDBGrid1.FetchRowStyle = True
    TDBGrid2.FetchRowStyle = True
    TDBGridLoad TDBGrid1
    TDBGridLoad TDBGrid2
    fAkhir = pServerDate
    DoQuery
End Sub

Private Sub fSave_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) = "1" Then
            a = "update t_PembelianSTTPembayaran set DebetAcc='" & x(i, col1("Kredit Acc").ColIndex) & "', KetAcc='" & x(i, col1("Ket Acc").ColIndex) & "' where IdNo=" & x(i, col1("IdNo").ColIndex)
            ExecMe a
        End If
    Next
    TDBGrid2.Update
    For i = 0 To y.UpperBound(1)
        If y(i, col2("updated").ColIndex) = "1" Then
            a = "update t_STTPembayaran" & pTipe & " set DebetAcc='" & y(i, col2("Debet Acc").ColIndex) & "', KetAcc='" & y(i, col1("Ket Acc").ColIndex) & "' where IdNo=" & y(i, col2("IdNo").ColIndex)
            ExecMe a
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoEvents
    DoQuery
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBDropDown1_DropDownClose()
On Error Resume Next
    If ActiveControl.Name = "TDBGrid1" Then
        col1("Ket Acc").Value = ""
        If Not IsNull(TDBDropDown1.Bookmark) Then col1("Ket Acc").Value = col3("Ket Acc").Value
    Else
        col2("Ket Acc").Value = ""
        If Not IsNull(TDBDropDown1.Bookmark) Then col2("Ket Acc").Value = col3("Ket Acc").Value
    End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    col1("updated").Value = 1
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("updated").ColIndex) = "1" Then
        RowStyle.BackColor = vbYellow
    End If
End Sub
Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If y(Bookmark, col2("updated").ColIndex) = "1" Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
On Error Resume Next
    a = col1(ColIndex).Caption
    If a = "Tanggal" Or a = "Tanggal Lunas" Then
        Value = cTanggal(Value)
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    fKet = "No STT: " & col1("No STT").Text & "  Tanggal: " & col1("Tanggal").Text & "  Cara Bayar: " & col1("Cara Bayar").Text & "  [" & col1("Supplier").Text & "]"
End Sub

Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    col2("updated").Value = 1
    TDBDropDown1_DropDownClose
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    fKet2 = "No STT: " & col2("No STT").Text & "  Cara Bayar: " & col2("Cara Bayar").Text & "  [" & col2("Customer").Text & "]"
End Sub

Private Sub TDBGrid2_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
On Error Resume Next
    a = col2(ColIndex).Caption
    If a = "Tanggal" Or a = "Tanggal Lunas" Then
        Value = cTanggal(Value)
    End If
End Sub




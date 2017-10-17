VERSION 5.00
Object = "{E2D3646A-2684-4DDE-BE47-3323E01328EE}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormHargaBeli 
   BackColor       =   &H00FFC0C0&   Caption         =   "HARGA BELI"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Tag             =   "13"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin UsrText.IText fSupplier 
      Height          =   270
      Left            =   1320
      TabIndex        =   6
      Top             =   600
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
   Begin VB.CommandButton fPrintTT 
      Caption         =   "PRINT TT"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7646
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
      Columns(2).Caption=   "No BTB"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tanggal BTB"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Supplier"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Total"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Tempo"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1032"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=953"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1508"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1429"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1852"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1773"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=7541"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=7461"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1958"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1879"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=65,.parent=2,.namedParent=67"
      _StyleDefs(17)  =   "FilterBarStyle:id=68,.parent=1,.namedParent=70"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=66,.parent=65"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=69,.parent=68"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=40,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=37,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=38,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=39,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=24,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=86,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=83,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=84,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=85,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=78,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=60,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=57,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=59,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=82,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=15"
      _StyleDefs(58)  =   "Named:id=29:Normal"
      _StyleDefs(59)  =   ":id=29,.parent=0"
      _StyleDefs(60)  =   "Named:id=30:Heading"
      _StyleDefs(61)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=30,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=31:Footing"
      _StyleDefs(64)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=32:Selected"
      _StyleDefs(66)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=33:Caption"
      _StyleDefs(68)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(69)  =   "Named:id=34:HighlightRow"
      _StyleDefs(70)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(71)  =   "Named:id=35:EvenRow"
      _StyleDefs(72)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=36:OddRow"
      _StyleDefs(74)  =   ":id=36,.parent=29"
      _StyleDefs(75)  =   "Named:id=67:RecordSelector"
      _StyleDefs(76)  =   ":id=67,.parent=30"
      _StyleDefs(77)  =   "Named:id=70:FilterBar"
      _StyleDefs(78)  =   ":id=70,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   9615
      _ExtentX        =   16960
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
      Columns(1).Caption=   "No PO"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nama Barang"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "QTY"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Satuan"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Curr"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Harga"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "PPN"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Total"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "IdStock"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1455"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1376"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=4392"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4313"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1482"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1402"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1191"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1111"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1402"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1323"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1614"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1535"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=820"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=741"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=1958"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1879"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=65,.parent=2,.namedParent=67"
      _StyleDefs(17)  =   "FilterBarStyle:id=68,.parent=1,.namedParent=70"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=66,.parent=65"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=69,.parent=68"
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
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=44,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=41,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=42,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=43,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=48,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=45,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=47,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=64,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=61,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=62,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=63,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=52,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=49,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=51,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=56,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=53,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=54,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=55,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=60,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=57,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=58,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=59,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=15"
      _StyleDefs(70)  =   "Named:id=29:Normal"
      _StyleDefs(71)  =   ":id=29,.parent=0"
      _StyleDefs(72)  =   "Named:id=30:Heading"
      _StyleDefs(73)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=30,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=31:Footing"
      _StyleDefs(76)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=32:Selected"
      _StyleDefs(78)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=33:Caption"
      _StyleDefs(80)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(81)  =   "Named:id=34:HighlightRow"
      _StyleDefs(82)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(83)  =   "Named:id=35:EvenRow"
      _StyleDefs(84)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(85)  =   "Named:id=36:OddRow"
      _StyleDefs(86)  =   ":id=36,.parent=29"
      _StyleDefs(87)  =   "Named:id=67:RecordSelector"
      _StyleDefs(88)  =   ":id=67,.parent=30"
      _StyleDefs(89)  =   "Named:id=70:FilterBar"
      _StyleDefs(90)  =   ":id=70,.parent=29"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PEMBELIAN BARANG BELUM LUNAS"
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
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label2"
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
      TabIndex        =   1
      Top             =   5400
      Width           =   9615
   End
End
Attribute VB_Name = "FormHargaBeli"
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

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set TDBGrid1.Array = x
    Set TDBGrid2.Array = y
    Set col1 = TDBGrid1.Columns
    Set col2 = TDBGrid2.Columns
    col1("Total").Alignment = dbgRight
    col1("Total").NumberFormat = "Standard"
    col1("Tanggal BTB").NumberFormat = "FormatText Event"
    col1(0).Alignment = dbgCenter
    col1(1).Alignment = dbgCenter
    col1(0).ValueItems.Presentation = dbgCheckBox
    col1(1).ValueItems.Presentation = dbgCheckBox
    TDBGrid2.FetchRowStyle = True
    col2("updated").Visible = False
    col2("QTY").Alignment = dbgRight
    col2("Harga").Alignment = dbgRight
    col2("Total").Alignment = dbgRight
    col2("Harga").NumberFormat = "Standard"
    col2("Total").NumberFormat = "Standard"
    col2("PPN").ValueItems.Presentation = dbgCheckBox
    For i = 0 To col2.Count - 1
        col2(i).Locked = True
    Next
    For i = 1 To col1.Count - 1
        col1(i).Locked = True
    Next
    col2("Harga").Locked = False
    col2("PPN").Locked = False
    DoQuery
End Sub
Sub DoQuery()
    a = "select PrintedTT-1 , PrintedTT*-1, NoBTB, TanggalBTB, NamaSupplier, Total, WaktuPembayaran from t_BTB where Pelunasan=0 " & MyFilter
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    GetDetail
End Sub

Function MyFilter()
    If Trim(fSupplier) <> "" Then MyFilter = " and NamaSupplier='" & fSupplier & "'"
End Function

Private Sub GetDetail()
On Error Resume Next
    a = "SELECT 0, t_BTBDetail.NoPODetail, m_StockBeli.NamaBarang, t_BTBDetail.QTY, m_StockBeli.Satuan, t_BTBDetail.MataUangDetail, t_BTBDetail.Harga, t_BTBDetail.PPNYesNo, t_BTBDetail.TotalHarga, t_BTBDetail.IdStock FROM t_BTBDetail left JOIN m_StockBeli ON t_BTBDetail.IdStock = m_StockBeli.IdStock where NoBTB='" & esc(col1("No BTB").Value) & "'"
    query a
    y.ReDim 0, 0, 0, col2.Count - 1
    y.DeleteRows 0
    If RS.RecordCount > 0 Then y.ReDim 0, 0, 0, col2.Count - 1
    y.DeleteRows 0
    query a
    y.LoadRows RS.GetRows
    TDBGrid2.Rebind
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid2.Width = ScaleWidth - 2 * TDBGrid2.Left
    TDBGrid1.Width = TDBGrid2.Width
    TDBGrid2.Height = ScaleHeight - TDBGrid2.Top - 100
End Sub

Private Sub fPrintTT_Click()
    TDBGrid1.Update
    c = ""
    For i = 0 To x.UpperBound(1)
        c = c & "','" & x(i, 2)
    Next
    a = "update t_BTB set PrintedTT=1 where NoBTB in('" & Mid(c, 4) & "')"
    ExecMe a
    'FormReport.LoadMe
    DoQuery
    
End Sub

Private Sub fSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DoQuery
    ElseIf KeyCode = 112 Then
        HelpMe "Nama Supplier", Me
    End If
End Sub

Sub FormHelpKeyDown(ByVal a As String)
    ActiveControl.Text = a
    DoQuery
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    If Not cekValid("EDIT", Tag) Then Exit Sub
    BeginTransaction
    TDBGrid2.Update
    For i = 0 To y.UpperBound(1)
        If y(i, col2("updated").ColIndex) <> "0" Then
            a = "update t_BTBDetail set Harga=" & cNum(y(i, col2("Harga").ColIndex)) & _
                ", TotalHarga=" & cNum(y(i, col2("Total").ColIndex)) & _
                ", PPNYesNo=" & IIf(y(i, col2("PPN").ColIndex) = 0, 0, 1) & " where NoBTB='" & esc(col1("No BTB").Value) & "' and IdStock=" & y(i, col2("IdStock").ColIndex)
            ExecMe a
        End If
    Next
    b = 0
    For i = 0 To y.UpperBound(1)
        b = b + y(i, col2("Total").ColIndex)
    Next
    col1("Total").Value = b
    a = "update t_BTB set Total=" & cNum(b) & " where NoBTB='" & esc(col1("No BTB").Value) & "'"
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    DoQuery
    Exit Sub
err:
    RollBackTransaction
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    a = col1(ColIndex).Caption
    If a = "Tanggal BTB" Then
        Value = cTanggal(Value)
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDetail
End Sub

Private Sub TDBGrid2_AfterColEdit(ByVal ColIndex As Integer)
    col2("updated").Value = "1"
    a = col2(ColIndex).Caption
    If a = "Harga" Or a = "PPN" Then
        col2("Total").Value = col2("QTY").Value * col2("Harga").Value * (1 + 10 * IIf(col2("PPN").Value = 0, 0, 1) / 100)
    End If
End Sub

Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If y(Bookmark, col2("updated").ColIndex) = "1" Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Private Sub TDBGrid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    fKet = col2("Nama Barang").Value
End Sub



VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Begin VB.Form FormRubahStockRetur 
   BackColor       =   &H00FFC0C0&
   Caption         =   "STOCK RETUR"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Tag             =   "49"
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2415
      Left            =   2400
      TabIndex        =   13
      Top             =   4200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode Barang"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Warna"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tube"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Sat Besar"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Sat Kecil"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3863"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3784"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=145"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1323"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1244"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=74843660"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1191"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1111"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=64"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1588"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1508"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1693"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton fQueryStock 
      Caption         =   "QUERY STOCK"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin UsrText.IText fCustomer 
      Height          =   270
      Left            =   5640
      TabIndex        =   10
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
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
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   3240
      TabIndex        =   8
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
      Left            =   2280
      TabIndex        =   7
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
   Begin VB.CheckBox fSemua 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SEMUA"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Height          =   2355
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   4154
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No Nota"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "IdStock"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tanggal"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Jenis"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Kode"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Warna"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "No Warna"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Tube"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Grade"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "SatBesar"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "SatKecil"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Jumlah1"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Jumlah2"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "IdDet"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "updated"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "cIdStock"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "cn1"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "cn2"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "ctanggal"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   19
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=19"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1111"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1032"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1482"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1402"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(2)._MinWidth=-2147483633"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=873"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=794"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(3)._MinWidth=-2147483633"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=4736"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=4657"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(4)._MinWidth=4"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=1005"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=926"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=1508"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=1429"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=1191"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=1111"
      Splits(0)._ColumnProps(35)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(36)=   "Column(8).Width=979"
      Splits(0)._ColumnProps(37)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(8)._WidthInPix=900"
      Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(40)=   "Column(9).Width=1376"
      Splits(0)._ColumnProps(41)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(9)._WidthInPix=1296"
      Splits(0)._ColumnProps(43)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(44)=   "Column(10).Width=1217"
      Splits(0)._ColumnProps(45)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(10)._WidthInPix=1138"
      Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(48)=   "Column(11).Width=1217"
      Splits(0)._ColumnProps(49)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(11)._WidthInPix=1138"
      Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(52)=   "Column(12).Width=1720"
      Splits(0)._ColumnProps(53)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(12)._WidthInPix=1640"
      Splits(0)._ColumnProps(55)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(56)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(57)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(59)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(60)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(61)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(63)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(64)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(65)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(67)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(68)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(69)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(71)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(72)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(73)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(75)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(76)=   "Column(17)._MinWidth=74178018"
      Splits(0)._ColumnProps(77)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(78)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(80)=   "Column(18).Order=19"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=106,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=70,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=74,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=82,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=90,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=94,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(16).Style:id=98,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(17).Style:id=102,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(17).HeadingStyle:id=99,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(17).FooterStyle:id=100,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(17).EditorStyle:id=101,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
      _StyleDefs(106) =   "Named:id=33:Normal"
      _StyleDefs(107) =   ":id=33,.parent=0"
      _StyleDefs(108) =   "Named:id=34:Heading"
      _StyleDefs(109) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(110) =   ":id=34,.wraptext=-1"
      _StyleDefs(111) =   "Named:id=35:Footing"
      _StyleDefs(112) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(113) =   "Named:id=36:Selected"
      _StyleDefs(114) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(115) =   "Named:id=37:Caption"
      _StyleDefs(116) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(117) =   "Named:id=38:HighlightRow"
      _StyleDefs(118) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(119) =   "Named:id=39:EvenRow"
      _StyleDefs(120) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(121) =   "Named:id=40:OddRow"
      _StyleDefs(122) =   ":id=40,.parent=33"
      _StyleDefs(123) =   "Named:id=41:RecordSelector"
      _StyleDefs(124) =   ":id=41,.parent=34"
      _StyleDefs(125) =   "Named:id=42:FilterBar"
      _StyleDefs(126) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   3731
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
      Columns(1).Caption=   "Tanggal NR"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nama Barang"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "SatBesar"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "SatKecil"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Jumlah1"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Jumlah2"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Ket"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "IdStock"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "IdDet"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4789"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4710"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1693"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(1)._MinWidth=164168260"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=5424"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=5345"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=1323"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1244"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=1244"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=1164"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1244"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1164"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=1429"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=1349"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(9)._MinWidth=165487776"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=212,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
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
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   10815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal NR"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Perubahan Stock"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Retur"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   1155
   End
End
Attribute VB_Name = "FormRubahStockRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid80.Columns
Dim col2 As TrueOleDBGrid80.Columns
Dim col3 As TrueOleDBGrid80.Columns
Dim x As New XArrayDB
Dim y As New XArrayDB
Dim z As New XArrayDB
Dim iNamaCustomer As Integer
Dim iTanggalNR As Integer
Dim iNamaBarang As Integer
Dim iSatBesar As Integer
Dim iSatKecil As Integer
Dim iJumlah1 As Integer
Dim iJumlah2 As Integer
Dim iKet As Integer
Dim iIdStock As Integer
Dim iIdDet As Integer

Dim jNoNota As Integer
Dim jIdStock As Integer
Dim jTanggal As Integer
Dim jJenis As Integer
Dim jKode As Integer
Dim jWarna As Integer
Dim jNoWarna As Integer
Dim jTube As Integer
Dim jGrade As Integer
Dim jSatBesar As Integer
Dim jSatKecil As Integer
Dim jJumlah1 As Integer
Dim jJumlah2 As Integer
Dim jIdDet As Integer
Dim jUpdated As Integer
Dim jcIdStock As Integer
Dim jcn1 As Integer
Dim jcn2 As Integer
Dim jcTanggal As Integer

Private Sub fAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub fCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub fDelete_Click()
On Error GoTo err
    BeginTransaction
    b = MsgBox("Yakin mau Hapus?", vbYesNo)
    If b = vbNo Then GoTo err
    a = "delete from t_InputStock~ where IdDetStockRetur=" & col2(jIdDet).Value
    If ExecMe(a) = 0 Then GoTo err
    a = "update m_Stock~ set JumlahBox=JumlahBox-" & cNum(col2(jcn1).Value) & ",JumlahKG=JumlahKG-" & cNum(col2(jcn2).Value) & " where IdStock=" & col2(jcIdStock).Value
    If ExecMe(a) = 0 Then GoTo err
    a = "delete from t_StockRetur~ where IdDet=" & col2(jIdDet).Value
    If ExecMe(a) = 0 Then GoTo err
    If y.UpperBound(1) = 0 Then
        a = "update t_NRDetail~ set StatusRubahStock=0 where IdDet=" & col1(iIdDet).Value
        If ExecMe(a) = 0 Then GoTo err
    End If
    CommitTransaction
    MsgBox "SUKSES"
    TDBGrid2.Delete
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub Form_Load()
    iNamaCustomer = 0
    iTanggalNR = 1
    iNamaBarang = 2
    iSatBesar = 3
    iSatKecil = 4
    iJumlah1 = 5
    iJumlah2 = 6
    iKet = 7
    iIdStock = 8
    iIdDet = 9
    
    jNoNota = 0
    jIdStock = 1
    jTanggal = 2
    jJenis = 3
    jKode = 4
    jWarna = 5
    jNoWarna = 6
    jTube = 7
    jGrade = 8
    jSatBesar = 9
    jSatKecil = 10
    jJumlah1 = 11
    jJumlah2 = 12
    jIdDet = 13
    jUpdated = 14
    jcIdStock = 15
    jcn1 = 16
    jcn2 = 17
    jcTanggal = 18
    Set col1 = TDBGrid1.Columns
    Set col2 = TDBGrid2.Columns
    Set col3 = TDBDropDown1.Columns
    Set TDBGrid1.Array = x
    Set TDBGrid2.Array = y
    Set TDBDropDown1.Array = z
    TDBGrid1.AllowUpdate = False
    TDBGrid2.AllowAddNew = True
    col1(iTanggalNR).Tag = "Date"
    col1(iJumlah1).Tag = "Integer"
    col1(iJumlah2).Tag = "Decimal"
    TDBGridLoad TDBGrid1
    TDBGridSetVisible TDBGrid1, iIdStock & "@" & iIdDet
    col2(jTanggal).Tag = "Date"
    col2(jJumlah1).Tag = "Integer"
    col2(jJumlah2).Tag = "Decimal"
    col2(jNoNota).Tag = "Integer"
    TDBGridLoad TDBGrid2
    TDBGridSetVisible TDBGrid2, jcTanggal & "@" & jIdStock & "@" & jIdDet & "@" & jcIdStock & "@" & jcn1 & "@" & jcn2 & "@" & jUpdated
    TDBGridSetLock TDBGrid2, jWarna & "@" & jTube & "@" & jSatBesar & "@" & jSatKecil, True
    TDBGrid2.FetchRowStyle = True
    col2(jKode).DropDown = TDBDropDown1
    col2(jKode).AutoDropDown = True
    DoQuery
    a = "select distinct KodeBarang, Warna, Tube, SatBesar, SatKecil from m_Stock~ where IsActive=1 order by KodeBarang, Warna, Tube, SatBesar, SatKecil"
    query a
    z.ReDim 0, 0, 0, col3.Count - 1
    z.DeleteRows 0
    If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
    TDBDropDown1.Rebind
End Sub

Private Function MyFilter() As String
    MyFilter = " and ReturBox>0"
    If fSemua.Value = 0 Then MyFilter = MyFilter & " and StatusRubahStock=0"
    If cD(fAwal) <> "A" Then MyFilter = MyFilter & " and TanggalNRDetail>=" & cD(fAwal)
    If cD(fAkhir) <> "A" Then MyFilter = MyFilter & " and TanggalNRDetail<=" & cD(fAkhir)
    If Trim(fCustomer) <> "" Then MyFilter = MyFilter & " and m_Customer.Nama like '%" & fCustomer & "%'"
End Function

Sub DoQuery()
    a = "select m_Customer.Nama, TanggalNRDetail, Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna+' '+Tube+' GRADE '+Grade as NamaBarang, SatBesar, SatKecil, ReturBox, ReturKG, KetDetail, m_Stock~.IdStock, t_NRDetail~.IdDet from (t_NRDetail~ left join m_Stock~ on t_NRDetail~.IdStock=m_Stock~.IdStock) left join m_Customer on m_Customer.Kode=t_NRDetail~.KodeCustomerDetail where 1=1 " & MyFilter & " order by Jenis, KodeBarang, Warna, NoWarna, Tube, Grade"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    For i = 0 To x.UpperBound(1)
        x(i, iTanggalNR) = cTanggal(x(i, iTanggalNR))
    Next
    TDBGrid1.Rebind
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left - 100
    TDBGrid2.Width = TDBGrid1.Width
    TDBGrid2.Height = ScaleHeight - TDBGrid2.Top - 100
End Sub

Private Sub fQueryStock_Click()
On Error Resume Next
    FormStock.LoadMe Me, col2(jJenis).Value, col2(jKode).Value, col2(jNoWarna).Value, col2(jTube).Value, col2(jGrade).Value, col2(jSatBesar)
End Sub

Sub SetOtherRowData(ByVal tNo As Long)
On Error GoTo err
    a = "select Jenis,KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil from m_stock~ where IsActive=1 and IdStock=" & tNo
    query a
    col2(jJenis).Value = RS.Fields("Jenis").Value
    col2(jKode).Value = RS.Fields("KodeBarang").Value
    col2(jWarna).Value = RS.Fields("Warna").Value
    col2(jNoWarna).Value = RS.Fields("NoWarna").Value
    col2(jTube).Value = RS.Fields("Tube").Value
    col2(jGrade).Value = RS.Fields("Grade").Value
    col2(jSatBesar).Value = RS.Fields("SatBesar").Value
    col2(jSatKecil).Value = RS.Fields("SatKecil").Value
    TDBGrid2.SetFocus
err:
End Sub

Private Sub fSemua_Click()
    DoQuery
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    TDBGrid2.Update
    BeginTransaction
    a = "update t_NRDetail~ set StatusRubahStock=1 where IdDet=" & col1(iIdDet).Value
    If ExecMe(a) = 0 Then GoTo err
    For i = 0 To y.UpperBound(1)
        If y(i, jUpdated) = "1" Then
            a = "select top 1 IdStock from m_Stock~ where Jenis='" & y(i, jJenis) & "' and KodeBarang='" & y(i, jKode) & "' and Warna='" & y(i, jWarna) & "' and NoWarna='" & y(i, jNoWarna) & "' and Tube='" & y(i, jTube) & "' and Grade='" & y(i, jGrade) & "' and SatBesar='" & y(i, jSatBesar) & "'"
            query a
            If RS.RecordCount = 0 Then GoTo err
            Dim kIdStock As Long
            kIdStock = RS.Fields(0).Value
            If y(i, jIdDet) = "" Then
                a = "insert into t_StockRetur~(IdStock, NoNota, Tanggal, n1, n2, IdDetNR) values(" & _
                    kIdStock & _
                    "," & cNum(y(i, jNoNota)) & _
                    "," & cD(y(i, jTanggal)) & _
                    "," & cNum(y(i, jJumlah1)) & _
                    "," & cNum(y(i, jJumlah2)) & _
                    "," & col1(iIdDet).Value & ")"
                If ExecMe(a) = 0 Then GoTo err
                a = "update m_Stock~ set JumlahBox=JumlahBox+" & cNum(y(i, jJumlah1)) & ",JumlahKG=JumlahKG+" & cNum(y(i, jJumlah2)) & " where IdStock=" & kIdStock
                If ExecMe(a) = 0 Then GoTo err
            Else
                a = "update t_StockRetur~ set " & _
                    "IdStock=" & kIdStock & _
                    ", NoNota=" & cNum(y(i, jNoNota)) & _
                    ", Tanggal=" & cD(y(i, jTanggal)) & _
                    ", n1=" & cNum(y(i, jJumlah1)) & _
                    ", n2=" & cNum(y(i, jJumlah2)) & " where IdDet=" & y(i, jIdDet)
                If ExecMe(a) = 0 Then GoTo err
                a = "update m_Stock~ set JumlahBox=JumlahBox-" & cNum(y(i, jcn1)) & ",JumlahKG=JumlahKG-" & cNum(y(i, jcn2)) & " where IdStock=" & y(i, jcIdStock)
                If ExecMe(a) = 0 Then GoTo err
                a = "update m_Stock~ set JumlahBox=JumlahBox+" & cNum(y(i, jJumlah1)) & ",JumlahKG=JumlahKG+" & cNum(y(i, jJumlah2)) & " where IdStock=" & kIdStock
                If ExecMe(a) = 0 Then GoTo err
            End If
            PrintedCode = Round(Rnd * 2 ^ 21)
            a = "select max(iddet) from t_StockRetur~ where IdDetNR=" & col1(iIdDet).Value
            query a
            IdDetStockRetur = RS.Fields(0).Value
            a = "update t_InputStock~ set Status=1, NoBukti=" & y(i, jNoNota) & ", PrintedCode=" & PrintedCode & ", IdStock=" & y(i, jIdStock) & ", n1=" & cNum(y(i, jJumlah1)) & ", n2=" & cNum(y(i, jJumlah2)) & " where IdDetStockRetur=" & y(i, jIdDet)
            If ExecMe(a) = 0 Then
                a = "insert into t_InputStock~(NoBukti, Tanggal, IdStock, Keterangan, n1, n2, PrintedCode, IdDetStockRetur,Status) values(" & _
                    y(i, jNoNota) & _
                    "," & cD(y(i, jTanggal)) & _
                    "," & kIdStock & _
                    ",'RETUR DARI " & col1(iNamaCustomer) & _
                    "'," & cNum(y(i, jJumlah1)) & _
                    "," & cNum(y(i, jJumlah2)) & _
                    "," & PrintedCode & "," & IdDetStockRetur & ",1)"
                If ExecMe(a) = 0 Then GoTo err
            End If
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    TDBGrid1_RowColChange -1, -1
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBDropDown1_DropDownClose()
On Error Resume Next
    col2(jKode).Value = col3("Kode Barang").Value
    col2(jWarna).Value = col3("Warna").Value
    col2(jTube).Value = col3("Tube").Value
    col2(jSatBesar).Value = col3("Sat Besar").Value
    col2(jSatKecil).Value = col3("Sat Kecil").Value
End Sub

Private Sub TDBDropDown1_Paint()
    TDBGrid2.SelLength = Len(TDBGrid2.Text)
End Sub

Private Sub TDBGrid1_DblClick()
    If y.UpperBound(1) = -1 Then
        y.AppendRows
        TDBGrid2.Rebind
    End If
    SetOtherRowData col1(iIdStock).Value
    col2("Tanggal").Value = col1("Tanggal NR").Value
    col2("Jumlah1").Value = col1("Jumlah1").Value
    col2("Jumlah2").Value = col1("Jumlah2").Value
    col2("updated").Value = 1
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    TDBGridKeyDown TDBGrid2, KeyCode
End Sub

Private Sub TDBGrid2_KeyPress(KeyAscii As Integer)
On Error Resume Next
    TDBGridKeyPress TDBGrid2, KeyAscii
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    fKet = col1(iNamaBarang).Value & " (" & col1(iNamaCustomer).Value & ")"
    If IsNull(col1(iIdDet).Value) Then Exit Sub
    a = "select NoNota, m_Stock~.IdStock, Tanggal, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, n1, n2, IdDet, 0, m_stock~.IdStock, n1, n2, Tanggal from t_StockRetur~ left join m_Stock~ on t_StockRetur~.IdStock=m_Stock~.IdStock where IdDetNR=" & col1(iIdDet).Value
    query a
    y.ReDim 0, 0, 0, col2.Count - 1
    y.DeleteRows 0
    If RS.RecordCount > 0 Then y.LoadRows RS.GetRows
    For i = 0 To y.UpperBound(1)
        y(i, jTanggal) = cTanggal(y(i, jTanggal))
    Next
    TDBGrid2.Rebind
End Sub

Private Sub TDBGrid2_AfterColEdit(ByVal ColIndex As Integer)
    col2(jUpdated).Value = "1"
End Sub

Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If y(Bookmark, jUpdated) = "1" Then RowStyle.BackColor = vbYellow
End Sub


VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMutasi 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MUTASI"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "FormMutasi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11760
   Tag             =   "25"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPost 
      Caption         =   "POST"
      Height          =   315
      Left            =   2160
      TabIndex        =   27
      Top             =   120
      Width           =   615
   End
   Begin UsrText.IText Text1 
      Height          =   330
      Index           =   4
      Left            =   7080
      TabIndex        =   5
      Tag             =   "Grade"
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText Text1 
      Height          =   330
      Index           =   2
      Left            =   5040
      TabIndex        =   3
      Tag             =   "NoWarna"
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText Text1 
      Height          =   330
      Index           =   3
      Left            =   6120
      TabIndex        =   4
      Tag             =   "Tube"
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText Text1 
      Height          =   330
      Index           =   1
      Left            =   4080
      TabIndex        =   12
      Tag             =   "Warna"
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText Text1 
      Height          =   330
      Index           =   5
      Left            =   8040
      TabIndex        =   13
      Tag             =   "IdStock"
      Top             =   720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText Text1 
      Height          =   330
      Index           =   6
      Left            =   2880
      TabIndex        =   17
      Tag             =   "SatBesar"
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox dLGrade 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Grade"
      Height          =   255
      Left            =   7080
      TabIndex        =   23
      Top             =   480
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox dLTube 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tube"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   480
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox dLNoWarna 
      BackColor       =   &H00FFC0C0&
      Caption         =   "No Warna"
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   480
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin UsrText.IText fKode 
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Tag             =   "KodeBarang"
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox fTidakAda 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tidak di"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox dLKode 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kode"
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   480
      Value           =   1  'Checked
      Width           =   855
   End
   Begin UsrText.IText Text1 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "Jenis"
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox dLJenis 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Jenis"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CommandButton fCopy 
      Caption         =   "&COPY"
      Height          =   315
      Left            =   4800
      TabIndex        =   18
      Top             =   60
      Width           =   975
   End
   Begin VB.CheckBox fDetail 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detail"
      Height          =   345
      Left            =   8880
      TabIndex        =   15
      Top             =   600
      Width           =   885
   End
   Begin VB.CommandButton fExcel 
      Caption         =   "&PRINT MUTASI"
      Height          =   315
      Left            =   3120
      TabIndex        =   14
      Top             =   60
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ada Mutasi Saja"
      Height          =   255
      Left            =   9060
      TabIndex        =   10
      Top             =   60
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Yang Ada Stock Saja"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   60
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   6015
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IdStock"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama Barang"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tanggal"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Masuk"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "MasukKG"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Keluar"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "KeluarKG"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Akhir"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "AkhirKG"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "NoBukti"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Keterangan"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "SatBesar"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "SatKecil"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Jenis"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Kode"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "AwalKgs"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "InKgs"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "OutKgs"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "AkhirKgs"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "AwalNoUrut"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "InNoUrut"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "OutNoUrut"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "AkhirNoUrut"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   23
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=23"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1058"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=979"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=532"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5927"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5847"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=532"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1535"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1455"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=532"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=926"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=847"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=532"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1588"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1508"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=532"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=926"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=847"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=532"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1693"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=532"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=847"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=767"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=532"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=1667"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1588"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=532"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2752"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2672"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=532"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=4577"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=4498"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=532"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=1720"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=1640"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=532"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=1561"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=1482"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=532"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=532"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(13)._MinWidth=66632096"
      Splits(0)._ColumnProps(72)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(73)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(75)=   "Column(14)._ColStyle=532"
      Splits(0)._ColumnProps(76)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(77)=   "Column(14)._MinWidth=66682848"
      Splits(0)._ColumnProps(78)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(79)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(81)=   "Column(15)._ColStyle=532"
      Splits(0)._ColumnProps(82)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(83)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(16)._ColStyle=532"
      Splits(0)._ColumnProps(87)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(88)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(89)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(91)=   "Column(17)._ColStyle=532"
      Splits(0)._ColumnProps(92)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(93)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(94)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(96)=   "Column(18)._ColStyle=532"
      Splits(0)._ColumnProps(97)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(98)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(99)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(101)=   "Column(19)._ColStyle=532"
      Splits(0)._ColumnProps(102)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(103)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(104)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(105)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(106)=   "Column(20)._ColStyle=532"
      Splits(0)._ColumnProps(107)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(108)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(109)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(111)=   "Column(21)._ColStyle=532"
      Splits(0)._ColumnProps(112)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(113)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(22)._ColStyle=532"
      Splits(0)._ColumnProps(117)=   "Column(22).Order=23"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(18)  =   "RecordSelectorStyle:id=133,.parent=2,.namedParent=135"
      _StyleDefs(19)  =   "FilterBarStyle:id=136,.parent=1,.namedParent=138"
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
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=134,.parent=133"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=137,.parent=136"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=40,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=37,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=38,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=39,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=80,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=77,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=78,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=79,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=92,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=89,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=90,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=91,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=60,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=57,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=58,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=59,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=116,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=113,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=114,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=115,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=64,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=61,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=62,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=63,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=120,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=117,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=118,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=119,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=72,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=69,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=70,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=71,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=124,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=121,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=122,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=123,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=108,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=105,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=106,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=107,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=104,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=101,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=102,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=103,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=128,.parent=11"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=125,.parent=12"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=126,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=127,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=132,.parent=11"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=129,.parent=12"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=130,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=131,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=28,.parent=11"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=25,.parent=12"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=26,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=27,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=44,.parent=11"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=41,.parent=12"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=42,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=43,.parent=15"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=24,.parent=11"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=21,.parent=12"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=22,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=23,.parent=15"
      _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=48,.parent=11"
      _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=45,.parent=12"
      _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=46,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=47,.parent=15"
      _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=52,.parent=11"
      _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=49,.parent=12"
      _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=50,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=51,.parent=15"
      _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=56,.parent=11"
      _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=53,.parent=12"
      _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=54,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=55,.parent=15"
      _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=68,.parent=11"
      _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=65,.parent=12"
      _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=66,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=67,.parent=15"
      _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=76,.parent=11"
      _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=73,.parent=12"
      _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=74,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=75,.parent=15"
      _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=84,.parent=11"
      _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=81,.parent=12"
      _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=82,.parent=13"
      _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=83,.parent=15"
      _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=88,.parent=11"
      _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=85,.parent=12"
      _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=86,.parent=13"
      _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=87,.parent=15"
      _StyleDefs(126) =   "Named:id=29:Normal"
      _StyleDefs(127) =   ":id=29,.parent=0,.valignment=2"
      _StyleDefs(128) =   "Named:id=30:Heading"
      _StyleDefs(129) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(130) =   ":id=30,.wraptext=-1"
      _StyleDefs(131) =   "Named:id=31:Footing"
      _StyleDefs(132) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(133) =   "Named:id=32:Selected"
      _StyleDefs(134) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(135) =   "Named:id=33:Caption"
      _StyleDefs(136) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(137) =   "Named:id=34:HighlightRow"
      _StyleDefs(138) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(139) =   "Named:id=35:EvenRow"
      _StyleDefs(140) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(141) =   "Named:id=36:OddRow"
      _StyleDefs(142) =   ":id=36,.parent=29"
      _StyleDefs(143) =   "Named:id=135:RecordSelector"
      _StyleDefs(144) =   ":id=135,.parent=30"
      _StyleDefs(145) =   "Named:id=138:FilterBar"
      _StyleDefs(146) =   ":id=138,.parent=29"
   End
   Begin UsrText.IText fAwal 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   120
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Warna"
      Height          =   255
      Left            =   4080
      TabIndex        =   26
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sat Besar"
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Stock"
      Height          =   255
      Left            =   8040
      TabIndex        =   24
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "atau"
      Height          =   255
      Left            =   8580
      TabIndex        =   11
      Top             =   60
      Width           =   435
   End
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   10695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
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
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FormMutasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB

Dim rsOriginal() As Variant
Dim rsOriginalCount As Long
Dim seb As String
Dim col1 As TrueOleDBGrid80.Columns
Dim iIdStock As Integer
Dim iNamaBarang As Integer
Dim iTanggal As Integer
Dim iMasuk As Integer
Dim iMasukKG As Integer
Dim iKeluar As Integer
Dim iKeluarKG As Integer
Dim iAkhir As Integer
Dim iAkhirKG As Integer
Dim iNoBukti As Integer
Dim iKet As Integer
Dim iSatBesar As Integer
Dim iSatKecil As Integer
Dim iJenis As Integer
Dim iKode As Integer
Dim iAwalKgs As String
Dim iInKgs As String
Dim iOutKgs As String
Dim iAkhirKgs As String

Private Sub post(ByVal postDate As Long)
    Dim s As String
    Dim lastDate As Long
    Dim seb As String
    t = GetTickCount
    query "select max(tgl) from mutasi"
    If IsNull(RS.Fields(0).value) Then
        lastDate = 10000000
    Else
        lastDate = RS.Fields(0).value
    End If
    If lastDate + 1 > postDate Then
        s = "delete from mutasi"
        ExecMe s
        s = "delete from mutasi_hist where tgl>=" & postDate
        ExecMe s
        query "select max(tgl) from mutasi_hist"
        If IsNull(RS.Fields(0).value) Then
            lastDate = 151231
        Else
            lastDate = RS.Fields(0).value
        End If
        If lastDate > 0 Then
            s = "insert into mutasi select * from mutasi_hist where tgl=" & lastDate
            ExecMe (s)
            s = "delete from mutasi_hist where tgl=" & lastDate
            ExecMe s
        End If
    End If
    'lastDate = Add_Tanggal2(lastDate, 1)
    'Debug.Print GetTickCount - t
    t = GetTickCount
    
    InputStock postDate, postDate
    OutputStock postDate, postDate
    'Debug.Print GetTickCount - t
    t = GetTickCount
    
    's = "select Tanggal as tgl, IDStock, n1 as box, n2 as kg, nDet as dtl into inputStock from t_InputStock~ where Tanggal between " & lastDate & " and " & postDate
    'ExecMe s
    's = "select TanggalDetail as tgl, IDStock, JumlahBox as box, JumlahKG as kg, dtl into outputStock from t_SPPDetail~ where TanggalDetail between " & lastDate & " and " & postDate
    'ExecMe s
    's = "insert into inputStock(tgl, IDStock, box,  kg) select Tanggal, IDStock, SelisihBox, SelisihKG from t_StockOpname~ where SelisihBox>0 and Tanggal between " & lastDate & " and " & postDate
    'ExecMe s
    's = "insert into outputStock(tgl, IDStock, box,  kg) select Tanggal, IDStock, -SelisihBox, -SelisihKG from t_StockOpname~ where SelisihBox<0 and Tanggal between " & lastDate & " and " & postDate
    'ExecMe s
    
    s = "insert into mutasi_hist select * from mutasi"
    ExecMe s
    
    'debug.print GetTickCount - t
    t = GetTickCount
    
    s = "delete from a"
    ExecMe s
    s = "insert into a select " & postDate & " as  tgl, a.IDStock, iif(isnull(c.box),0,c.box) as inBox, iif(isnull(c.kg),0,c.kg) as inKg" & _
" , iif(isnull(d.box),0,d.box) as outBox, iif(isnull(d.kg),0,d.kg) as outKg, iif(isnull(b.akhirBox),0,b.akhirBox) as awalBox" & _
" , iif(isnull(b.akhirKg),0, b.akhirKg) as awalKg, iif(isnull(b.akhirBox),0, b.akhirBox) as akhirBox, iif(isnull(b.akhirKg),0, b.akhirKg) as akhirKg" & _
" , b.AkhirNoUrut as awalNoUrut, c.NoUrut as InNoUrut, d.NoUrut as OutNoUrut, b.AkhirNoUrut as AkhirNoUrut " & _
" , b.AkhirKgs as awalKgs, c.kgs as InKgs, d.kgs as OutKgs, b.AkhirKgs as AkhirKgs from (((" & _
" select IDStock from mutasi union select IDStock from inputStock union select IDStock from outputStock) a" & _
" left join mutasi b on a.IDStock=b.IDStock)" & _
" left join inputStock c on c.IDStock=a.IDStock)" & _
" left join outputStock d on d.IDStock=a.IDStock "
    ExecMe s
    s = "select * from a"
    query s
    Dim kgsStock As String
    Dim NoUrut As String
    Dim inBox As Long
    Dim outBox As Long
    Dim inKg As Double
    Dim outKg As Double
    Dim akhirBox As Long
    While Not RS.EOF
        If ifNullEmpty(RS.Fields("inKgs").value) <> "" Or ifNullEmpty(RS.Fields("outKgs").value) <> "" Then
            kgsStock = IIf(ifNullEmpty(RS.Fields("awalKgs").value) = "", "", "_" & RS.Fields("awalKgs").value)
            kgsStock = kgsStock & IIf(ifNullEmpty(RS.Fields("inKgs").value) = "", "", "_" & RS.Fields("inKgs").value)
            If Not IsNull(RS.Fields("outKgs").value) Then
                genKgsAkhir kgsStock, RS.Fields("outKgs").value
            End If
            kgsStock = Mid(kgsStock, 2)
            NoUrut = IIf(ifNullEmpty(RS.Fields("awalNoUrut").value) = "", "", "_" & RS.Fields("awalNoUrut").value)
            NoUrut = NoUrut & IIf(ifNullEmpty(RS.Fields("inNoUrut").value) = "", "", "_" & RS.Fields("inNoUrut").value)
            If Not IsNull(RS.Fields("outNoUrut").value) Then
                genNoUrutAkhir NoUrut, RS.Fields("outNoUrut").value
            End If
            NoUrut = Mid(NoUrut, 2)
            
            s = "update a set akhirNoUrut=@akhirNoUrut, akhirKgs=@akhirKgs, akhirBox=@akhirBox, akhirKg=@akhirKg where IDStock=@IDStock"
            s = Replace(s, "@IDStock", RS.Fields("IDStock").value)
            inBox = ifNullZero(RS.Fields("inBox").value)
            inKg = ifNullZero(RS.Fields("inKg").value)
            outBox = ifNullZero(RS.Fields("outBox").value)
            outKg = ifNullZero(RS.Fields("outKg").value)
            akhirBox = ifNullZero(RS.Fields("awalBox").value) + inBox - outBox
            s = Replace(s, "@akhirBox", akhirBox)
            If akhirBox = 0 Then kgsStock = ""
            s = Replace(s, "@akhirKgs", "'" & kgsStock & "'")
            s = Replace(s, "@akhirNoUrut", "'" & NoUrut & "'")
            s = Replace(s, "@akhirKg", cNum(ifNullZero(RS.Fields("awalKg").value) + inKg - outKg))
            
            ExecMe s
        End If
        
        
        RS.MoveNext
    Wend
    'Debug.Print GetTickCount - t
    t = GetTickCount
    
    's = " select @tgl as tgl, IDStock, sum(iif(tgl=@tgl, inBox,0)) as inBox, sum(iif(tgl=@tgl, inKg,0)) as inKg" & _
'" , sum(iif(tgl=@tgl, outBox,0)) as outBox, sum(iif(tgl=@tgl, outKg,0)) as outKg" & _
'" , sum(akhirBox)+sum(inBox)-sum(outBox) as akhirBox" & _
'" , sum(akhirKg)+sum(inKg)-sum(outKg) as akhirKg" & _
'"   into a from (" & _
'" select tgl, IDStock, akhirBox, akhirKg, 0 as inBox, 0.0 as inKg, 0 as outBox, 0.0 as outKg from mutasi" & _
'" union all select tgl, IDStock, 0,0,box,kg, 0,0 from inputStock " & _
'" union all select tgl, IDStock, 0,0,0,0,box, kg from outputStock ) as a" & _
'" group by IDStock"
'    s = Replace(s, "@tgl", cD(fAwal.Text))
'    ExecMe s
    s = "delete from mutasi"
    ExecMe s
    s = "insert into mutasi select * from a"
    ExecMe s
    Debug.Print GetTickCount - t
    t = GetTickCount

End Sub
Private Sub cmdPost_Click()
    Dim postDate As Long

    postDate = cD(fAwal.Text)
    While postDate < cD(date)
        post postDate
        postDate = Add_Tanggal2(postDate, 1)
    Wend
    
    
    BuatKartuStock
End Sub
Private Function ifNullZero(ByVal obj As Variant)
    If IsNull(obj) Then
        ifNullZero = 0
        Exit Function
    End If
    ifNullZero = obj
End Function
Private Function ifNullEmpty(ByVal obj As Variant)
    If IsNull(obj) Then
        ifNullEmpty = ""
        Exit Function
    End If
    ifNullEmpty = obj
End Function
Private Sub genKgsAkhir(ByRef kgs As String, kgsOut As String)
    If kgs = "_" Then Exit Sub
    Dim kg() As String
    kg = Split(kgsOut, "_")

    For i = 0 To UBound(kg)
        kgs = Replace(kgs, "_" & kg(i), "", 1, 1)
    Next
    
End Sub
Private Sub genNoUrutAkhir(ByRef NoUrut As String, NoUrutOut As String)
    If NoUrut = "_" Then Exit Sub
    Dim no_urut() As String
    no_urut = Split(NoUrutOut, "_")

    For i = 0 To UBound(no_urut)
        NoUrut = Replace(NoUrut, "_" & no_urut(i), "", 1, 1)
    Next
    
End Sub
Private Sub fAkhir_Validate(Cancel As Boolean)
    If cD(fAkhir) = "A" Then fAkhir = pServerDate
End Sub

Private Sub fCopy_Click()
    CopyGrid TDBGrid1
End Sub

Private Sub InputStock(ByVal startDate As Long, ByVal endDate As Long)
    s = "delete from inputStock"
    ExecMe s
    s = "select iif(isnull(b.TanggalGudang), b.Tanggal, b.TanggalGudang) as tgl, b.IDStock, a.NoUrut, a.Kg from t_InputStockDetail~ as a" & _
        " left join t_InputStock~ as b on a.NoBukti=b.NoBukti where b.Tanggal=" & startDate & " order by b.IdStock"
    query s
    If RS.RecordCount = 0 Then Exit Sub
    
    seb = RS.Fields("IDStock").value
    Dim kgs As String
    Dim NoUrut As String
    Dim jumlahBox As Long
    Dim jumlahKG As Double
    Dim Tgl As Long
    kgs = ""
    NoUrut = ""
    While Not RS.EOF
        If seb <> RS.Fields("IDStock").value Then
            If kgs <> "" Then
                kgs = Mid(kgs, 2)
                NoUrut = Mid(NoUrut, 2)
                s = "insert into inputStock values(" & Tgl & _
                    "," & seb & ", " & jumlahBox & "," & cNum(jumlahKG) & _
                    ",'" & kgs & "','" & NoUrut & "')"
                ExecMe s
            End If
            kgs = ""
            NoUrut = ""
            seb = RS.Fields("IDStock").value
                
            jumlahBox = 0
            jumlahKG = 0
        End If
        kgs = kgs & "_" & RS.Fields("kg").value
        NoUrut = NoUrut & "_" & CLng(RS.Fields("NoUrut").value)
        Tgl = RS.Fields("tgl").value
            
        jumlahBox = jumlahBox + 1
        jumlahKG = jumlahKG + RS.Fields("kg").value
        RS.MoveNext
    Wend
    kgs = Mid(kgs, 2)
    NoUrut = Mid(NoUrut, 2)
    s = "insert into inputStock values(" & Tgl & _
        "," & seb & ", " & jumlahBox & "," & cNum(jumlahKG) & _
        ",'" & kgs & "','" & NoUrut & "')"
    ExecMe s
End Sub

Private Sub OutputStock(ByVal startDate As Long, ByVal endDate As Long)
    s = "delete from outputStock"
    ExecMe s
    s = "select TanggalDetail as tgl, IDStock, jumlahBox as box, jumlahKg as kg, dtl, dtl2 from t_SPPDetail~ where TanggalDetail between " & startDate & " and " & endDate & " order by IdStock"
    query s
    If RS.RecordCount = 0 Then Exit Sub
    seb = RS.Fields("IDStock").value
    Dim kgs As String
    Dim NoUrut As String
    Dim jumlahBox As Long
    Dim jumlahKG As Double
    Dim Tgl As Long
    kgs = ""
    NoUrut = ""
    While Not RS.EOF
        If seb <> RS.Fields("IDStock").value Then
            If kgs <> "" Then kgs = Mid(kgs, 2)
            If NoUrut <> "" Then NoUrut = Mid(NoUrut, 2)
            s = "insert into outputStock values(" & Tgl & _
                "," & seb & ", " & jumlahBox & "," & cNum(jumlahKG) & _
                ",'" & kgs & "','" & NoUrut & "')"
            ExecMe s
            kgs = ""
            NoUrut = ""
            seb = RS.Fields("IDStock").value
                
            jumlahBox = 0
            jumlahKG = 0
        End If
        If ifNullEmpty(RS.Fields("dtl2").value) <> "" Then kgs = kgs & "_" & RS.Fields("dtl2").value
        If ifNullEmpty(RS.Fields("dtl").value) <> "" Then NoUrut = NoUrut & "_" & RS.Fields("dtl").value
        Tgl = RS.Fields("tgl").value
            
        jumlahBox = jumlahBox + RS.Fields("box").value
        jumlahKG = jumlahKG + RS.Fields("kg").value
        RS.MoveNext
    Wend
    If kgs <> "" Then kgs = Mid(kgs, 2)
    If NoUrut <> "" Then NoUrut = Mid(NoUrut, 2)
    s = "insert into outputStock values(" & Tgl & _
        "," & seb & ", " & jumlahBox & "," & cNum(jumlahKG) & _
        ",'" & kgs & "','" & NoUrut & "')"
    ExecMe s
End Sub


Private Sub fDetail_Click()
    BuatKartuStock
End Sub

Private Sub fExcel_Click()
On Error GoTo err
Dim res2() As Variant
Dim ColumnCount As Byte
    BuatKartuStock
    DoEvents
    If Not fDetail Then
        'FormPreview.SetData Me, "M utasi", fAwal & " - " & fAkhir, res2
        FormPreview.LoadFromData Me, "Mutasi", x, fAwal
    End If
err:
End Sub

Private Sub fKode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DoQuery
    ElseIf KeyCode = 112 Then
        HelpMe "Kode Barang", Me
    End If
End Sub

Sub FormHelpKeyDown(ByVal tVal As String)
    ActiveControl.Text = tVal
    DoQuery
End Sub

Private Sub Check1_Click()
    DoQuery
End Sub

Private Sub Check2_Click()
    DoQuery
End Sub

Private Sub fAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub Form_Resize()
On Error Resume Next
    fKet.Width = ScaleWidth - 2 * fKet.Left
    TDBGrid1.Width = fKet.Width
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Function MyFilter() As String
Dim p As Boolean
Dim MFilter As String
    MyFilter = ""
    MFilter = ""
    For i = 0 To Text1.count - 1
        If Trim(Text1(i)) <> "" Then MFilter = MFilter & " and b." & Text1(i).Tag & " like '" & Text1(i) & "'"
    Next
    If Trim(fKode) <> "" Then
        If fTidakAda Then
            MFilter = MFilter & " and b.KodeBarang not like '%" & fKode & "%'"
        Else
            If dLKode.value = 0 Then
                MFilter = MFilter & " and b.KodeBarang like '%" & fKode & "%'"
            Else
                MFilter = MFilter & " and b.KodeBarang like '" & fKode & "'"
            
            End If
            
        End If
    End If

    MyFilter = MFilter & MyFilter
End Function

Sub DoQuery()
    BuatKartuStock
    TDBGrid1_RowColChange -1, -1
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    x.ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x.DeleteRows 0
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    fAwal = pServerDate
    fAkhir = fAwal
    iIdStock = col1("IdStock").ColIndex
    iNamaBarang = col1("Nama Barang").ColIndex
    iTanggal = col1("Tanggal").ColIndex
    iMasuk = col1("Masuk").ColIndex
    iMasukKG = col1("MasukKG").ColIndex
    iKeluar = col1("Keluar").ColIndex
    iKeluarKG = col1("KeluarKG").ColIndex
    iAkhir = col1("Akhir").ColIndex
    iAkhirKG = col1("AkhirKG").ColIndex
    iNoBukti = col1("NoBukti").ColIndex
    iKet = col1("Keterangan").ColIndex
    iSatBesar = col1("SatBesar").ColIndex
    iSatKecil = col1("SatKecil").ColIndex
    iJenis = col1("Jenis").ColIndex
    iKode = col1("Kode").ColIndex
    col1("Tanggal").Tag = "Date"
    col1("Masuk").Tag = "Integer"
    col1("Keluar").Tag = "Integer"
    col1("Akhir").Tag = "Integer"
    col1("MasukKG").Tag = "Decimal"
    col1("KeluarKG").Tag = "Decimal"
    col1("AkhirKG").Tag = "Decimal"
    TDBGridLoad TDBGrid1
    TDBGridSetVisible TDBGrid1, "Jenis@Kode", False
    TDBGrid1.Rebind
    TDBGrid1.FetchRowStyle = True
    For i = 0 To TDBGrid1.Columns.count - 1
        TDBGrid1.ColumnFooters = True
    Next
    'BuatKartuStock
End Sub
Sub BuatKartuStock()
    Dim s As String
    MousePointer = vbHourglass
    Dim awal As Long
    awal = cD(fAwal)
    Dim akhir As Long
    Dim table As String
    akhir = awal
    s = "select * from mutasi where tgl=" & awal
    query s
    If RS.RecordCount > 0 Then
        table = "mutasi"
    Else
        table = "mutasi_hist"
    End If
    If fDetail Then
        s1 = "select Tanggal, NoBukti, Lot as Keterangan, IdStock, n1 as inBox, n2 as inKg,0 as outBox,0.0 as outKg,0 as akhirBox,0.0 as akhirKg from t_InputStock~ where Tanggal between " & awal & " and " & akhir
        s1 = s1 & " union all "
        s1 = s1 & "select a.TanggalDetail, a.NoSJ, b.Nama as Keterangan, a.IdStock, 0 as inBox, 0 as inKg, a.JumlahBox as outBox, a.JumlahKG as outKg,0 as akhirBox,0.0 as akhirKg from t_SPPDetail~ as a left join m_customer as b on a.KodeCustomerDetail=b.Kode where a.TanggalDetail between " & awal & " and " & akhir & " and a.statusDetail>1"
        s1 = s1 & " union all "
        s1 = s1 & "select a.Tanggal, '', 'STOCK OPNAME' as Keterangan, a.IdStock, iif(a.SelisihBox>0, a.SelisihBox, 0) as inBox, iif(a.SelisihKg>0, a.SelisihKg, 0) as inKg, iif(a.SelisihBox<0, -a.SelisihBox, 0) as outBox, iif(a.SelisihKg<0, -a.SelisihKg, 0) as outKg,0 as akhirBox,0.0 as akhirKg from t_StockOpname~ as a where a.Tanggal between " & awal & " and " & akhir
        s1 = s1 & " union all "
        s1 = s1 & "select tgl, '', 'AKHIR' as Keterangan, IdStock, 0,0,0,0, akhirBox-inBox+outBox, akhirKg-inKg+outKg from " & table & " where tgl=" & awal
        
        s = "select a.IdStock, b.Jenis&' '&b.KodeBarang&' '&b.Warna&' '&b.NoWarna&' '&b.Tube&' '&b.Grade as NamaBarang"
        s = s & ", a.Tanggal, a.inBox, a.inKg, a.outBox, a.outKg, a.akhirBox, a.akhirKg"
        s = s & ", a.NoBukti, a.Keterangan, b.SatBesar, b.SatKecil, b.Jenis, b.KodeBarang"
        s = s & " from (" & s1 & ") as a left join m_stock~ as b on a.IdStock=b.IdStock"
        s = s & " where 1=1" & MyFilter
        s = s & " order by b.IdJenis, b.IdKodeBarang, b.SatBesar, b.D, b.F, b.KodeBarang, b.IdGrade, b.WarnaDasar, b.NoWarna, b.IdTube, a.Tanggal, a.outKG, a.inKg"
    Else
        s = "select a.IdStock, b.Jenis&' '&b.KodeBarang&' '&b.Warna&' '&b.NoWarna&' '&b.Tube&' '&b.Grade as NamaBarang"
        s = s & ", a.tgl, a.inBox, a.inKg, a.outBox, a.outKg, a.akhirBox, a.akhirKg"
        s = s & ", '', '', b.SatBesar, b.SatKecil, b.Jenis, b.KodeBarang, a.awalKgs, a.inKgs, a.outKgs, a.akhirKgs, a.awalNoUrut, a.inNoUrut, a.outNoUrut, a.akhirNoUrut"
        s = s & " from " & table & " as a left join m_stock~ as b on a.IdStock=b.IdStock"
        s = s & " where a.Tgl=" & awal & " and (a.inKg<>0 or a.outKg<>0 or a.akhirKg<>0)" & MyFilter
        s = s & " order by b.IdJenis, b.IdKodeBarang, b.SatBesar, b.D, b.F, b.KodeBarang, b.IdGrade, b.WarnaDasar, b.NoWarna, b.IdTube, a.Tgl, a.outKG"
    End If
    query s
    If RS.RecordCount = 0 Then
        x.ReDim 0, -1, 0, -1
        TDBGrid1.Rebind
        MousePointer = vbDefault
        
        Exit Sub
    End If
    Dim y As New XArrayDB
    y.LoadRows RS.GetRows
    x.ReDim 0, y.UpperBound(1), 0, y.UpperBound(2)
    Dim i As Integer
    Dim k As Integer
    For i = 0 To x.UpperBound(1)
        If check_flag(i, y) Then
            For j = 0 To y.UpperBound(2)
                x(k, j) = y(i, j)
                
            Next
            k = k + 1
        End If
    Next
    x.ReDim 0, k - 1, 0, x.UpperBound(2)
    
    TDBGridSetVisible TDBGrid1, "Tanggal@NoBukti@Keterangan", fDetail
    'TDBGridSetVisible TDBGrid1, "IdStock"
    TDBGrid1.Rebind
    MousePointer = vbDefault
End Sub

Private Sub fTidakAda_Click()
    DoQuery
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(Bookmark, col1("Masuk").ColIndex) = 0 And x(Bookmark, col1("Keluar").ColIndex) = 0 Then
        RowStyle.ForeColor = RGB(0, 155, 0)
    ElseIf x(Bookmark, col1("Masuk").ColIndex) = 0 Then
        RowStyle.ForeColor = RGB(155, 0, 0)
    ElseIf x(Bookmark, col1("Keluar").ColIndex) = 0 Then
        RowStyle.ForeColor = RGB(0, 0, 155)
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If LastRow <> TDBGrid1.Bookmark Then
        fKet = col1("Nama Barang").Text
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DoQuery
    End If
End Sub

Private Function check_flag(i As Integer, Optional ByVal x1 As XArrayDB = Null)
    If IsNull(x1) Then
        Set x1 = x
    End If
    If Check1 Then
        If x1(i, iAkhirKG) <> 0 Then
            check_flag = True
            Exit Function
        End If
    End If
    If Check2 Then
        If x1(i, iMasukKG) <> 0 Then
            check_flag = True
            Exit Function
        End If
        If x1(i, iKeluarKG) <> 0 Then
            check_flag = True
            Exit Function
        End If
    End If
    check_flag = False
End Function
Private Sub print_result(k As Integer, ByVal i As Integer, ByVal awal As Double, ByVal awalkg As Double, ByVal masuk As Double, ByVal masukkg As Double, ByVal keluar As Double, ByVal keluarkg As Double)
    If Not check_flag(i) Then
        k = k - 1
        Exit Sub
    End If
    For j = 0 To x.UpperBound(2)
        x(k, j) = x(i - 1, j)
    Next
    x(k, iMasuk) = masuk
    x(k, iMasukKG) = masukkg
    x(k, iKeluar) = keluar
    x(k, iKeluarKG) = keluarkg
    x(k, iAkhir) = awal + masuk - keluar
    x(k, iAkhirKG) = awalkg + masukkg - keluarkg
End Sub

VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{F6D22ACD-8630-4FE1-97C4-D56AB4AD4DEA}#1.0#0"; "usrtruecombo.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormNPB 
   BackColor       =   &H00FFC0C0&
   Caption         =   "NOTA PENERIMAAN BARANG"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Tag             =   "18"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fBatal 
      Caption         =   "&BATAL"
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton fOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton fList 
      Caption         =   "&LIST"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton fNew 
      Caption         =   "&NEW"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin UsrText.IText fNoNPB 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   1440
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
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO BTB"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NAMA BARANG"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "QTY Gudang"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "QTY Ambil"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "SATUAN"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "IdStock"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "StatusProses"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "IdDet"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=2359295"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5318"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5239"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=2359295"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1984"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1905"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=70518768"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1614"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1535"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1376"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1296"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(6)._MinWidth=75649680"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=74,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=40,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=60,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=57,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=58,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=59,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=44,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=41,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=42,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=43,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=64,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=61,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=62,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=63,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=24,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=48,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=45,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=46,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=47,.parent=15"
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
      _StyleDefs(79)  =   "Named:id=67:RecordSelector"
      _StyleDefs(80)  =   ":id=67,.parent=30"
      _StyleDefs(81)  =   "Named:id=70:FilterBar"
      _StyleDefs(82)  =   ":id=70,.parent=29"
   End
   Begin UsrTrueCombo.ITrueCombo fDepartemen 
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
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
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
      MaxLength       =   10
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No NPB"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Detil Barang"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Departemen"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA PENERIMAAN BARANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "FormNPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim m_mode As String
Dim col1 As TrueOleDBGrid80.Columns

Private Sub fList_Click()
    If Not cekValid("SEE", Tag) Then Exit Sub
    FormList.LoadMe "NPB", _
"select NoNPB, TanggalNPB, Dept from t_NPB where 1=1", _
"TanggalNPB@Departemen", "TanggalNPB@Dept", _
"1000@2000", "Date@String", _
"NO NPB@TANGGAL@DEPARTEMEN", _
"1500@1000@1500", _
"String@Date@String", Me, " order by NoNPB Desc"
End Sub

Private Sub ClearScreen()
On Error Resume Next
    fNoNPB = ""
    fTanggal = pServerDate
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    Set TDBGrid1.Array = x
    TDBGrid1.Rebind
    fDepartemen.SetFocus
End Sub

Sub GetResult(ByVal tNo As String)
On Error Resume Next
    a = "select NoNPB, TanggalNPB, Dept, Status from t_NPB where NoNPB='" & esc(tNo) & "'"
    query a
    If RS.RecordCount < 1 Then
        Exit Sub
    End If
    m_mode = "SEE"
    GoEvent
    fNoNPB = RS.Fields("NoNPB").Value
    fTanggal = cTanggal(RS.Fields("TanggalNPB").Value)
    fDepartemen = RS.Fields("Dept").Value
    fDepartemen.FindIndex
    fOK.Enabled = RS.Fields("Status").Value = 0
    fBatal.Enabled = Not fOK.Enabled
    fDelete.Enabled = fOK.Enabled
With TDBGrid1
    a = "select NoBTB, NamaBarang, QTYGudang, QTY, Satuan, m_stockBeli.IdStock, TipeStock, IdDet from t_NPBDetail left join m_StockBeli on t_NPBDetail.IdStock=m_stockBeli.IdStock where NoNPB='" & esc(tNo) & "'"
    query a
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    .Rebind
    .SetFocus
End With
End Sub

Private Sub fNew_Click()
    ClearScreen
    m_mode = "NEW"
    GoEvent
End Sub

Private Sub GoEvent()
    v = m_mode = "NEW"
    fSave.Enabled = v
    fDelete.Enabled = Not v
    fPrint.Enabled = Not v
    fOK.Enabled = Not v
    fBatal.Enabled = Not v
    fNoNPB.Enabled = v
    fTanggal.Enabled = v
    fDepartemen.Enabled = v
    col1("QTY Ambil").Locked = Not v
    TDBGrid1.AllowUpdate = v
    TDBGrid1.AllowDelete = v
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    Set TDBGrid1.Array = x
    TDBGrid1.AllowDelete = True
    TDBGrid1.Rebind
    fDepartemen.ZOrder 0
    TDBGrid1.HeadingStyle.Alignment = dbgCenter
    
    fDepartemen.SetHeader "Departemen@Kode"
    fDepartemen.SetWidth "2500@1000"
    fDepartemen.SetType "String@String"
    a = "select Departemen, KdDept from m_Departemen order by Departemen"
    query a
Dim rs1() As Variant
    rs1 = RS.GetRows
    fDepartemen.SetDB rs1
    
    col1("QTY Gudang").Alignment = dbgRight
    col1("QTY Ambil").Alignment = dbgRight
        
    col1("NAMA BARANG").Locked = True
    col1("SATUAN").Locked = True
    col1("NO BTB").Locked = True
    col1("IdStock").Visible = False
    col1("QTY Gudang").Locked = True
    col1("StatusProses").Visible = False
    col1("IdDet").Visible = False
    fNew_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fPrint_Click()
    'FormReport.LoadMe "NPB.rpt", fNoNPB
End Sub

Private Sub fSave_Click()
On Error GoTo err
    BeginTransaction
    fDepartemen.Validate
With TDBGrid1
    .Update
    a = "insert into t_NPB(NoNPB,TanggalNPB,Dept) values('" & _
        fNoNPB & _
        "'," & cD(fTanggal) & _
        ",'" & fDepartemen & "')"
    If ExecMe(a) = 0 Then GoTo err
    For i = 0 To x.UpperBound(1)
        If x(i, .Columns("QTY Ambil").ColIndex) > 0 Then
            a = "update t_BTBDetail set NoNPB=NoNPB&'," & fNoNPB & "',QTYAmbil=QTYAmbil+" & cNum(x(i, .Columns("QTY Ambil").ColIndex)) & " where NoBTB='" & esc(x(i, .Columns("No BTB").ColIndex)) & "' and IdStock=" & x(i, .Columns("IdStock").ColIndex)
            ExecMe a
            a = "insert into t_NPBDetail(NoBTB,NoNPB,TanggalNPB,QTY,QTYGudang,Dept,IdStock) values('" & _
                x(i, .Columns("No BTB").ColIndex) & _
                "','" & fNoNPB & _
                "'," & cD(fTanggal) & _
                "," & cNum(x(i, .Columns("QTY Ambil").ColIndex)) & _
                "," & cNum(x(i, .Columns("QTY Gudang").ColIndex)) & _
                ",'" & fDepartemen & _
                "'," & x(i, .Columns("IdStock").ColIndex) & ")"
            If ExecMe(a) = 0 Then GoTo err
        End If
    Next
End With
    CommitTransaction
    MsgBox "SUKSES"
    GetResult fNoNPB
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fOK_Click()
On Error GoTo err
    BeginTransaction
    For i = 0 To x.UpperBound(1)
        If x(i, col1("StatusProses").ColIndex) = 1 Then
            a = "update m_StockProses set Jumlah=Jumlah+" & cNum(x(i, col1("QTY Ambil").ColIndex)) & " where IdStock=" & x(i, col1("IdStock").ColIndex) & " and Dept='" & esc(fDepartemen) & "'"
            If ExecMe(a) = 0 Then GoTo err
            a = "insert into t_Out(Tanggal, IdStock, QTY, Dept, DeptAsal, FromNPB, StatusPindahStock) values(" & _
                cD(fTanggal) & _
                "," & x(i, col1("IdStock").ColIndex) & _
                "," & cNum(x(i, col1("QTY Ambil").ColIndex)) & _
                ",'" & esc(fDepartemen) & _
                "','GUDANG'," & x(i, col1("IdDet").ColIndex) & ",1)"
            If ExecMe(a) = 0 Then GoTo err
        End If
        a = "update m_stockBeli set Jumlah=Jumlah-" & cNum(x(i, col1("QTY Ambil").ColIndex)) & " where IdStock=" & x(i, col1("IdStock").ColIndex)
        If ExecMe(a) = 0 Then GoTo err
    Next
    a = "update t_NPB set Status=1 where NoNPB='" & esc(fNoNPB) & "'"
    ExecMe a
    a = "update t_NPBDetail set Status=1 where NoNPB='" & esc(fNoNPB) & "'"
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    GetResult fNoNPB
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fBatal_Click()
On Error GoTo err
    BeginTransaction
    a = "update t_NPB set Status=0 where NoNPB='" & esc(fNoNPB) & "'"
    ExecMe a
    a = "update t_NPBDetail set Status=0 where NoNPB='" & esc(fNoNPB) & "'"
    ExecMe a
    For i = 0 To x.UpperBound(1)
        If x(i, col1("StatusProses").ColIndex) = 1 Then
            a = "update m_stockProses set Jumlah=Jumlah-" & cNum(x(i, col1("QTY Ambil").ColIndex)) & " where IdStock=" & x(i, col1("IdStock").ColIndex) & " and Dept='" & esc(fDepartemen) & "'"
            If ExecMe(a) = 0 Then GoTo err
            a = "delete from t_Out where FromNPB=" & x(i, col1("IdDet").ColIndex)
            If ExecMe(a) = 0 Then GoTo err
        End If
        a = "update m_stockBeli set Jumlah=Jumlah+" & cNum(x(i, col1("QTY Ambil").ColIndex)) & " where IdStock=" & x(i, col1("IdStock").ColIndex)
        If ExecMe(a) = 0 Then GoTo err
    Next
    CommitTransaction
    MsgBox "SUKSES"
    GetResult fNoNPB
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fDelete_Click()
On Error GoTo err
    If Not cekValid("DELETE", Tag) Then Exit Sub
    BeginTransaction
    With TDBGrid1
Dim NoNPB As String
    For i = 0 To x.UpperBound(1)
        a = "select top 1 FromNPB from t_In where FromNPB=" & x(i, col1("IdDet").ColIndex)
        query a
        If RS.RecordCount > 0 Then
            MsgBox "Sudah ada di Nota Input Proses!!!"
            GoTo err
        End If
        a = "select NoNPB from t_BTBDetail where NoBTB='" & esc(x(i, .Columns("No BTB").ColIndex)) & "'"
        query a
        NoNPB = Replace(RS.Fields(0).Value & "", "," & fNoNPB, "", 1, 1)
        a = "update t_BTBDetail set NoNPB='" & NoNPB & "', QTYAmbil=QTYAmbil-" & cNum(x(i, .Columns("QTY Ambil").ColIndex)) & " where IdStock=" & x(i, .Columns("IdStock").ColIndex) & " and NoBTB='" & esc(x(i, .Columns("NO BTB").ColIndex)) & "'"
        ExecMe a
    Next
    a = "delete from t_NPB where NoNPB='" & esc(fNoNPB) & "'"
    ExecMe a
    a = "delete from t_NPBDetail where NoNPB='" & esc(fNoNPB) & "'"
    ExecMe a
    End With
    CommitTransaction
    MsgBox "SUKSES"
    m_mode = "NEW"
    GoEvent
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fDepartemen_LostFocus()
    If fDepartemen = "" Or Not fDepartemen.Enabled Then Exit Sub
    a = "select max(NoNPB) from t_NPB where TanggalNPB>" & pAddNoLong & " and Dept='" & esc(fDepartemen) & "'"
    query a
    If IsNull(RS.Fields(0).Value) Then
        fNoNPB = "00001/" & zerofill(fDepartemen.GetData("Kode"), 2) & "/" & Right(pServerDate, 2)
    Else
        fNoNPB = zerofill(Left(RS.Fields(0).Value, 5) + 1, 5) & "/" & zerofill(fDepartemen.GetData("Kode"), 2) & "/" & Right(pServerDate, 2)
    End If
    a = "select NoBTB, NamaBarang, QTY-QTYAmbil, QTY-QTYAmbil, Satuan, m_stockBeli.IdStock, m_stockBeli.TipeStock from t_BTBDetail left join m_StockBeli on t_BTBDetail.IdStock=m_stockBeli.IdStock where QTY>QTYAmbil and t_BTBDetail.Dept='" & esc(fDepartemen) & "'"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub


VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormStockProses 
   BackColor       =   &H00FFC0C0&
   Caption         =   "STOCK PROSES"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   4
      Left            =   6360
      TabIndex        =   9
      Tag             =   "Departemen"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   3
      Left            =   4800
      TabIndex        =   8
      Tag             =   "Grade"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Tag             =   "NoWarna"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Tag             =   "KodeBarang"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "NamaBarang"
      Top             =   360
      Width           =   1455
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7223
      _LayoutType     =   4
      _RowHeight      =   -2147483647
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
      Columns(2).Caption=   "Kode Barang"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No Warna"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Grade"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Dept"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Satuan"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Jumlah"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Terpakai"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1032"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4498"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4419"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1693"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1429"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1349"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=900"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=820"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2064"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1984"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1032"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=953"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=1693"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1614"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=1984"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1905"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(8)._MinWidth=166480384"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
      DeadAreaBackColor=   14215660
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
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=74,.parent=73"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=77,.parent=76"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=82,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=24,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=52,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=49,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=51,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=44,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=41,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=42,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=43,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=60,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=57,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=59,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=64,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=61,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=62,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=63,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=72,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=69,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=70,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=71,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=40,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=37,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=38,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=39,.parent=15"
      _StyleDefs(66)  =   "Named:id=29:Normal"
      _StyleDefs(67)  =   ":id=29,.parent=0"
      _StyleDefs(68)  =   "Named:id=30:Heading"
      _StyleDefs(69)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=30,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=31:Footing"
      _StyleDefs(72)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=32:Selected"
      _StyleDefs(74)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=33:Caption"
      _StyleDefs(76)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(77)  =   "Named:id=34:HighlightRow"
      _StyleDefs(78)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(79)  =   "Named:id=35:EvenRow"
      _StyleDefs(80)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=36:OddRow"
      _StyleDefs(82)  =   ":id=36,.parent=29"
      _StyleDefs(83)  =   "Named:id=75:RecordSelector"
      _StyleDefs(84)  =   ":id=75,.parent=30"
      _StyleDefs(85)  =   "Named:id=78:FilterBar"
      _StyleDefs(86)  =   ":id=78,.parent=29"
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Departemen"
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No Warna"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FormStockProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim f As Form


Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    Set TDBGrid1.Array = x
    TDBGrid1.AllowUpdate = False
    TDBGrid1.Rebind
    TDBGrid1.Columns("Jumlah").Tag = "Decimal"
    TDBGrid1.Columns("Terpakai").Tag = "Decimal"
    TDBGridLoad TDBGrid1
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        f.SetOtherRowData TDBGrid1.Columns("IdStock").Value
        Visible = False
    ElseIf KeyCode = 120 Then
        f.SetOtherRowData TDBGrid1.Columns("IdStock").Value
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ShowResult
    End If
End Sub

Private Function MyFilter()
    MyFilter = ""
    If Trim(Text(0)) <> "" Then MyFilter = MyFilter & " and NamaBarang like '%" & Text(0) & "%'"
    If Trim(Text(1)) <> "" Then MyFilter = MyFilter & " and KodeBarang like '" & Text(1) & "%'"
    If Trim(Text(2)) <> "" Then MyFilter = MyFilter & " and NoWarna like '" & Text(2) & "%'"
    If Trim(Text(3)) <> "" Then MyFilter = MyFilter & " and Grade like '" & Text(3) & "%'"
    If Trim(Text(4)) <> "" Then MyFilter = MyFilter & " and Dept like '" & Text(4) & "%'"
End Function

Private Sub ShowResult()
    a = "select m_StockProses.IdStock, NamaBarang, KodeBarang, NoWarna, Grade, Dept, Satuan, m_StockProses.Jumlah, JumlahTerpakai from m_StockProses left join m_stockBeli on m_StockProses.IdStock=m_StockBeli.IdStock where 1=1" & MyFilter
    query a
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    If RS.RecordCount <> 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub

Sub LoadMe(tForm As Form, Optional ByVal tDepartemen As String = "", Optional ByVal tNamaBarang As String = "")
    Me.WindowState = 0
    Set f = tForm
    Text(4) = tDepartemen
    Text(0) = tNamaBarang
    ShowResult
    Show , f
End Sub





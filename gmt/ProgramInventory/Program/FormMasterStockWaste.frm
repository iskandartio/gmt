VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMasterStockWaste 
   BackColor       =   &H00FFC0C0&   Caption         =   "STOCK WASTE"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Tag             =   "19"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton fStock 
      Caption         =   "EDIT STOCK"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   120
      Width           =   1575
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
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
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
      Columns(2).Caption=   "IdStock"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nama Barang"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "QTY"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Satuan"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Harga"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(4)._MinWidth=149"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(5)._MinWidth=54391100"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
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
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Named:id=33:Normal"
      _StyleDefs(59)  =   ":id=33,.parent=0"
      _StyleDefs(60)  =   "Named:id=34:Heading"
      _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=34,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=35:Footing"
      _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=36:Selected"
      _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=37:Caption"
      _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(69)  =   "Named:id=38:HighlightRow"
      _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=39:EvenRow"
      _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=40:OddRow"
      _StyleDefs(74)  =   ":id=40,.parent=33"
      _StyleDefs(75)  =   "Named:id=41:RecordSelector"
      _StyleDefs(76)  =   ":id=41,.parent=34"
      _StyleDefs(77)  =   "Named:id=42:FilterBar"
      _StyleDefs(78)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Click Update untuk Update Stock Waste"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "FormMasterStockWaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid80.Columns
Dim x As New XArrayDB

Private Sub fDelete_Click()
On Error GoTo err
    b = MsgBox("Yakin Hapus?", vbYesNo)
    If b = vbNo Then Exit Sub
    BeginTransaction
    a = "delete from m_StockWaste where IdStock=" & col1("cKey").Value
    If ExecMe(a) = 0 Then GoTo err
    a = "update m_StockBeli set StockWaste=0 where IdStock=" & col1("cKey").Value
    If ExecMe(a) = 0 Then GoTo err
    CommitTransaction
    TDBGrid1.Delete
    Exit Sub
err:
    RollBackTransaction
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub DoQuery()
    a = "select 0, m_StockWaste.IdStock, m_StockWaste.IdStock, NamaBarang, m_StockWaste.Jumlah, Satuan, m_StockWaste.Harga from m_StockWaste left join m_StockBeli on m_StockWaste.IdStock=m_StockBeli.IdStock order by NamaBarang"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    DoQuery
    col1("updated").Visible = False
    col1("cKey").Visible = False
    TDBGrid1.FetchRowStyle = True
    col1("Harga").NumberFormat = "Standard"
    col1("IdStock").Locked = True
    col1("Nama Barang").Locked = True
    col1("QTY").Locked = True
    col1("Satuan").Locked = True
    col1("Harga").Alignment = dbgRight
    col1("QTY").Alignment = dbgRight
    TDBGrid1.AllowAddNew = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fStock_Click()
    FormStockProses.LoadMe Me
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) = "1" Then
            If x(i, col1("cKey").ColIndex) = "" Then
                a = "insert into m_StockWaste(IdStock, Harga) values(" & _
                    x(i, col1("IdStock").ColIndex) & _
                    "," & cNum(x(i, col1("Harga").ColIndex)) & ")"
                If ExecMe(a) = 0 Then GoTo err
            Else
                a = "update m_StockWaste set IdStock=" & x(i, col1("IdStock").ColIndex) & ",Harga=" & cNum(x(i, col1("Harga").ColIndex)) & " where IdStock=" & x(i, col1("cKey").ColIndex)
                If ExecMe(a) = 0 Then GoTo err
            End If
            a = "update m_StockBeli set StockWaste=1 where IdStock=" & x(i, col1("IdStock").ColIndex)
            If ExecMe(a) = 0 Then GoTo err
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoQuery
    Exit Sub
err:
    CommitTransaction
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

Sub SetOtherRowData(ByVal tIdStock As Long)
    col1("Updated").Value = "1"
    a = "select IdStock, NamaBarang, Jumlah, Satuan from m_StockBeli where IdStock=" & tIdStock
    query a
    If RS.RecordCount > 0 Then
        col1("IdStock").Value = tIdStock
        col1("Nama Barang").Value = RS.Fields("NamaBarang").Value
        col1("QTY").Value = RS.Fields("Jumlah").Value
        col1("Satuan").Value = RS.Fields("Satuan").Value
    Else
        col1("IdStock").Value = ""
        col1("Nama Barang").Value = ""
        col1("QTY").Value = ""
        col1("Satuan").Value = ""
    End If
    TDBGrid1.SetFocus
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

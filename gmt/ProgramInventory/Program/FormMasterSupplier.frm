VERSION 5.00
Object = "{E2D3646A-2684-4DDE-BE47-3323E01328EE}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMasterSupplier 
   BackColor       =   &H00FFC0C0&   Caption         =   "MASTER SUPPLIER"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Tag             =   "23"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin UsrText.IText Text1 
      Height          =   270
      Left            =   3600
      TabIndex        =   5
      Top             =   240
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
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
      Columns(1).Caption=   "Supplier"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nama Supplier"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Alamat"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Kota"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Telepon"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Fax"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Contact Person"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Tempo Bayar"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Mata Uang"
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
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=3784"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3704"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2355"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(3)._MinWidth=74451708"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=1852"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=1773"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(4)._MinWidth=48"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=1852"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1773"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(5)._MinWidth=48"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=1826"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=1746"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(36)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(40)=   "Column(8)._MinWidth=75728656"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=52,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=78,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=54,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=58,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
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
      Left            =   2160
      TabIndex        =   2
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
      TabIndex        =   6
      Top             =   600
      Width           =   7935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipe"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Find"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FormMasterSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid80.Columns
Dim x As New XArrayDB

Private Sub fFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        fFind.Cancel = True
        Dim cCol As Integer
        cCol = TDBGrid1.Col
        n = TDBGrid1.Bookmark
        m = n + 1
        If m = x.UpperBound(1) + 1 Then m = 0
        Do While m <> n
            If InStr(1, x(m, cCol), fFind, vbTextCompare) <> 0 Then Exit Do
            m = m + 1
            If m = x.UpperBound(1) + 1 Then m = 0
        Loop
        If InStr(1, x(m, cCol), fFind, vbTextCompare) = 0 Then MsgBox "Tidak Ketemu"
        TDBGrid1.Bookmark = m
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    fKet.Width = TDBGrid1.Width
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    col1("updated").Visible = False
    col1("Supplier").Visible = False
    DoQuery
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
    TDBGrid1.FetchRowStyle = True
End Sub

Private Sub fSave_Click()
    UpdateDB
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    col1("updated").Value = 1
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If x(Bookmark, col1("updated").ColIndex) = 1 Then
        RowStyle.BackColor = vbYellow
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        fFind_KeyDown 13, 0
    End If
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub
Private Sub DoQuery()
    a = "select 0, Nama, Nama, Alamat, Kota, Telepon, Fax, ContactPerson, WaktuPembayaran, MataUang, Tipe from m_Supplier" & MyFilter & " order by Nama"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub

Private Function MyFilter() As String
    MyFilter = ""
    If Trim(Text1) <> "" Then MyFilter = " where Tipe like '%" & esc(Text1) & "%'"
End Function

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)
On Error GoTo err
    BeginTransaction
    a = "delete from m_Supplier where Nama='" & esc(x(TDBGrid1.Bookmark, col1("Nama Supplier").ColIndex)) & "'"
    ExecMe a
    CommitTransaction
    Exit Sub
err:
    RollBackTransaction
    Cancel = True
End Sub

Function UpdateDB() As Boolean
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) <> 0 Then
            If x(i, col1("Supplier").ColIndex) <> "" Then
                a = "update m_Supplier set Nama='" & x(i, col1("Nama Supplier").ColIndex) & _
                "', Alamat='" & x(i, col1("Alamat").ColIndex) & _
                "', Kota='" & x(i, col1("Kota").ColIndex) & _
                "', Telepon='" & x(i, col1("Telepon").ColIndex) & _
                "', Fax='" & x(i, col1("Fax").ColIndex) & _
                "', ContactPerson='" & x(i, col1("Contact Person").ColIndex) & _
                "', WaktuPembayaran='" & x(i, col1("Tempo Bayar").ColIndex) & _
                "', MataUang='" & x(i, col1("Mata Uang").ColIndex) & _
                "', Tipe='" & x(i, col1("Tipe").ColIndex) & "' where Nama='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                ExecMe a
                If x(i, col1("Nama Supplier").ColIndex) <> x(i, col1("Supplier").ColIndex) Then
                    a = "update t_PR set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_BTB set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_BTBDetail set NamaSupplierDetail='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplierDetail='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_Jasa set NamaPemberiJasa='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaPemberiJasa='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_JasaDetail set NamaPemberiJasa='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaPemberiJasa='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_NRPembelian set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_NRPembelianDetail set NamaSupplierDetail='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplierDetail='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_PembelianSTT set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_PembelianSTTPelunasan set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_PembelianSTTPembayaran set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_PembelianSTTPelunasan set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                    a = "update t_PembelianSTTPotongan set NamaSupplier='" & x(i, col1("Nama Supplier").ColIndex) & "' where NamaSupplier='" & esc(x(i, col1("Supplier").ColIndex)) & "'"
                    ExecMe a
                End If
            Else
                a = "insert into m_Supplier(Nama, Alamat, Kota, Telepon, Fax, ContactPerson, WaktuPembayaran, MataUang, Tipe) values('" & _
                x(i, col1("Nama Supplier").ColIndex) & _
                "','" & x(i, col1("Alamat").ColIndex) & _
                "','" & x(i, col1("Kota").ColIndex) & _
                "','" & x(i, col1("Telepon").ColIndex) & _
                "','" & x(i, col1("Fax").ColIndex) & _
                "','" & x(i, col1("Contact Person").ColIndex) & _
                "','" & x(i, col1("Tempo Bayar").ColIndex) & _
                "','" & x(i, col1("Mata Uang").ColIndex) & _
                "','" & x(i, col1("Tipe").ColIndex) & "')"
                If ExecMe(a) = 0 Then GoTo err
            End If
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoQuery
    Exit Function
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Function

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    fKet = col1(TDBGrid1.Col).Value
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DoQuery
    End If
End Sub

VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Object = "{F6D22ACD-8630-4FE1-97C4-D56AB4AD4DEA}#1.0#0"; "UsrTrueCombo.ocx"
Begin VB.Form FormPembelianJasa 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PEMBELIAN JASA"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Tag             =   "15"
   WindowState     =   2  'Maximized
   Begin UsrTrueCombo.ITrueCombo fPemberiJasa 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
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
   Begin UsrText.IText fTotal 
      Height          =   270
      Left            =   3360
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
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
      Enabled         =   0   'False
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5106
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama Jasa"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nilai"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Debet Acc"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Ket Acc"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=8652"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8573"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2328"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2249"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2090"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2011"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=10,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin GMT.iToolbar iToolbar1 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   873
   End
   Begin UsrText.IText fNo 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
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
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
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
   Begin UsrText.IText fNoFakturJasa 
      Height          =   270
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
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
   Begin UsrText.IText fMataUang 
      Height          =   270
      Left            =   6360
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   476
      Text            =   "RP"
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
      BackStyle       =   0  'Transparent
      Caption         =   "Mata Uang"
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "No Faktur Jasa"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pemberi Jasa"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No Bukti"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PEMBELIAN JASA"
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
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FormPembelianJasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim col1 As TrueOleDBGrid80.Columns
Dim LPemberiJasa As Boolean
Dim m_mode As String

Private Sub ClearScreen()
    fPemberiJasa = ""
    fNo = ""
    fTanggal = "__/__/__"
    fNoFakturJasa = ""
    fMataUang = "RP"
    fTotal = "0"
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
End Sub

Sub GetResult(ByVal tNo As String)
On Error Resume Next
Dim a As String
    ClearScreen
    fPemberiJasa_KeyDown 0, 0
    iToolbar1.SetQuick tNo
    a = "select NamaPemberiJasa,NoBukti, Tanggal, NamaJasa, Nilai, DebetAcc, KetAcc, NoFakturJasa, MataUang from t_JasaDetail where NoBukti='" & esc(tNo) & "'"
    query a
    If RS.RecordCount < 1 Then
        GoEvent "SEE"
        Exit Sub
    End If
    fPemberiJasa = RS.Fields("NamaPemberiJasa").Value
    fPemberiJasa.FindIndex
    fNo = RS.Fields("NoBukti").Value
    fTanggal = cTanggal(RS.Fields("Tanggal").Value)
    fNoFakturJasa = RS.Fields("NoFakturJasa").Value
    fMataUang = RS.Fields("MataUang").Value
    x.ReDim 0, RS.RecordCount - 1, 0, col1.Count - 1
    b = 0
    For i = 0 To RS.RecordCount - 1
        x(i, col1("Nama Jasa").ColIndex) = RS.Fields("NamaJasa").Value
        x(i, col1("Nilai").ColIndex) = RS.Fields("Nilai").Value
        x(i, col1("Debet Acc").ColIndex) = RS.Fields("DebetAcc").Value
        x(i, col1("Ket Acc").ColIndex) = RS.Fields("KetAcc").Value
        b = b + RS.Fields("Nilai").Value
        RS.MoveNext
    Next
    fTotal = cDecimal(b)
    TDBGrid1.Rebind
    GoEvent "SEE"
End Sub

Private Sub fNo_LostFocus()
    If m_mode = "NEW" Then iToolbar1.SetQuick SetNomor(fNo, fTanggal, "select max(NoBukti) from t_Jasa where Tanggal>" & pAddNoLong)
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    LPemberiJasa = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fPemberiJasa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
Dim a As String
    If Not LPemberiJasa Then
        a = "select Nama, Telepon from m_Supplier order by Nama"
        query a
        Dim rs1() As Variant
        rs1 = RS.GetRows
        fPemberiJasa.SetDB rs1
        fPemberiJasa.SetType "String"
        LPemberiJasa = True
    End If
err:
End Sub

Private Sub Form_Load()
Dim a As String
    Caption = Caption & "---" & pTipe
    TDBGrid1.AllowDelete = True
    TDBGrid1.AllowAddNew = True
    fPemberiJasa.SetHeader "Nama@Telepon"
    fPemberiJasa.SetWidth "2500@2500"
    fPemberiJasa.SetType "String@String"
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    a = "select Max(NoBukti) from t_Jasa where Tanggal>" & pAddNoLong
    query a
    If IsNull(RS.Fields(0).Value) Then
        iToolbar1.SetNoData True
        GoEvent "SEE"
    Else
        GetResult RS.Fields(0).Value
    End If
    col1("Nilai").Alignment = dbgRight
    col1("Nilai").NumberFormat = "Standard"
End Sub

Private Sub GoEvent(Optional ByVal tEvent As String = "")
Dim v As Boolean
    If tEvent = "" Then tEvent = m_mode Else m_mode = tEvent
    v = tEvent = "SEE"
    fPemberiJasa.Enabled = Not v
    fNo.Enabled = Not v
    fTanggal.Enabled = Not v
    fNoFakturJasa.Enabled = Not v
    TDBGrid1.AllowUpdate = Not v
    TDBGrid1.AllowDelete = Not v
    fMataUang.Enabled = Not v
End Sub

Private Sub fTanggal_LostFocus()
    fNo_LostFocus
End Sub

Private Sub iToolbar1_NewClick()
On Error Resume Next
Dim a As String
Dim b As Long
    ClearScreen
    a = "select Max(NoBukti) from t_Jasa where Tanggal>" & pAddNoLong
    query a
    If IsNull(RS.Fields(0).Value) Then
        b = 1
    Else
        b = Left(RS.Fields(0).Value, 5) + 1
    End If
    fNo = zerofill(b, 5) & pAddNo
    fTanggal = pServerDate
    GoEvent "NEW"
    fPemberiJasa.SetFocus
err:
End Sub

Private Sub iToolbar1_NextClick()
On Error Resume Next
Dim a As String
    a = iToolbar1.GetText
    Mid(a, 1) = zerofill(Left(a, 5) + 1, 5)
    GetResult a
End Sub

Private Sub iToolbar1_PrevClick()
On Error Resume Next
Dim a As String
    a = iToolbar1.GetText
    Mid(a, 1) = zerofill(Left(a, 5) - 1, 5)
    GetResult a
End Sub

Private Sub iToolbar1_BottomClick()
On Error Resume Next
    a = "select max(NoBukti) from t_Jasa where Tanggal>" & pAddNoLong
    query a
    If Not IsNull(RS.Fields(0).Value) Then
        GetResult RS.Fields(0).Value
    Else
        iToolbar1.SetNoData True
        ClearScreen
        GoEvent "SEE"
    End If
End Sub

Private Sub iToolbar1_SaveClick()
On Error GoTo err
    BeginTransaction
    HitungTotal
    If m_mode <> "NEW" Then
        a = "delete from t_Jasa where NoBukti='" & esc(fNo.Tag) & "'"
        ExecMe a
        a = "delete from t_JasaDetail where NoBukti='" & esc(fNo.Tag) & "'"
        ExecMe a
    End If
    a = "insert into t_Jasa(NamaPemberiJasa,NoBukti,Tanggal,Total,NoFakturJasa,MataUang) values('" & _
        fPemberiJasa & _
        "','" & fNo & _
        "'," & cD(fTanggal) & _
        "," & cNum(fTotal) & _
        ",'" & fNoFakturJasa & _
        "','" & fMataUang & "')"
    If ExecMe(a) = 0 Then GoTo err
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        a = "insert into t_JasaDetail(NamaPemberiJasa,NoBukti,Tanggal,Nilai, NamaJasa,DebetAcc,KetAcc, NoFakturJasa, MataUang) values('" & _
            fPemberiJasa & _
            "','" & fNo & _
            "'," & cD(fTanggal) & _
            "," & cNum(x(i, col1("Nilai").ColIndex)) & _
            ",'" & x(i, col1("Nama Jasa").ColIndex) & _
            "','" & x(i, col1("Debet Acc").ColIndex) & _
            "','" & x(i, col1("Ket Acc").ColIndex) & _
            "','" & fNoFakturJasa & _
            "','" & fMataUang & "')"
        If ExecMe(a) = 0 Then GoTo err
    Next
    CommitTransaction
    MsgBox "SUKSES"
    GetResult fNo
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
    iToolbar1.SetGagalSave
End Sub

Private Sub iToolbar1_TextEnter()
    GetResult iToolbar1.GetText
End Sub

Private Sub iToolbar1_TopClick()
On Error Resume Next
    a = "select min(NoBukti) from t_Jasa where Tanggal>" & pAddNoLong
    query a
    If Not IsNull(RS.Fields(0).Value) Then
        GetResult RS.Fields(0).Value
    Else
        iToolbar1.SetNoData True
        ClearScreen
        GoEvent "SEE"
    End If
End Sub

Private Sub iToolbar1_DeleteClick()
On Error GoTo err
    BeginTransaction
    a = "delete from t_Jasa where NoBukti='" & esc(fNo) & "'"
    If ExecMe(a) = 0 Then GoTo err
    a = "delete from t_JasaDetail where NoBukti='" & esc(fNo) & "'"
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    iToolbar1_BottomClick
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
    iToolbar1.SetGagalSave
End Sub

Private Sub iToolbar1_EditClick()
    m_mode = "EDIT"
    GoEvent
    TDBGrid1.SetFocus
    fNo.Tag = fNo
End Sub

Private Sub iToolbar1_ExitClick()
    Unload Me
End Sub

Private Sub HitungTotal()
On Error Resume Next
Dim a As Double
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        a = a + x(i, col1("Nilai").ColIndex)
    Next
    fTotal = cDecimal(a)
End Sub
Private Sub iToolbar1_ListClick()
    FormList.LoadMe "BELUM LUNAS@SUDAH LUNAS", _
        "select NoBukti, Tanggal, NamaPemberiJasa, Total from t_Jasa where Pelunasan<Total@select NoBukti, Tanggal, NamaPemberiJasa, Total from t_Jasa where Pelunasan>=Total", _
        "Tanggal@Pemberi Jasa", "Tanggal@NamaPemberiJasa", "1000@2500", "Date@String", _
        "No Bukti@Tanggal@Pemberi Jasa@Total", "1500@1500@2500@1500", "String@Date@String@Decimal", Me, "", 0
End Sub

Private Sub iToolbar1_CancelClick()
    GetResult iToolbar1.GetText
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If col1(ColIndex).Caption = "Nilai" Then HitungTotal
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    a = ActiveControl.Name
    If a <> "TDBGrid1" Then
        If IsNull(LastRow) Then TDBGrid1.Row = 0
        If x.UpperBound(1) > 0 Then TDBGrid1.SetFocus
    End If
End Sub




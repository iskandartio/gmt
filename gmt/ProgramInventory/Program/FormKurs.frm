VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormKurs 
   BackColor       =   &H00FFC0C0&
   Caption         =   "KURS"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Tag             =   "35"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAdd 
      Caption         =   "TAMBAH TAHUN"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin UsrText.IText fAkhir 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Top             =   360
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
      Left            =   120
      TabIndex        =   2
      Top             =   360
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
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tanggal"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Beli"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Jual"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nilai"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Kurs Pajak"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "updated"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=2000"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=1"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=1"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(3)._MinWidth=1"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(5)._MinWidth=93"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
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
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=54,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=46,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=15"
      _StyleDefs(54)  =   "Named:id=29:Normal"
      _StyleDefs(55)  =   ":id=29,.parent=0"
      _StyleDefs(56)  =   "Named:id=30:Heading"
      _StyleDefs(57)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(58)  =   ":id=30,.wraptext=-1"
      _StyleDefs(59)  =   "Named:id=31:Footing"
      _StyleDefs(60)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   "Named:id=32:Selected"
      _StyleDefs(62)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=33:Caption"
      _StyleDefs(64)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(65)  =   "Named:id=34:HighlightRow"
      _StyleDefs(66)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(67)  =   "Named:id=35:EvenRow"
      _StyleDefs(68)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=36:OddRow"
      _StyleDefs(70)  =   ":id=36,.parent=29"
      _StyleDefs(71)  =   "Named:id=39:RecordSelector"
      _StyleDefs(72)  =   ":id=39,.parent=30"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=29"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FormKurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim mControl As Boolean
Dim col1 As TrueOleDBGrid80.Columns

Private Sub cmdAdd_Click()
    InputTanggal
End Sub

Private Sub fAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub fAkhir_Validate(Cancel As Boolean)
    If cD(fAkhir) > cD(pServerDate) Or cD(fAkhir) = "A" Then fAkhir = pServerDate
    DoQuery
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub


Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    fAkhir = pServerDate
    fAwal = add_tanggal(fAkhir, -20)
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    col1("Tanggal").NumberFormat = "FormatText Event"
    col1("Tanggal").Locked = True
    col1("Beli").Tag = "Decimal"
    col1("Jual").Tag = "Decimal"
    col1("Nilai").Tag = "Decimal"
    col1("Nilai").Locked = True
    col1("Kurs Pajak").Tag = "Decimal"
    TDBGrid1.FetchRowStyle = True
    TDBGridSetVisible TDBGrid1, "updated"
    TDBGridLoad TDBGrid1
    DoQuery
End Sub

Private Sub DoQuery()
    Dim maxTanggal As Integer
    a = "select Tanggal, Beli, Jual, Nilai, KursPajak, 0  from m_kurs where Tanggal>=" & cD(fAwal) & " and Tanggal<=" & cD(fAkhir) & " order by Tanggal"
    query a
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left - 100
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If x(i, col1("updated").ColIndex) = 1 Then
            Beli = x(i, col1("Beli").ColIndex)
            Jual = x(i, col1("Jual").ColIndex)
            Nilai = x(i, col1("Nilai").ColIndex)
            KursPajak = x(i, col1("Kurs Pajak").ColIndex)
            If Nilai = "" Then Nilai = 0
            a = "update m_kurs set Beli=" & cNum(Beli) & ", Jual=" & cNum(Jual) & ", KursPajak=" & cNum(KursPajak) & ", Nilai=" & cNum(Nilai) & " where Tanggal=" & x(i, col1("Tanggal").ColIndex)
            b = ExecMe(a)
            If b = 0 Then
                a = "insert into m_kurs(Tanggal, Beli, Jual, Nilai, KursPajak) values(" & _
                    x(i, col1("Tanggal").ColIndex) & _
                    "," & cNum(Beli) & _
                    "," & cNum(Jual) & _
                    "," & cNum(Nilai) & _
                    "," & cNum(KursPajak) & ")"
                If ExecMe(a) = 0 Then GoTo err
            End If
        End If
    Next
    MsgBox "SUKSES"
    DoEvents
    DoQuery
err:
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo err
    col1("updated").Value = 1
    col1("Nilai").Value = (col1("Beli").Value - -col1("Jual").Value) / 2
    Exit Sub
err:
    col1("Nilai").Value = 0
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(Bookmark, col1("updated").ColIndex) = 1 Then RowStyle.BackColor = vbYellow
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    a = col1(ColIndex).Caption
    If a = "Tanggal" Then
        Value = cTanggal(Value)
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    TDBGridKeyDown TDBGrid1, KeyCode
    If KeyCode = 86 Then
        col1("updated").Value = 1
        col1("Nilai").Value = (col1("Beli").Value - -col1("Jual").Value) / 2
    End If
    Exit Sub
err:
    col1("Nilai").Value = 0
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub



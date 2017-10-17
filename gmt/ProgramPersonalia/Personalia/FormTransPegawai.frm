VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "UsrText.ocx"
Begin VB.Form FormTransKaryawan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "TRANSAKSI KARYAWAN"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "54"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fTunjangan 
      Caption         =   "TUNJANGAN"
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame FrameFilter 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   8535
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   3
         Left            =   6720
         TabIndex        =   16
         Tag             =   "Ket"
         Top             =   240
         Width           =   1335
      End
      Begin UsrText.IText fTanggalAwal 
         Height          =   270
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         Text            =   "__/__/____"
         DataType        =   6
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
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Tag             =   "m_Karyawan.NIK"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Tag             =   "m_Karyawan.Departemen"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Tag             =   "m_Karyawan.Nama"
         Top             =   240
         Width           =   1335
      End
      Begin UsrText.IText fTanggalAkhir 
         Height          =   270
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         Text            =   "__/__/____"
         DataType        =   6
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ket"
         Height          =   255
         Left            =   6720
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "NIK"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemen"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox fFind 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
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
      _StyleDefs(30)  =   "Named:id=29:Normal"
      _StyleDefs(31)  =   ":id=29,.parent=0"
      _StyleDefs(32)  =   "Named:id=30:Heading"
      _StyleDefs(33)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(34)  =   ":id=30,.wraptext=-1"
      _StyleDefs(35)  =   "Named:id=31:Footing"
      _StyleDefs(36)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(37)  =   "Named:id=32:Selected"
      _StyleDefs(38)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(39)  =   "Named:id=33:Caption"
      _StyleDefs(40)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(41)  =   "Named:id=34:HighlightRow"
      _StyleDefs(42)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(43)  =   "Named:id=35:EvenRow"
      _StyleDefs(44)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(45)  =   "Named:id=36:OddRow"
      _StyleDefs(46)  =   ":id=36,.parent=29"
      _StyleDefs(47)  =   "Named:id=39:RecordSelector"
      _StyleDefs(48)  =   ":id=39,.parent=30"
      _StyleDefs(49)  =   "Named:id=42:FilterBar"
      _StyleDefs(50)  =   ":id=42,.parent=29"
   End
   Begin VB.CommandButton fLembur 
      Caption         =   "LEMBUR"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton fAbsensi 
      Caption         =   "ABSENSI"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Find"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "FormTransKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDBObject.XArrayDB
Dim mTitle As String
Dim mCancelUpdate As Boolean
Dim col1 As TrueOleDBGrid80.Columns


Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub fFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim cCol As Integer
        cCol = TDBGrid1.Col
        m = TDBGrid1.Bookmark
        n = m + 1
        If n = x.UpperBound(1) + 1 Then n = 0
        Do While m <> n
            If InStr(1, x(n, cCol), fFind, vbTextCompare) <> 0 Then Exit Do
            n = n + 1
            If n = x.UpperBound(1) + 1 Then n = 0
        Loop
        If InStr(1, x(n, cCol), fFind, vbTextCompare) = 0 Then
            MsgBox "Not Found"
        End If
        TDBGrid1.Bookmark = n
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    TDBGrid1.FetchRowStyle = True
    FrameFilter.Visible = False
    TDBGrid1.AllowAddNew = True
    TDBGrid1.AllowDelete = True
End Sub


Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left * 2
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 200
End Sub


Private Sub fTanggalAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    col1("updated").Value = "1"
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = col1("NIK").ColIndex Then
        qKaryawan
    End If
End Sub

Private Sub qKaryawan()
On Error GoTo err
    s = "select * from m_Karyawan where NIK=" & cNum(col1("NIK").Value)
    query s
    col1("Nama").Value = rs!Nama
    col1("Departemen").Value = rs!Departemen
    Exit Sub
err:
    col1("NIK").Value = ""
    col1("Nama").Value = ""
    col1("Departemen").Value = ""
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(Bookmark, col1("updated").ColIndex) = "1" Then RowStyle.BackColor = vbYellow
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        fFind_KeyDown 13, 0
    ElseIf KeyCode = 112 Then
        HelpMe "Karyawan", Me, TDBGrid1.Left + 1000, TDBGrid1.Top + 1000
    Else
        TDBGridKeyDown TDBGrid1, KeyCode
    End If
End Sub

Sub FormHelpKeyDown(ByVal tRetVal As Variant)
    col1("NIK").Value = tRetVal
    qKaryawan
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_MouseDown(Button As Integer, Shift As Integer, tX As Single, tY As Single)
On Error Resume Next
    a = ActiveControl.Name
    If a <> "TDBGrid1" Then
        If IsNull(LastRow) Then TDBGrid1.Row = 0
        If x.UpperBound(1) > 0 Then TDBGrid1.SetFocus
    End If
End Sub

Private Sub Texts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DoQuery
    End If
End Sub


Private Sub fAbsensi_Click()
    mTitle = "ABSENSI"
    DoQuery
End Sub

Private Sub DoQuery()
Dim MyFilter As String
    TDBGrid1.Bookmark = 0
    While col1.Count <> 0
        col1.Remove 0
    Wend
    MyFilter = ""
    For i = 0 To Texts.Count - 1
        If Trim(Texts(i)) <> "" Then MyFilter = MyFilter & " and " & Texts(i).Tag & " like '%" & Texts(i) & "%'"
    Next
    If fTanggalAwal.Text <> "__/__/____" Then MyFilter = MyFilter & " and Tanggal>=" & cD(fTanggalAwal.Text)
    If fTanggalAkhir.Text <> "__/__/____" Then MyFilter = MyFilter & " and Tanggal<=" & cD(fTanggalAkhir.Text)
    If mTitle = "ABSENSI" Then
        s = "select 0 as updated, tAuto as pKey, Tanggal, m_Karyawan.NIK, m_Karyawan.Nama, m_Karyawan.Departemen, st, ket from t_Absen left join m_Karyawan on t_Absen.NIK=m_Karyawan.NIK where 1=1" & MyFilter & " order by m_Karyawan.NIK, Tanggal"
    ElseIf mTitle = "LEMBUR" Then
        s = "select 0 as updated, tAuto as pKey, Tanggal, m_Karyawan.NIK, m_Karyawan.Nama, m_Karyawan.Departemen, Jam, JamAsli, Ket from t_Lembur left join m_Karyawan on t_Lembur.NIK=m_Karyawan.NIK where 1=1" & MyFilter & " order by m_Karyawan.NIK, Tanggal"
    ElseIf mTitle = "TUNJANGAN" Then
        s = "select 0 as updated, tAuto as pKey, Tanggal, m_Karyawan.NIK, m_Karyawan.Nama, m_Karyawan.Departemen, HariShift, TJabatan, TKhusus, TLain, Ket from t_Tunjangan left join m_Karyawan on t_Tunjangan.NIK=m_Karyawan.NIK where 1=1" & MyFilter & " order by m_Karyawan.NIK, Tanggal"
    End If
    query s
    For i = 0 To rs.Fields.Count - 1
        col1.Add(i).Caption = rs.Fields(i).Name
        If i > 0 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1("updated").Visible = False
    col1("NIK").NumberFormat = "000000"
    col1("pKey").Visible = False
    col1("Tanggal").Tag = "DateLong"
    If mTitle = "LEMBUR" Then
        col1("Jam").Tag = "Decimal"
        col1("JamAsli").Tag = "Decimal"
    ElseIf mTitle = "TUNJANGAN" Then
        col1("HariShift").Tag = "Integer"
        col1("TKhusus").Tag = "Decimal"
        col1("TLain").Tag = "Decimal"
        col1("TJabatan").Tag = "Decimal"
    End If
    TDBGridSetLock TDBGrid1, "Nama@Departemen", True
    TDBGridLoad TDBGrid1
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not rs.EOF Then x.LoadRows rs.GetRows
    For i = 0 To x.UpperBound(1)
        x(i, col1("Tanggal").ColIndex) = cTanggal(x(i, col1("Tanggal").ColIndex), True)
    Next
    TDBGrid1.ReBind
    
    FrameFilter.Visible = True
End Sub

Private Sub fLembur_Click()
    mTitle = "LEMBUR"
    DoQuery
End Sub

Private Sub fTunjangan_Click()
    mTitle = "TUNJANGAN"
    DoQuery
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If CLng(x(i, 0)) = 1 Then
            c = x(i, col1("pKey").ColIndex)
            If mTitle = "ABSENSI" Then
                If cNum(c) = 0 Then
                    s = "insert into t_Absen(NIK, Tanggal, St, Ket) values(" & _
                        cNum(x(i, col1("NIK").ColIndex)) & _
                        "," & cD(x(i, col1("Tanggal").ColIndex), True) & _
                        ",'" & esc(x(i, col1("St").ColIndex)) & _
                        "','" & esc(x(i, col1("Ket").ColIndex)) & "')"
                    If ExecMe(s) < 1 Then GoTo err
                Else
                    s = "update t_Absen set NIK=" & cNum(x(i, col1("NIK").ColIndex)) & _
                        ", Tanggal=" & cD(x(i, col1("Tanggal").ColIndex), True) & _
                        ", St='" & esc(x(i, col1("St").ColIndex)) & _
                        "', Ket='" & esc(x(i, col1("Ket").ColIndex)) & "' where tAuto=" & c
                    If ExecMe(s) < 1 Then GoTo err
                End If
            ElseIf mTitle = "LEMBUR" Then
                If cNum(c) = 0 Then
                    s = "insert into t_Lembur(NIK, Tanggal, Jam, JamAsli, Ket) values(" & _
                        cNum(x(i, col1("NIK").ColIndex)) & _
                        "," & cD(x(i, col1("Tanggal").ColIndex), True) & _
                        "," & cNum(x(i, col1("Jam").ColIndex)) & _
                        "," & cNum(x(i, col1("JamAsli").ColIndex)) & _
                        ",'" & esc(x(i, col1("Ket").ColIndex)) & "')"
                    If ExecMe(s) < 1 Then GoTo err
                Else
                    s = "update t_Lembur set NIK=" & cNum(x(i, col1("NIK").ColIndex)) & _
                        ", Tanggal=" & cD(x(i, col1("Tanggal").ColIndex), True) & _
                        ", Jam=" & cNum(x(i, col1("Jam").ColIndex)) & _
                        ", JamAsli=" & cNum(x(i, col1("JamAsli").ColIndex)) & _
                        ", Ket='" & esc(x(i, col1("Ket").ColIndex)) & "' where tAuto=" & c
                    If ExecMe(s) < 1 Then GoTo err
                End If
            ElseIf mTitle = "TUNJANGAN" Then
                If CDbl(x(i, col1("HariShift").ColIndex)) > 1 Then
                    MsgBox "Hari Shift diisi 1 atau 0"
                    GoTo err
                End If
                If cNum(c) = 0 Then
                    s = "insert into t_Tunjangan(NIK, Tanggal, HariShift, TJabatan, TKhusus, TLain, Ket) values(" & _
                        cNum(x(i, col1("NIK").ColIndex)) & _
                        "," & cD(x(i, col1("Tanggal").ColIndex), True) & _
                        "," & cNum(x(i, col1("HariShift").ColIndex)) & _
                        "," & cNum(x(i, col1("TJabatan").ColIndex)) & _
                        "," & cNum(x(i, col1("TKhusus").ColIndex)) & _
                        "," & cNum(x(i, col1("TLain").ColIndex)) & _
                        ",'" & esc(x(i, col1("Ket").ColIndex)) & "')"
                    If ExecMe(s) < 1 Then GoTo err
                Else
                    s = "update t_Tunjangan set NIK=" & cNum(x(i, col1("NIK").ColIndex)) & _
                        ", Tanggal=" & cD(x(i, col1("Tanggal").ColIndex), True) & _
                        ", HariShift=" & cNum(x(i, col1("HariShift").ColIndex)) & _
                        ", TJabatan=" & cNum(x(i, col1("TJabatan").ColIndex)) & _
                        ", TKhusus=" & cNum(x(i, col1("TKhusus").ColIndex)) & _
                        ", TLain=" & cNum(x(i, col1("TLain").ColIndex)) & _
                        ", Ket='" & esc(x(i, col1("Ket").ColIndex)) & "' where tAuto=" & c
                    If ExecMe(s) < 1 Then GoTo err
                End If
            End If
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


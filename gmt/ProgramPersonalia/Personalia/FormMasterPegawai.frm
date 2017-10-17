VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FormMasterKaryawan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MASTER KARYAWAN"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "53"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   120
      TabIndex        =   11
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
      Begin VB.CommandButton fTambah 
         Caption         =   "+"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Tag             =   "Departemen"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Tag             =   "Nama"
         Top             =   240
         Width           =   1335
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
   Begin VB.CommandButton fDepartemen 
      Caption         =   "DEPARTEMEN"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton fKaryawan 
      Caption         =   "KARYAWAN"
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
Attribute VB_Name = "FormMasterKaryawan"
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
    s = "update m_Karyawan set NIK=0 where 1=0"
    If ExecMe(s) = -1 Then
        s = "create table m_Karyawan(NIK long PRIMARY KEY, Nama text(255), Departemen text(50))"
        ExecMe s
    End If
    s = "update t_TransKaryawan set NIK=0 where 1=0"
    If ExecMe(s) = -1 Then
        s = ""
        For i = 1 To 31
            s = s & ",a" & i & " text(1), b" & i & " int, c" & i & " text(200) "
        Next

        s = "create table t_TransKaryawan(NIK long, Periode int" & s & ", constraint pKeyNIK primary key(NIK, Periode))"
        ExecMe s
    End If
    Caption = Caption & "---" & pTipe
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    TDBGrid1.FetchRowStyle = True
    FrameFilter.Visible = False
End Sub

Sub DoQuery()
On Error Resume Next
    If mTitle = "KARYAWAN" Then
        fKaryawan_Click
    ElseIf mTitle = "DEPARTEMEN" Then
        fDepartemen_Click
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left * 2
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 200
End Sub

Private Sub fKaryawan_Click()
    TDBGrid1.Bookmark = 0
    While col1.Count <> 0
        col1.Remove 0
    Wend
    MyFilter = ""
    For i = 0 To Texts.Count - 1
        If Trim(Texts(i)) <> "" Then MyFilter = MyFilter & " and " & Texts(i).Tag & " like '%" & Texts(i) & "%'"
    Next
    a = "select 0 as updated, NIK as pKey, NIK, Nama, Departemen from m_Karyawan where 1=1 " & MyFilter & " order by NIK"
    query a
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 0 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1("updated").Visible = False
    col1("NIK").NumberFormat = "000000"
    col1("pKey").Visible = False
    TDBGridLoad TDBGrid1
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.ReBind
    mTitle = "KARYAWAN"
    FrameFilter.Visible = True
End Sub

Private Sub fDepartemen_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, Departemen as pKey, Departemen from m_Karyawan order by Departemen"
    query a
    While col1.Count <> 0
        col1.Remove 0
    Wend
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 1 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1(0).Visible = False
    col1(1).Visible = False
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.ReBind
    mTitle = "DEPARTEMEN"
    FrameFilter.Visible = False
End Sub

Private Sub fTambah_Click()
On Error GoTo err
Dim tNIK As Long
    TDBGrid1.Update
    a = "select Max(NIK) from m_Karyawan"
    query a
    If Not IsNull(RS.Fields(0).Value) Then tNIK = RS.Fields(0).Value + 1 Else tNIK = 1
    a = "insert into m_Karyawan(NIK) values(" & tNIK & ")"
    If ExecMe(a) = 0 Then GoTo err
    x.AppendRows
    x(x.UpperBound(1), col1("NIK").ColIndex) = tNIK
    x(x.UpperBound(1), col1("pKey").ColIndex) = tNIK
    x(x.UpperBound(1), col1("updated").ColIndex) = 1
    TDBGrid1.ReBind
    TDBGrid1.MoveLast
    TDBGrid1.SetFocus
err:
End Sub

Private Sub fUpdate_Click()
'On Error GoTo err

    BeginTransaction
    TDBGrid1.Update
    For i = 0 To x.UpperBound(1)
        If CLng(x(i, 0)) = 1 Then
            c = x(i, col1("pKey").ColIndex)
            If mTitle = "KARYAWAN" Then
                s = "update m_Karyawan set NIK='" & x(i, col1("NIK").ColIndex) & "', Nama='" & x(i, col1("Nama").ColIndex) & "', Departemen='" & x(i, col1("Departemen").ColIndex) & "'" & _
                    " where NIK=" & c
                If ExecMe(s) = 0 Then GoTo err
            ElseIf mTitle = "DEPARTEMEN" Then
                
                s = "update m_Karyawan set Departemen='" & x(i, col1("Departemen").ColIndex) & "' where Departemen='" & c & "'"
                If ExecMe(s) = 0 Then GoTo err
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


Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    col1("updated").Value = "1"
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(Bookmark, col1("updated").ColIndex) = "1" Then RowStyle.BackColor = vbYellow
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        fFind_KeyDown 13, 0
    Else
        TDBGridKeyDown TDBGrid1, KeyCode
    End If
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



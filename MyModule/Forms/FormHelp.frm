VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormHelp 
   Caption         =   "HELP"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   13160660
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Named:id=29:Normal"
      _StyleDefs(39)  =   ":id=29,.parent=0"
      _StyleDefs(40)  =   "Named:id=30:Heading"
      _StyleDefs(41)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=30,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=31:Footing"
      _StyleDefs(44)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=32:Selected"
      _StyleDefs(46)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=33:Caption"
      _StyleDefs(48)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(49)  =   "Named:id=34:HighlightRow"
      _StyleDefs(50)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(51)  =   "Named:id=35:EvenRow"
      _StyleDefs(52)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=36:OddRow"
      _StyleDefs(54)  =   ":id=36,.parent=29"
      _StyleDefs(55)  =   "Named:id=39:RecordSelector"
      _StyleDefs(56)  =   ":id=39,.parent=30"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=29"
   End
   Begin VB.TextBox fCari 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4200
      Top             =   720
   End
End
Attribute VB_Name = "FormHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim f As Form
Dim mRetCol As String
Dim fFind As String
Dim mUnload As Boolean


Sub LoadMe(ByVal tSQL As String, ByVal tHeaders As String, ByVal tWidths As String, ByVal tTypes As String, ByVal tf As Form, Optional ByVal tLeft As Single = 0, Optional ByVal tTop As Single = 0, Optional ByVal tRetCol As String = "0", Optional ByVal tRowHeight As Single = 0)
On Error GoTo err
    query tSQL
With TDBGrid1
    While .Columns.Count > 0
        .Columns.Remove 0
    Wend
Dim a() As String
Dim b() As String
Dim w As Double
    a = Split(tHeaders, "@")
    b = Split(tWidths, "@")
    c = Split(tTypes, "@")
    For i = 0 To UBound(a)
        .Columns.Add(i).Caption = a(i)
        .Columns(i).Visible = Left(a(i), 1) <> "*"
        If i <= UBound(b) Then .Columns(i).Width = b(i) Else .Columns(i).Width = "3000"
        If i <= UBound(c) Then .Columns(i).Tag = c(i) Else .Columns(i).Tag = "String"
        w = w + b(i)
    Next
    Set .Array = x
    TDBGridLoad TDBGrid1
    If tRowHeight > 0 Then TDBGrid1.RowHeight = tRowHeight
    If RS.RecordCount > 0 Then x.LoadRows RS.GetRows
    Width = w + 200
    For j = 0 To x.UpperBound(2)
        If .Columns(j).Tag = "Date" Then
            For i = 0 To x.UpperBound(1)
                x(i, j) = cTanggal2(x(i, j))
            Next
        End If
    Next
    .Rebind
End With
    Set f = tf
    mRetCol = tRetCol
    fFind = f.ActiveControl.Text
    Cari
    If f.ActiveControl.Container.Name = f.Name Then
        t = f.Top
        l = f.Left
    Else
        l = f.Left + f.ActiveControl.Container.Left
        t = f.Top + f.ActiveControl.Container.Top
    End If
    If tLeft = 0 Then Left = l + f.ActiveControl.Left Else Left = tLeft
    If tTop = 0 Then Top = t + f.ActiveControl.Top + f.ActiveControl.Height + 270 Else Top = tTop
    Timer1.Enabled = False
    Show vbModal, f
err:
End Sub


Private Sub Form_Load()
    TDBGrid1.Font.Size = 10
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth
    TDBGrid1.Height = ScaleHeight
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then
        b = Split(mRetCol, "@")
        For i = 0 To UBound(b)
            a = a & "@" & TDBGrid1.Columns(CLng(b(i))).Value
        Next
        f.FormHelpKeyDown Mid(a, 2)
        mUnload = True
        fCari.SetFocus
    End If
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii > 47 And KeyAscii < 123 Then
        fFind = fFind & Chr(KeyAscii)
    End If
    fCari = fFind
    Timer1.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub TDBGrid1_LostFocus()
    If mUnload Then Unload Me
End Sub

Private Sub Timer1_Timer()
    Cari
    Timer1.Enabled = False
End Sub

Private Sub Cari()
    If fFind = "" Then Exit Sub
Dim cCol As Integer
    cCol = TDBGrid1.Col
    n = TDBGrid1.Bookmark
    m = n + 1
    If m = x.UpperBound(1) + 1 Then m = 0
    Do While m <> n
        If InStr(1, x(m, cCol), fFind, vbTextCompare) = 1 Then Exit Do
        m = m + 1
        If m = x.UpperBound(1) + 1 Then m = 0
    Loop
    TDBGrid1.Bookmark = m
    While TDBGrid1.SelBookmarks.Count <> 0
        TDBGrid1.SelBookmarks.Remove 0
    Wend
    fFind = ""
    TDBGrid1.SelBookmarks.Add m
    Timer1.Enabled = False
End Sub



VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.UserControl ITrueCombo 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   4905
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ITrueCombo.ctx":0000
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8070
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
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14215660
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(23)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(44)  =   "Named:id=29:Normal"
      _StyleDefs(45)  =   ":id=29,.parent=0"
      _StyleDefs(46)  =   "Named:id=30:Heading"
      _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=30,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=31:Footing"
      _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=32:Selected"
      _StyleDefs(52)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=33:Caption"
      _StyleDefs(54)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(55)  =   "Named:id=34:HighlightRow"
      _StyleDefs(56)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(57)  =   "Named:id=35:EvenRow"
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "ITrueCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim col1 As TrueOleDBGrid80.Columns
Dim z As New XArrayDBObject.XArrayDB
Dim Loaded As Boolean
Dim mCancel As Boolean
Dim mRun As Boolean
Dim mListIndex As Integer
Dim mType() As String
Dim mname() As String
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
'Event Declarations:
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."

Function ListCount() As Long
    ListCount = z.UpperBound(1) + 1
    
End Function

Sub Clear()
    z.ReDim 0, z.UpperBound(1), 0, z.UpperBound(2)
    z.DeleteRows 0
    TDBGrid1.ReBind
End Sub

Sub SetListIndex(ByVal i As Long)
    mListIndex = i
    While TDBGrid1.SelBookmarks.Count > 0
        TDBGrid1.SelBookmarks.Remove 0
    Wend
    TDBGrid1.SelBookmarks.Add i
    TDBGrid1.Bookmark = i
End Sub

Function zz(ByVal i As Long, ByVal tColumnName As String) As Variant
    j = col1(tColumnName).ColIndex
    If j > UBound(mname) Then zz = "" Else zz = z(i, j)
End Function

Sub SetDropDownHeight(ByVal tHeight As Single)
    TDBGrid1.Height = tHeight
End Sub

Sub Cancel()
    mCancel = True
End Sub

Function Validate() As Boolean
    Validate = True
    If mListIndex = -1 Then
        Validate = False
        SetFocus
    End If
End Function

Private Sub TDBGrid1_Click()
    Text1.SetFocus
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
On Error GoTo err
    If mType(ColIndex) = "Date" Then
        Value = cTanggal(Value)
    ElseIf mType(ColIndex) = "Currency" Then
        Value = cDecimal(Value)
    ElseIf Mid(mType(ColIndex), 2) = "Zerofill" Then
        Value = zerofill(Value, CByte(Left(mType(ColIndex), 1)))
    End If
    Exit Sub
err:
    Value = ""
End Sub

Function cDecimal(ByVal Value As Double)
On Error Resume Next
    cDecimal = Format(Value, "###,###,##0.00")
End Function

Function cTanggal(ByVal a As Long)
    b = zerofill(a, 6)
    cTanggal = Right(b, 2) & "/" & Mid(b, 3, 2) & "/" & Left(b, 2)
End Function

Private Function zerofill(ByVal tStr As String, ByVal num As Integer) As String
On Error Resume Next
    zerofill = String(num, "0")
    Mid(zerofill, num + 1 - Len(tStr)) = tStr
End Function

Private Sub change()
    If TDBGrid1.SelBookmarks.Count <> 0 Then
        Text1 = TDBGrid1.Columns(0).Value
    End If
End Sub

Sub InvisibleMode()
    TDBGrid1.Visible = False
    Height = Text1.Height
End Sub

Sub VisibleMode()
    TDBGrid1.Visible = True
    Height = Text1.Height + TDBGrid1.Height
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If ActiveControl.Name = "Text1" Then Exit Sub
    Text1.Text = TDBGrid1.Columns(0).Value
    Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Function ListIndex()
    ListIndex = mListIndex
End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
    mRun = True
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    f = TDBGrid1.Bookmark
    If KeyCode = 27 Then
        InvisibleMode
    ElseIf KeyCode = 33 Then
        f = f - 10
    ElseIf KeyCode = 34 Then
        f = f + 10
    ElseIf KeyCode = 40 Then
        If TDBGrid1.Visible Then
            f = f + 1
        Else
            VisibleMode
            TDBGrid1.ZOrder 0
            Text1.SetFocus
        End If
    ElseIf KeyCode = 38 Then
        f = f - 1
    ElseIf KeyCode = 13 Then
        change
        SendKeys Chr(9)
    End If
    If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 33 Or KeyCode = 34 Then
        KeyCode = 0
        If f < 0 Then f = 0
        If f > z.UpperBound(1) Then f = z.UpperBound(1)
        If TDBGrid1.SelBookmarks.Count = 1 Then TDBGrid1.SelBookmarks.Remove 0
        TDBGrid1.SelBookmarks.Add f
        TDBGrid1.Bookmark = f
        mListIndex = f
        change
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
    End If
End Sub

Sub FindIndex(Optional ByVal tString As String = "", Optional ByVal tIndex As Long = 0)
On Error GoTo err
    If tString <> "" Then a = tString Else a = Text1
    h = z.UpperBound(1)
    l = 0
    mListIndex = -1
    f = -1
    If a = "" Then GoTo err
    If mType(tIndex) = "Integer" Then
        a = CLng(a)
        Do
            If l > h Then Exit Do
            m = (l + h) \ 2
            If z(m, tIndex) = a Then
                f = m
                h = m - 1
            ElseIf z(m, tIndex) < a Then
                l = m + 1
            Else
                h = m - 1
            End If
        Loop
    Else
        a1 = Len(a)
        Do
            If l > h Then Exit Do
            m = (l + h) \ 2
            v = StrComp(Left(z(m, tIndex), a1), a, vbTextCompare)
            If v = 0 Then
                f = m
                h = m - 1
            ElseIf v = -1 Then
                l = m + 1
            Else
                h = m - 1
            End If
        Loop
    End If
    If TDBGrid1.SelBookmarks.Count = 1 Then TDBGrid1.SelBookmarks.Remove 0
    If f <> -1 Then
        TDBGrid1.SelBookmarks.Add f
        TDBGrid1.Bookmark = f
        mListIndex = f
    Else
        TDBGrid1.Bookmark = -1
    End If
    Exit Sub
err:
    TDBGrid1.Bookmark = -1
    If TDBGrid1.SelBookmarks.Count = 1 Then TDBGrid1.SelBookmarks.Remove 0
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim a As String
    RaiseEvent KeyUp(KeyCode, Shift)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    If KeyCode = 0 Or KeyCode = 40 Or KeyCode = 38 Or KeyCode = 33 Or KeyCode = 34 Then Exit Sub
    FindIndex
End Sub

Private Sub Text1_LostFocus()
    If Not ActiveControl Is Nothing Then
        Exit Sub
    Else
        InvisibleMode
    End If
End Sub

Sub SetDB(rs1() As Variant)
    z.LoadRows rs1
    TDBGrid1.ReBind
    mListIndex = -1
    Text1_KeyUp 1, 0
    Loaded = True
End Sub

Sub SetType(ByVal tType As String)
    Dim a() As String
    a = Split(tType, "@")
    ReDim mType(UBound(a))
    For i = 0 To UBound(a)
        mType(i) = a(i)
        If a(i) = "Integer" Then
            TDBGrid1.Columns(i).Alignment = dbgRight
        ElseIf a(i) = "Date" Or Mid(a(i), 2) = "Zerofill" Then
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
        ElseIf a(i) = "Currency" Then
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
            TDBGrid1.Columns(i).Alignment = dbgRight
        ElseIf a(i) = "Decimal" Then
            TDBGrid1.Columns(i).NumberFormat = "Standard"
            TDBGrid1.Columns(i).Alignment = dbgRight
        End If
    Next
End Sub

Sub SetHeader(ByVal tHeader As String)
    Dim a() As String
    a = Split(tHeader, "@")
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    ReDim mname(UBound(a))
    For i = 0 To UBound(a)
        If Left(a(i), 1) = "*" Then
            TDBGrid1.Columns.Add(i).Caption = Trim(Mid(a(i), 2))
        Else
            TDBGrid1.Columns.Add(i).Caption = Trim(a(i))
            TDBGrid1.Columns(i).Visible = True
        End If
        mname(i) = TDBGrid1.Columns(i).Caption
        TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    Next
    z.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    Set TDBGrid1.Array = z
    TDBGrid1.ReBind
    mRun = True
End Sub

Sub SetWidth(ByVal tWidth As String)
On Error GoTo err
    Dim a() As String
    a = Split(tWidth, "@")
    For i = 0 To UBound(a)
        TDBGrid1.Columns(i).Width = CLng(a(i))
        cc = cc + TDBGrid1.Columns(i).Width
    Next
    Width = cc + 700
    TDBGrid1.Width = Width
    Exit Sub
err:
End Sub

Private Sub UserControl_GotFocus()
    If Not Enabled Then SendKeys Chr(9)
End Sub

Private Sub UserControl_Initialize()
    Set col1 = TDBGrid1.Columns
End Sub

Private Sub UserControl_InitProperties()
    TDBGrid1.Left = Text1.Left
End Sub

Private Sub UserControl_Resize()
    If Not mRun Then
        Text1.Width = Width
        FontSize = Text1.FontSize
        Text1.Height = TextHeight("I") * 1.4
        Height = Text1.Height
    End If
End Sub

Public Property Get GetData(ByVal tColumnName As Variant) As Variant
    GetData = ""
    If mListIndex = -1 Then Exit Property
    GetData = TDBGrid1.Columns(tColumnName).Value
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Text.VB_UserMemId = 0
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Text1.Text = PropBag.ReadProperty("Text", "")
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    TDBGrid1.ForeColor = PropBag.ReadProperty("ForeColor2", &H80000008)
    Set TDBGrid1.Font = PropBag.ReadProperty("Font2", Ambient.Font)
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    Height = PropBag.ReadProperty("Height")
    Text1.Height = Height
    TDBGrid1.Top = Text1.Top + Text1.Height
    TDBGrid1.ExtendRightColumn = True
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", Text1.Text, "")
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ForeColor2", TDBGrid1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Font2", TDBGrid1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
    'Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
    FontSize = Text1.FontSize
    Height = 1.4 * TextHeight("I")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Font"
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,ForeColor
Public Property Get ForeColor2() As OLE_COLOR
Attribute ForeColor2.VB_ProcData.VB_Invoke_Property = ";Font"
    ForeColor2 = TDBGrid1.ForeColor
End Property

Public Property Let ForeColor2(ByVal New_ForeColor2 As OLE_COLOR)
    TDBGrid1.ForeColor() = New_ForeColor2
    PropertyChanged "ForeColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,Font
Public Property Get Font2() As Font
Attribute Font2.VB_ProcData.VB_Invoke_Property = ";Font"
    Set Font2 = TDBGrid1.Font
End Property

Public Property Set Font2(ByVal New_Font2 As Font)
    Set TDBGrid1.Font = New_Font2
    PropertyChanged "Font2"
    FontSize = TDBGrid1.Font.Size
    TDBGrid1.RowHeight = TextHeight("I") * 1.1
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub


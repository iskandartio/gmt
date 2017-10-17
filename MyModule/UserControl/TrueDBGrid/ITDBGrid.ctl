VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.UserControl ITDBGrid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ITDBGrid.ctx":0000
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5530
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "ITDBGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim y As New XArrayDB
Dim mCancel As Boolean
Dim mLocal As Boolean
'Event Declarations:
Event BeforeDelete(Cancel As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,BeforeDelete
Attribute BeforeDelete.VB_Description = "Occurs before record deletion from grid"
Event RowColChange(LastRow As Variant, ByVal LastCol As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,RowColChange
Attribute RowColChange.VB_Description = "Occurs after a different cell becomes current in the grid"
Event BeforeInsert(Cancel As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,BeforeInsert
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,KeyPress
Event BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,BeforeColEdit
Event BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,BeforeColUpdate
Event AfterColEdit(ByVal ColIndex As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,AfterColEdit
Event AfterColUpdate(ByVal ColIndex As Integer) 'MappingInfo=TDBGrid1,TDBGrid1,-1,AfterColUpdate
'Default Property Values:
Const m_def_MaxRow = 0
'Property Variables:
Dim m_MaxRow As Byte
Dim mType() As String
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Sub SetDim(ByVal tX As Long, tY As Long)
    y.ReDim 0, tX, 0, tY
End Sub

Sub SetType(ByVal tTypes As String)
    mType = Split(tTypes, ".")
    For i = 0 To UBound(mType)
        If mType(i) = "Integer" Then
            TDBGrid1.Columns(i).Alignment = dbgRight
        ElseIf mType(i) = "Decimal" Then
            TDBGrid1.Columns(i).Alignment = dbgRight
            TDBGrid1.Columns(i).NumberFormat = "Standard"
        ElseIf mType(i) = "NoCents" Then
            TDBGrid1.Columns(i).Alignment = dbgRight
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
        ElseIf Left(mType(i), 8) = "Zerofill" Then
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
        ElseIf mType(i) = "YesNo" Then
            TDBGrid1.Columns(i).Alignment = dbgCenter
        ElseIf mType(i) = "Date" Then
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
        End If
    Next
End Sub

Sub SetProp(ByVal tFields As String, ByVal tProperty As String, tVal As Variant)
On Error Resume Next
    Dim a() As String
    a = Split(tFields, ".")
    For i = 0 To UBound(a)
        CallByName TDBGrid1.Columns(a(i)), tProperty, VbLet, tVal
    Next
End Sub

Sub SetValueItems(ByVal tField As String, tVal() As Variant)
    TDBGrid1.Columns(tField).ValueItems.Presentation = dbgComboBox
    TDBGrid1.Columns(tField).ValueItems.Clear
    Dim c As New TrueOleDBGrid80.ValueItem
    For i = 0 To UBound(tVal)
        c.Value = tVal(i)
        TDBGrid1.Columns(tField).ValueItems.Add c
    Next
End Sub

Sub InsertRow(ByVal tRow As Long)
    TDBGrid1.Update
    If y.UpperBound(1) + 1 < m_MaxRow Then
        y.InsertRows tRow
    End If
End Sub

Sub AppendRow(ByVal tRow As Long)
    TDBGrid1.Update
    If y.UpperBound(1) < m_MaxRow Then
        If tRow = y.UpperBound(1) Then
            y.AppendRows
        Else
            y.InsertRows tRow + 1
        End If
    End If
End Sub

Sub Clear()
    y.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    y.DeleteRows 0
    Set TDBGrid1.Array = y
    TDBGrid1.ReBind
End Sub

Sub DeleteRows(ByVal tRow As Long)
    y.DeleteRows tRow
End Sub
Sub SetHeader(ByVal tHeader As String)
    a = Split(tHeader, ".")
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    For i = 0 To UBound(a)
        b = Trim(a(i))
        If Left(b, 1) = "*" Then
            TDBGrid1.Columns.Add(i).Caption = Mid(b, 2)
            TDBGrid1.Columns(i).AllowSizing = False
        Else
            TDBGrid1.Columns.Add(i).Visible = True
            TDBGrid1.Columns(i).Caption = b
            TDBGrid1.Columns(i).HeadAlignment = dbgCenter
        End If
    Next
    Clear
End Sub

Sub SetWidth(ByVal tWidth As String)
On Error GoTo err
    a = Split(tWidth, ".")
    For i = 0 To UBound(a)
        TDBGrid1.Columns(i).Width = CLng(a(i))
    Next
    Exit Sub
err:
End Sub

Sub AutoSize()
    For i = 0 To TDBGrid1.Columns.Count - 1
        TDBGrid1.Columns(i).AutoSize
    Next
End Sub

Sub SetDB(rs1() As Variant)
    y.LoadRows rs1
    TDBGrid1.ReBind
End Sub

Function GetColIndex(ByVal tStr As String)
    GetColIndex = -1
    GetColIndex = TDBGrid1.Columns(tStr).ColIndex
End Function

Public Property Get z(ByVal tRow As Long, ByVal tCol As Variant) As Variant
Attribute z.VB_UserMemId = 0
    If TypeName(tCol) = "String" Then
        tCol = GetColIndex(tCol)
    End If
    z = y(tRow, tCol)
End Property

Public Property Let z(ByVal tRow As Long, ByVal tCol As Variant, ByVal tVal As Variant)
    If TypeName(tCol) = "String" Then
        tCol = GetColIndex(tCol)
    End If
    y(tRow, tCol) = tVal
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,Columns
Public Property Get Columns() As Columns
    Set Columns = TDBGrid1.Columns
End Property

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    RaiseEvent BeforeColEdit(ColIndex, KeyAscii, Cancel)
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    RaiseEvent BeforeColUpdate(ColIndex, OldValue, Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,AllowDelete
Public Property Get AllowDelete() As Boolean
    AllowDelete = TDBGrid1.AllowDelete
End Property

Public Property Let AllowDelete(ByVal New_AllowDelete As Boolean)
    TDBGrid1.AllowDelete() = New_AllowDelete
    PropertyChanged "AllowDelete"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,AllowArrows
Public Property Get AllowArrows() As Boolean
    AllowArrows = TDBGrid1.AllowArrows
End Property

Public Property Let AllowArrows(ByVal New_AllowArrows As Boolean)
    TDBGrid1.AllowArrows() = New_AllowArrows
    PropertyChanged "AllowArrows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,AllowAddNew
Public Property Get AllowAddNew() As Boolean
    AllowAddNew = TDBGrid1.AllowAddNew
End Property

Public Property Let AllowAddNew(ByVal New_AllowAddNew As Boolean)
    TDBGrid1.AllowAddNew() = New_AllowAddNew
    PropertyChanged "AllowAddNew"
End Property

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    RaiseEvent AfterColEdit(ColIndex)
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    RaiseEvent AfterColUpdate(ColIndex)
    If mCancel = True Then
        mCancel = False
        Exit Sub
    End If

End Sub


Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
On Error Resume Next
    If Left(mType(ColIndex), 8) = "Zerofill" Then
        L = CLng(Mid(mType(ColIndex), 9))
        While Len(Value) < L
            Value = "0" & Value
        Wend
    ElseIf mType(ColIndex) = "NoCents" Then
        Value = cDecimal(Value)
    ElseIf mType(ColIndex) = "Date" Then
        Value = cTanggal(Value)
    End If
End Sub

Private Function cTanggal(ByVal a As Long)
On Error GoTo err
    If a = 0 Then GoTo err
    cTanggal = Right(a, 2) & "/" & Mid(a, 2, 2) & "/0" & Left(a, 1)
    Exit Function
err:
    cTanggal = "__/__/__"
End Function

Function cDecimal(ByVal Value As Double)
On Error Resume Next
    cDecimal = Format(Value, "###,###,##0.00")
End Function

Sub SetNumberFormat(ByVal ColIndexes As String, ByVal Format As String)
    Dim a() As String
    a = Split(ColIndexes, ".")
    If Format <> "No Cents" Then
        For i = 0 To UBound(a)
            TDBGrid1.Columns(CByte(i)).NumberFormat = Format
        Next
    Else
        For i = 0 To UBound(a)
            TDBGrid1.Columns(CInt(a(i))).NumberFormat = "FormatText Event"
            TDBGrid1.Columns(CInt(a(i))).DefaultValue = Format
        Next
    End If
End Sub

Private Sub UserControl_Initialize()
Dim dl As Long
Dim buffer As String * 100
    dl = GetLocaleInfo(0, &HF, buffer, 100) 'Digit Gruping
    If Left(buffer, dl - 1) = "." Then
        mLocal = True
    End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    TDBGrid1.AllowDelete = PropBag.ReadProperty("AllowDelete", True)
    TDBGrid1.AllowArrows = PropBag.ReadProperty("AllowArrows", True)
    TDBGrid1.AllowAddNew = PropBag.ReadProperty("AllowAddNew", True)
    TDBGrid1.AllowUpdate = PropBag.ReadProperty("AllowUpdate", True)
    TDBGrid1.Col = PropBag.ReadProperty("Col", 0)
    TDBGrid1.Row = PropBag.ReadProperty("Row", -1)
    TDBGrid1.HeadLines = PropBag.ReadProperty("HeadLines", 1)
    m_MaxRow = PropBag.ReadProperty("MaxRow", m_def_MaxRow)
    Set TDBGrid1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    TDBGrid1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    TDBGrid1.Bookmark = PropBag.ReadProperty("Bookmark", 0)
    TDBGrid1.SelStart = PropBag.ReadProperty("SelStart", 0)
    TDBGrid1.SelLength = PropBag.ReadProperty("SelLength", 0)
End Sub

Private Sub UserControl_Resize()
    TDBGrid1.Width = ScaleWidth
    TDBGrid1.Height = ScaleHeight
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AllowDelete", TDBGrid1.AllowDelete, True)
    Call PropBag.WriteProperty("AllowArrows", TDBGrid1.AllowArrows, True)
    Call PropBag.WriteProperty("AllowAddNew", TDBGrid1.AllowAddNew, True)
    Call PropBag.WriteProperty("AllowUpdate", TDBGrid1.AllowUpdate, True)
'    Call PropBag.WriteProperty("MaxRow", m_MaxRow, m_def_MaxRow)
    Call PropBag.WriteProperty("Col", TDBGrid1.Col, 0)
    Call PropBag.WriteProperty("Row", TDBGrid1.Row, 0)
    Call PropBag.WriteProperty("HeadLines", TDBGrid1.HeadLines, 1)
    Call PropBag.WriteProperty("MaxRow", m_MaxRow, m_def_MaxRow)
    Call PropBag.WriteProperty("Font", TDBGrid1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", TDBGrid1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Bookmark", TDBGrid1.Bookmark, 0)
    Call PropBag.WriteProperty("SelStart", TDBGrid1.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", TDBGrid1.SelLength, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,AllowUpdate
Public Property Get AllowUpdate() As Boolean
    AllowUpdate = TDBGrid1.AllowUpdate
End Property

Public Property Let AllowUpdate(ByVal New_AllowUpdate As Boolean)
    TDBGrid1.AllowUpdate() = New_AllowUpdate
    PropertyChanged "AllowUpdate"
End Property

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    ColIndex = TDBGrid1.Col
    t = TDBGrid1.Columns(ColIndex).Value
    If mType(ColIndex) = "Integer" Or Left(mType(ColIndex), 8) = "Zerofill" Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
    ElseIf mType(ColIndex) = "Decimal" Then
        If KeyAscii = 46 And mLocal Then KeyAscii = 44
        If KeyAscii = 44 And Not mLocal Then KeyAscii = 46
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii = 8 Then
                If Len(t) > 0 Then
                    t = Left(t, Len(t) - 1)
                End If
            ElseIf KeyAscii = 44 Then
                If InStr(t, ",") <> 0 Then KeyAscii = 0
            ElseIf KeyAscii = 46 Then
                If InStr(t, ".") <> 0 Then KeyAscii = 0
            Else
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    Col = TDBGrid1.Col
    If KeyCode = 13 Then
        Dim cCol As Integer
        Dim cRow As Integer
        KeyCode = 0
        cCol = TDBGrid1.Col
        cRow = TDBGrid1.Row
        Do
            cCol = cCol + 1
            If cCol = TDBGrid1.Columns.Count Then
                If TDBGrid1.AddNewMode = 0 Then cRow = cRow + 1
                cCol = 0
            End If
            If TDBGrid1.Columns(cCol).Visible And Not TDBGrid1.Columns(cCol).Locked Then Exit Do
        Loop
        TDBGrid1.Row = cRow
        TDBGrid1.Col = cCol
    ElseIf KeyCode = 117 Then
        TDBGrid1.Font.Size = TDBGrid1.Font.Size - 2
    ElseIf KeyCode = 118 Then
        TDBGrid1.Font.Size = TDBGrid1.Font.Size + 2
    ElseIf KeyCode = 119 Then
        For i = 0 To TDBGrid1.Columns.Count - 1
            TDBGrid1.Columns(i).AutoSize
        Next
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MaxRow = m_def_MaxRow
End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    RaiseEvent BeforeInsert(Cancel)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    If m_MaxRow = 0 Then Exit Sub
    If TDBGrid1.Row = m_MaxRow Then Cancel = True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,Col
Public Property Get Col() As Integer
    Col = TDBGrid1.Col
End Property

Public Property Let Col(ByVal New_Col As Integer)
    TDBGrid1.Col() = New_Col
    PropertyChanged "Col"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,Row
Public Property Get Row() As Integer
    Row = TDBGrid1.Row
End Property

Public Property Let Row(ByVal New_Row As Integer)
    TDBGrid1.Row() = New_Row
    PropertyChanged "Row"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,HeadLines
Public Property Get HeadLines() As Single
    HeadLines = TDBGrid1.HeadLines
End Property

Public Property Let HeadLines(ByVal New_HeadLines As Single)
    TDBGrid1.HeadLines() = New_HeadLines
    PropertyChanged "HeadLines"
End Property

Public Property Get RowCount() As Long
    RowCount = y.UpperBound(1) + 1
End Property
Public Property Let Cancel(ByVal tCancel As Boolean)
    mCancel = tCancel
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get MaxRow() As Byte
    MaxRow = m_MaxRow
End Property

Public Property Let MaxRow(ByVal New_MaxRow As Byte)
    m_MaxRow = New_MaxRow
    PropertyChanged "MaxRow"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,Update
Public Sub Update()
    TDBGrid1.Update
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,ReBind
Public Sub ReBind()
    TDBGrid1.ReBind
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,Font
Public Property Get Font() As Font
    Set Font = TDBGrid1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set TDBGrid1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = TDBGrid1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    TDBGrid1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    RaiseEvent RowColChange(LastRow, LastCol)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,Bookmark
Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_Description = "Sets/returns bookmark of current row"
    Bookmark = TDBGrid1.Bookmark
End Property

Public Property Let Bookmark(ByVal New_Bookmark As Variant)
    TDBGrid1.Bookmark() = New_Bookmark
    PropertyChanged "Bookmark"
End Property

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)
    RaiseEvent BeforeDelete(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Sets/returns start of selected text"
    SelStart = TDBGrid1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    TDBGrid1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TDBGrid1,TDBGrid1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Sets/returns length of selected text"
    SelLength = TDBGrid1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    TDBGrid1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property


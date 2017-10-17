Attribute VB_Name = "Control"
Private Declare Function ClipCursor Lib "user32" (lpRect As RECT) As Long
Private Declare Function ClipCursorByNum Lib "user32" Alias "ClipCursor" (ByVal _
    lpRect As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As _
    Integer
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
    lpPoint As POINTAPI) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function InvalidateRectByNum Lib "user32" Alias _
    "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, _
    ByVal bErase As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
Private Declare Sub ReleaseCapture Lib "user32" ()

Sub FindKeyDown(KeyCode As Integer, TDBGrid1 As TrueOleDBgrid80.TDBGrid, x As Object, fFind As Object)
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

Sub LoadTDBGrid(ByVal QueryString As String, TDBGrid1 As TrueOleDBgrid80.TDBGrid, y As Object)
    query QueryString
    Dim c As TrueOleDBgrid80.Column
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    Caption = Cap
    For i = 0 To RS.Fields.Count - 1
        Set c = TDBGrid1.Columns.Add(i)
        c.Caption = RS.Fields(i).Name
        c.HeadAlignment = dbgCenter
        c.Width = TDBGrid1.Parent.TextWidth(c.Caption) + 200
        If RS.Fields(i).Type = 3 Or RS.Fields(i).Type = 5 Then c.Alignment = dbgRight
        If RS.Fields(i).Type = 5 Then c.NumberFormat = "standard"
        c.Visible = True
    Next
    If RS.RecordCount = 0 Then Exit Sub
    y.LoadRows RS.GetRows
    Set TDBGrid1.Array = y
    TDBGrid1.Rebind
    For i = 0 To TDBGrid1.Columns.Count - 1
        cc = cc + TDBGrid1.Columns(i).Width
    Next
    TDBGrid1.Width = cc + 500
End Sub

Sub UnloadAll()
    Dim f As Form
    For Each f In Forms
        Unload f
    Next
    'ExitProcess 0
End Sub

Sub MustNumber(key As Integer)
    If key <> 8 And (key < 48 Or key > 57) Then key = 0
End Sub

Sub LookUp(key As Integer)
On Error GoTo err
Dim s As String, l As Long
Dim c As ComboBox
    Set c = Screen.ActiveControl
    s = Left$(c.Text, c.SelStart) & Chr$(key)
    l = SendMessage(c.hwnd, &H14C, -1, ByVal s)
    If l <> -1 Then
        With c
            .ListIndex = l
            .Text = .List(l)
            .SelStart = Len(s)
            .SelLength = Len(.Text)
        End With
        key = 0
    End If
    Exit Sub
err:
End Sub

Sub LoadCaption(obj As Object, ByVal Cap As String, Optional ByVal tCount As Long = -1, Optional ByVal pTop As Long = -1)
On Error Resume Next
    i = obj.Count
    If tCount >= i Then obj(i - 1).Container.LoadLabel obj Else i = tCount
    If pTop = -1 Then
        If i > 1 Then obj(i).Top = obj(i - 1).Top + obj(i - 1).Height
    Else
        obj(i).Top = pTop
    End If
    obj(i).Visible = True
    obj(i) = Cap
End Sub


Sub ClearGrid(TDBGrid As TrueOleDBgrid80.TDBGrid, y As Object)
    y.ReDim 0, 0, 0, 0
    y(0, 0) = ""
    y.ReDim 0, 0, 0, TDBGrid.Columns.Count - 1
    TDBGrid.Rebind
End Sub
Sub ClearDD(TDBDropDown As TrueOleDBgrid80.TDBDropDown, y As Object)
    y.ReDim 0, 0, 0, 0
    y(0, 0) = ""
    y.ReDim 0, 0, 0, TDBDropDown.Columns.Count - 1
    TDBDropDown.Rebind
End Sub

Function FindIndex(ByVal StringToFind As String, ByVal combo As ComboBox, Optional ByVal Incremental As Boolean = False, Optional ByVal NextData As Boolean = False)
    FindIndex = -1
    If Not Incremental Then
        For i = 0 To combo.ListCount - 1
            If StrComp(combo.List(i), StringToFind, vbTextCompare) = 0 Then
                FindIndex = i
                Exit Function
            End If
        Next
    Else
        l = 0
        h = combo.ListCount - 1
        Do
            If l > h Then Exit Do
            m = (l + h) \ 2
            If StrComp(StringToFind, combo.List(m), vbTextCompare) = 0 Then
                FindIndex = m
                If Not NextData Then Exit Function
                h = m - 1
            ElseIf StrComp(StringToFind, combo.List(m), vbTextCompare) = 1 Then
                l = m + 1
            ElseIf StrComp(StringToFind, combo.List(m), vbTextCompare) = -1 Then
                h = m - 1
            End If
        Loop
    End If
End Function

Function FindIndex2(ByVal StringToFind As String, ByVal NumIndex As Byte, ByRef rs1() As Variant, ByVal num As Long, Optional ByVal Incremental As Boolean = True, Optional ByVal NextData As Boolean = True)
    FindIndex2 = -1
    If Not Incremental Then
        For i = 0 To num - 1
            If StrComp(rs1(NumIndex, i), StringToFind, vbTextCompare) = 0 Then
                FindIndex2 = i
                Exit Function
            End If
        Next
    Else
        l = 0
        h = num - 1
        Do
            If l > h Then Exit Do
            m = (l + h) \ 2
            If StrComp(StringToFind, rs1(NumIndex, m), vbTextCompare) = 0 Then
                FindIndex2 = m
                If Not NextData Then Exit Function
                h = m - 1
            ElseIf StrComp(StringToFind, rs1(NumIndex, m), vbTextCompare) = 1 Then
                l = m + 1
            ElseIf StrComp(StringToFind, rs1(NumIndex, m), vbTextCompare) = -1 Then
                h = m - 1
            End If
        Loop
    End If
End Function



' Drag a control until the user releases all mouse buttons
'
' You should call this routine from the MouseDown event procedures
' of the controls that you want to make draggable, after
' you determine that the user has initiated a drag operation.
' For example, if you want to let the user drag controls
' using the Ctrl+Right button combination, add this code
' to their MouseDown procedure:
'
' Private Sub Command1_MouseDown(...)
'    If Button = vbRightButton And Shift = vbCtrlMask Then
'        DragControl Command1
'    End If
' End Sub
'
' From that point on, this procedure takes the control and
' exits only when the user releases all mouse buttons

Sub DragControl(ctrl As Object)
    Dim startButton As Integer
    Dim startPoint As POINTAPI
    Dim currPoint As POINTAPI
    Dim contRect As RECT
    Dim contScaleMode As Integer
    
    ' get mouse position and buttons pressed
    GetCursorPos startPoint
    If GetAsyncKeyState(vbLeftButton) Then startButton = vbLeftButton
    If GetAsyncKeyState(vbRightButton) Then startButton = startButton Or _
        vbRightButton
    If GetAsyncKeyState(vbMiddleButton) Then startButton = startButton Or _
        vbMiddleButton
        
    ' get container upper-left corner position
    ' in screen coordinates (currPoint is Zero)
    ClientToScreen ctrl.Container.hwnd, currPoint
    ' get container size
    GetClientRect ctrl.Container.hwnd, contRect
    ' convert to screen coordintes
    contRect.Left = currPoint.x
    contRect.Top = currPoint.y
    contRect.Right = contRect.Right + currPoint.x
    contRect.Bottom = contRect.Bottom + currPoint.y
    ' limit the cursor within the parent control
    ClipCursor contRect
    
    ' get the ScaleMode that is active for the control
    ' this is the ScaleMode of its container, or it
    ' is vbTwips if its container does not support
    ' the ScaleMode property
    On Error Resume Next
    contScaleMode = vbTwips
    ' ignore next assignement if the container
    ' dows not support ScaleMode property
    contScaleMode = ctrl.Container.ScaleMode
    
    Do
        ' exit if all mouse buttons are released
        If (startButton And vbLeftButton) = 0 Or GetAsyncKeyState(vbLeftButton) _
            = 0 Then
            If (startButton And vbRightButton) = 0 Or GetAsyncKeyState _
                (vbRightButton) = 0 Then
                If (startButton And vbMiddleButton) = 0 Or GetAsyncKeyState _
                    (vbMiddleButton) = 0 Then
                    Exit Do
                End If
            End If
        End If
        
        ' get current mouse position
        GetCursorPos currPoint
        If currPoint.x <> startPoint.x Or currPoint.y <> startPoint.y Then
            With ctrl.Parent
                ctrl.Move ctrl.Left + .ScaleX(currPoint.x - startPoint.x, _
                    vbPixels, contScaleMode), ctrl.Top + .ScaleY(currPoint.y - _
                    startPoint.y, vbPixels, contScaleMode)
                InvalidateRectByNum .hwnd, 0, False
                .Refresh
            End With
            LSet startPoint = currPoint
        End If
        DoEvents
    Loop
    ClipCursorByNum 0
End Sub

Sub SetControlProp(f As Object, ByVal tControls As String, ByVal tProperties As String, tVal As Variant)
    Dim a() As String
    Dim b() As String
    a = Split(tControls, ".")
    b = Split(tProperties, ".")
    For i = 0 To UBound(a)
        For j = 0 To UBound(b)
            CallByName f.Controls(a(i)), b(j), VbLet, tVal
        Next
    Next
End Sub

Sub FormMouseMove(ByVal Button As Integer, f As Form)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(f.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

    

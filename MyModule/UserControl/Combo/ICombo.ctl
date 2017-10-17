VERSION 5.00
Begin VB.UserControl ICombo 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ICombo.ctx":0000
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "ICombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mCancel As Boolean
Dim mSorted As Boolean
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
'Event Declarations:
Event DropDown() 'MappingInfo=Combo1,Combo1,-1,DropDown
Attribute DropDown.VB_Description = "Occurs when the list portion of a ComboBox control is about to drop down."
'Event Click() 'MappingInfo=Combo1,Combo1,-1,Click

Public Function Validate()
    Validate = True
    If FindIndex(Combo1) = -1 Then Validate = False
    If Not Validate Then SetFocus
End Function

Public Sub Clear()
    Combo1.Clear
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    If KeyCode = 13 Then SendKeys Chr(9)
End Sub

Function FindIndex(ByVal s As String) As Integer
    ll = Len(s)
    FindIndex = -1
    If s = "" Then Exit Function
    With Combo1
        If Not mSorted Then
            For i = 0 To .ListCount - 1
                If InStr(1, .List(i), s, vbTextCompare) = 1 Then
                    .ListIndex = i
                    .Text = .List(i)
                    FindIndex = i
                    Exit For
                End If
            Next
        Else
            L = 0
            h = .ListCount - 1
            f = -1
            Do
                If L > h Then Exit Do
                m = (L + h) \ 2
                If InStr(1, .List(m), s, vbTextCompare) = 1 Then
                    f = m
                    h = m - 1
                    .ListIndex = m
                    .Text = .List(m)
                    FindIndex = m
                    Exit Do
                ElseIf StrComp(Left(.List(m), ll), s, vbTextCompare) = -1 Then
                    L = m + 1
                ElseIf StrComp(Left(.List(m), ll), s, vbTextCompare) = 1 Then
                    h = m - 1
                End If
            Loop
        End If
    End With
End Function

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    If Combo1.Locked Then Exit Sub
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    s = Left$(Combo1, Combo1.SelStart) & Chr$(KeyAscii)
    If FindIndex(s) = -1 Then Combo1 = s
    Combo1.SelStart = Len(s)
    Combo1.SelLength = Len(Combo1)
    KeyAscii = 0
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Sorted = mSorted
End Property

Public Property Let Sorted(ByVal tSorted As Boolean)
    mSorted = tSorted
End Property

Private Sub Swap(ByVal L As Long, ByVal r As Long)
    t = Combo1.List(L)
    Combo1.List(L) = Combo1.List(r)
    Combo1.List(r) = t
End Sub

Private Sub Combo1_LostFocus()
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    If FindIndex(Combo1) = -1 Then
        Combo1 = ""
    End If
End Sub

Private Sub UserControl_Resize()
    Combo1.Width = ScaleWidth
    Height = 315
End Sub

Private Sub Combo1_GotFocus()
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1)
End Sub

Public Property Let Cancel(tCancel As Boolean)
    mCancel = tCancel
End Property
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Text.VB_UserMemId = 0
    Text = Combo1.Text
End Property

Public Property Let Text(ByVal tText As String)
    Combo1.Text = tText
    PropertyChanged "Text"
End Property

Public Property Get ListCount() As Long
    ListCount = Combo1.ListCount
End Property

Public Property Get List(ByVal tIndex As Long) As String
Attribute List.VB_ProcData.VB_Invoke_Property = ";Data"
    List = Combo1.List(tIndex)
End Property

Public Property Let List(ByVal tIndex As Long, ByVal tList As String)
    For i = Combo1.ListCount To tIndex - 1
        Combo1.List(i) = ""
        Combo1.ItemData(i) = ""
    Next
    Combo1.List(tIndex) = tList
End Property
Public Property Get ItemData(ByVal tIndex As Long) As String
Attribute ItemData.VB_ProcData.VB_Invoke_Property = ";Data"
    ItemData = Combo1.ItemData(tIndex)
End Property

Public Property Let ItemData(ByVal tIndex As Long, ByVal tItemData As String)
    Combo1.ItemData(tIndex) = tItemData
End Property

Public Property Get Enabled() As Boolean
    Enabled = Combo1.Enabled
End Property

Public Property Let Enabled(ByVal tEnabled As Boolean)
    Combo1.Enabled = tEnabled
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Combo1.Text = PropBag.ReadProperty("Text", "")
    Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    Combo1.Locked = PropBag.ReadProperty("Locked", False)
    mSorted = PropBag.ReadProperty("Sorted", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", Combo1.Text, ""
    PropBag.WriteProperty "Enabled", Combo1.Enabled, True
    PropBag.WriteProperty "Sorted", mSorted, False
    Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)
    Call PropBag.WriteProperty("Locked", Combo1.Locked, False)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
    Set Font = Combo1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Combo1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Font"
    ForeColor = Combo1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Combo1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Combo1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = Combo1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Combo1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
'
'Private Sub Combo1_Click()
'    RaiseEvent Click
'End Sub
'
Private Sub Combo1_DropDown()
    RaiseEvent DropDown
End Sub


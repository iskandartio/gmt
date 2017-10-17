VERSION 5.00
Begin VB.UserControl IText 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "IText.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Tag             =   "0"
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "IText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mDataType As mType
Private mCancel As Boolean
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Enum mType
    tString = 0
    tInteger = 1
    tCurrency = 2
    tDate = 3
    tPassword = 4
    tUcaseString = 5
    tLongDate = 6
End Enum

Public Function aa()
    aa = Replace(Text1, "'", "''")
End Function
Public Sub SelText()
    Text1_GotFocus
End Sub
Public Function Validate()
    Validate = True
    If mDataType = 3 Or mDataType = 6 Then
        If Not DateType(Text1) Then Validate = False
    Else
        If Text1.Text = "" Then Validate = False
    End If
    If Not Validate Then SetFocus
End Function

Private Function search_max(ByVal m As Integer, ByVal y As Integer)
    search_max = 31
    If m = 4 Or m = 6 Or m = 9 Or m = 11 Then
        search_max = 30
    ElseIf m = 2 Then
        If y Mod 4 = 0 Then cari_max = 29 Else search_max = 28
    End If
End Function

Function DateType(ByVal a As String)
On Error GoTo err
    DateType = False
    tanggal = CInt(Left(a, 2))
    bulan = CInt(Mid(a, 4, 2))
    tahun = CInt(Right(a, 2))
    maximum = search_max(bulan, tahun)
    If tanggal > 0 And bulan > 0 And tanggal <= maximum And bulan < 13 Then
        DateType = True
        Exit Function
    End If
err:
End Function
   
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    If KeyCode = 13 Then SendKeys Chr(9)
    If Text1.Locked Then Exit Sub
    If KeyCode = 46 And (mDataType = 3 Or mDataType = 6) Then
        a = Text1.Text
        l = Text1.SelStart
        If l = 2 Or l = 5 Then
            KeyCode = 0
            Exit Sub
        End If
        Text1 = Left(a, l + 1) & "_" & Mid(a, l + 2)
        Text1.SelStart = l
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    On Error Resume Next
    If Text1.Locked Then Exit Sub
    If mDataType = 1 Then
        If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
    ElseIf mDataType = 2 Then
        a = Text1
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii = 8 Then
                a = Left(a, Len(a) - 1)
            ElseIf KeyAscii = 46 Then
                If InStr(a, ",") <> 0 Then KeyAscii = 0
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf mDataType = 3 Or mDataType = 6 Then
        If KeyAscii <> 47 And KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 8 Then
            a = Text1
            l = Text1.SelStart
            If l = 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            If l = 3 Or l = 6 Then l = l - 1
            Text1 = Left(a, l) & "_" & Mid(a, l + 1)
            Text1.SelStart = l
        ElseIf KeyAscii <> 0 Then
            a = Text1
            l = Text1.SelStart
            If (l = 8 And mDataType = 3) Or (l = 10 And mDataType = 6) Then KeyAscii = 0
            If l = 2 Or l = 5 Then l = l + 1
            Text1.Text = Left(a, l) & Mid(a, l + 2)
            Text1.SelStart = l
        End If
    ElseIf mDataType = tUcaseString Then
        If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    If mDataType = 3 And KeyCode <> 37 Then
        If Text1.SelStart = 2 Or Text1.SelStart = 5 Then Text1.SelStart = Text1.SelStart + 1
    End If
End Sub

Private Sub UserControl_GotFocus()
    If Not Enabled Then SendKeys Chr(9)
End Sub

Sub FocusSelect()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub UserControl_Resize()
    Text1.Width = ScaleWidth
    If mDataType = 3 Then
        Width = 900
    ElseIf mDataType = 6 Then
        Width = 1100
    End If
    Height = TextHeight("I") * 1.4
End Sub

Private Sub Text1_GotFocus()
    If mCancel Then
        mCancel = False
        Exit Sub
    End If
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Public Property Get DataType() As mType
Attribute DataType.VB_ProcData.VB_Invoke_Property = ";Data"
    DataType = mDataType
End Property

Public Property Let DataType(ByVal tDataType As mType)
    mDataType = tDataType
    Text1 = ""
    Text1.PasswordChar = ""
    Text1.MaxLength = 0
    If mDataType = 3 Then
        'Text1.MaxLength = 8
        Text1 = "__/__/__"
        Width = 900
    ElseIf mDataType = 4 Then
        Text1.PasswordChar = "*"
    ElseIf mDataType = 6 Then
        
        Text1.Text = "__/__/____"
        Width = 1100
    End If
    PropertyChanged "DataType"
End Property

Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Text.VB_UserMemId = 0
    Text = Text1
End Property

Public Property Let Text(ByVal tText As String)
On Error GoTo err
    If mDataType = 3 Or mDataType = 6 Then
        Mid(tText, 3, 1) = "/"
        Mid(tText, 6, 1) = "/"
    ElseIf mDataType = 1 Then
        tText = CLng(tText)
    ElseIf mDataType = 2 Then
        tText = CDbl(tText)
    End If
    Text1 = tText
    PropertyChanged "Text"
    Exit Property
err:
    Text1 = ""
End Property

Public Property Let Cancel(ByVal tCancel As Boolean)
Attribute Cancel.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    mCancel = tCancel
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Text1 = PropBag.ReadProperty("Text", "")
    mDataType = PropBag.ReadProperty("DataType", 0)
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    FontSize = Text1.FontSize
    Height = 1.4 * TextHeight("I")
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    Text1.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", Text1.Text, ""
    PropBag.WriteProperty "DataType", mDataType, 0
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
    Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
    Call PropBag.WriteProperty("Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("PasswordChar", Text1.PasswordChar, "")
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
    Height = TextHeight("I") * 1.4
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
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
Attribute MaxLength.VB_ProcData.VB_Invoke_Property = ";Data"
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = Text1.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    Text1.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property


VERSION 5.00
Begin VB.UserControl iButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "iButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mCaption As String
Dim mUp1 As Long
Dim mDown1 As Long
Dim mUp2 As Long
Dim mDown2 As Long
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Dim mBackColor As Long
Dim mForeColor As Long
Dim mLevel As Long
Dim mClickedColor As OLE_COLOR

Public Property Get Level() As Byte
    Level = mLevel
End Property
Public Property Let Level(tLevel As Byte)
    mLevel = tLevel
    PaintControl
End Property
Public Property Get ClickedColor() As OLE_COLOR
    ClickedColor = mClickedColor
End Property
Public Property Let ClickedColor(tColor As OLE_COLOR)
    mClickedColor = tColor
End Property
Public Property Get Caption() As String
    Caption = mCaption
End Property
Public Property Let Caption(tCaption As String)
    mCaption = tCaption
    PaintControl
End Property

Private Sub UserControl_Initialize()
    mBackColor = GetSysColor(&HF)
    mForeColor = GetSysColor(&H12)
End Sub

Private Sub UserControl_LostFocus()
    PaintControl
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cls
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), mDown2, B
    For i = 0 To mLevel
        Line (1 + i, 1 + i)-(ScaleWidth - 2 - i, ScaleHeight - 2 - i), mDown1, B
    Next
    PrintClicked
    PrintCaption
End Sub
Private Sub PrintCaption()
    b = Split(mCaption, "`")
    CurrentY = (ScaleHeight - (UBound(b) + 1) * TextHeight("I")) / 2
    For i = 0 To UBound(b)
        CurrentX = (ScaleWidth - TextWidth(b(i))) / 2
        Print b(i)
    Next
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PaintControl True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mCaption = PropBag.ReadProperty("Caption", "")
    mLevel = PropBag.ReadProperty("Level", 0)
    mClickedColor = PropBag.ReadProperty("ClickedColor", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", mBackColor)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", mForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    PaintControl
End Sub

Private Sub UserControl_Show()
    ChooseColor
    PaintControl
End Sub
Private Sub ChooseColor()
    If BackColor < 0 Then BackColor = mBackColor
    r = BackColor Mod &H100
    g = (BackColor \ &H100) Mod &H100
    b = BackColor \ &H10000
    l = r
    h = r
    If g < r Then l = g
    If b < r Then l = b
    If g > h Then h = g
    If b > h Then h = b
    lum = (l + h) / 255 * 120
    If lum = 240 Then
        r = 255
        g = 255
        b = 255
    ElseIf lum = 0 Then
        r = 0
        g = 0
        b = 0
    ElseIf lum > 120 Then
        r = (120 * r - 255 * (lum - 120)) / (240 - lum)
        g = (120 * g - 255 * (lum - 120)) / (240 - lum)
        b = (120 * b - 255 * (lum - 120)) / (240 - lum)
    Else
        r = r * 120 / lum
        g = g * 120 / lum
        b = b * 120 / lum
    End If
    r = IIf(r < 0, 0, r)
    g = IIf(g < 0, 0, g)
    b = IIf(b < 0, 0, b)
    Lum2 = lum + 0.7 * (240 - lum)
    mUp1 = GetColorByLum(Lum2, r, g, b)
    Lum2 = lum + 0.8 * (240 - lum)
    mUp2 = GetColorByLum(Lum2, r, g, b)
    Lum2 = 0.9 * lum
    mDown1 = GetColorByLum(Lum2, r, g, b)
    Lum2 = 0.8 * lum
    mDown2 = GetColorByLum(Lum2, r, g, b)
End Sub

Private Function GetColorByLum(ByVal Lum2 As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)
    If Lum2 > 120 Then
        GetColorByLum = RGB(r + (255 - r) * (Lum2 / 120 - 1), g + (255 - g) * (Lum2 / 120 - 1), b + (255 - b) * (Lum2 / 120 - 1))
    Else
        GetColorByLum = RGB(r / 120 * Lum2, g / 120 * Lum2, b / 120 * Lum2)
    End If
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", mCaption, ""
    PropBag.WriteProperty "ClickedColor", mClickedColor, 0
    PropBag.WriteProperty "Level", mLevel, 0
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, mBackColor)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, mForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Sub PaintControl(Optional ByVal tClicked As Boolean = False)
On Error Resume Next
    Cls
    Line (1, 1)-(ScaleWidth - 2, 1), mUp2
    Line (1, 1)-(1, ScaleHeight - 2), mUp2
    Line (1, ScaleHeight - 2)-(ScaleWidth - 2, ScaleHeight - 2), mDown2
    Line (ScaleWidth - 2, 1)-(ScaleWidth - 2, ScaleHeight - 1), mDown2
    For i = 0 To mLevel
        Line (2, 2 + i)-(ScaleWidth - 2 - i, 2 + i), mUp1
        Line (2 + i, 2)-(2 + i, ScaleHeight - 2 - i), mUp1
        Line (2 + i, ScaleHeight - 3 - i)-(ScaleWidth - 3, ScaleHeight - 3 - i), mDown1
        Line (ScaleWidth - 3 - i, 2 + i)-(ScaleWidth - 3 - i, ScaleHeight - 2), mDown1
    Next
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), Parent.BackColor, B
    PrintCaption
    If tClicked Then PrintClicked
End Sub
Private Sub PrintClicked()
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), mClickedColor, B
    For i = 5 + mLevel To ScaleWidth - mLevel - 6 Step 2
        PSet (i, 4 + mLevel), mClickedColor
        PSet (i, ScaleHeight - 5 - mLevel), mClickedColor
    Next
    For i = 5 + mLevel To ScaleHeight - mLevel - 6 Step 2
        PSet (4 + mLevel, i), mClickedColor
        PSet (ScaleWidth - 5 - mLevel, i), mClickedColor
    Next
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    ChooseColor
    PaintControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    PaintControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    PaintControl
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property



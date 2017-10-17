VERSION 5.00
Begin VB.UserControl iToolbar 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   ScaleHeight     =   630
   ScaleWidth      =   10305
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7920
      TabIndex        =   12
      Top             =   10
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   495
      Index           =   11
      Left            =   7440
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   495
      Index           =   10
      Left            =   6960
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   495
      Index           =   9
      Left            =   6480
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   495
      Index           =   8
      Left            =   6000
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&XIT"
      Height          =   495
      Index           =   7
      Left            =   5400
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&LIST"
      Height          =   495
      Index           =   6
      Left            =   4680
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&PRINT"
      Height          =   495
      Index           =   5
      Left            =   3840
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&CANCEL"
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   2880
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SAVE"
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&DELETE"
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&EDIT"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&NEW"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "iToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event NewClick()
Event EditClick()
Event DeleteClick()
Event SaveClick()
Event CancelClick()
Event PrintClick()
Event ListClick()
Event ExitClick()
Event TopClick()
Event BottomClick()
Event PrevClick()
Event NextClick()
Event TextEnter()

Dim mIndex As Integer
Dim mCekValid As Boolean
Dim mBagSee1 As Long
Dim mBagAdd1 As Long
Dim mBagEdit1 As Long
Dim mBagDelete1 As Long
Dim mBagPrint1 As Long
Dim mBagSee2 As Long
Dim mBagAdd2 As Long
Dim mBagEdit2 As Long
Dim mBagDelete2 As Long
Dim mBagPrint2 As Long
Dim mGagal As Boolean
Dim mNoData As Boolean

Sub AuthorizationKey(ByVal tSee1 As Long, ByVal tSee2 As Long, ByVal tAdd1 As Long, ByVal tAdd2 As Long, _
    ByVal tEdit1 As Long, ByVal tEdit2 As Long, ByVal tDelete1 As Long, ByVal tDelete2 As Long, ByVal tPrint1 As Long, ByVal tPrint2 As Long)
    mBagSee1 = tSee1
    mBagSee2 = tSee2
    mBagAdd1 = tAdd1
    mBagAdd2 = tAdd2
    mBagEdit1 = tEdit1
    mBagEdit2 = tEdit2
    mBagDelete1 = tDelete1
    mBagDelete2 = tDelete2
    mBagPrint1 = tPrint1
    mBagPrint2 = tPrint2
    mCekValid = True
End Sub

Sub SetQuick(ByVal tText As String)
    mNoData = False
    Text1 = tText
End Sub

Private Sub Command1_Click(Index As Integer)
    mIndex = Index
    If mCekValid Then
        If Not Cek_Valid(mIndex, Parent.Tag) Then Exit Sub
    End If
    If mIndex = 0 Then
        RaiseEvent NewClick
    ElseIf mIndex = 1 Then
        RaiseEvent EditClick
    ElseIf mIndex = 2 Then
        RaiseEvent DeleteClick
    ElseIf mIndex = 3 Then
        RaiseEvent SaveClick
    ElseIf mIndex = 4 Then
        RaiseEvent CancelClick
    ElseIf mIndex = 5 Then
        RaiseEvent PrintClick
    ElseIf mIndex = 6 Then
        RaiseEvent ListClick
    ElseIf mIndex = 7 Then
        RaiseEvent ExitClick
    ElseIf mIndex = 8 Then
        RaiseEvent TopClick
    ElseIf mIndex = 9 Then
        RaiseEvent PrevClick
    ElseIf mIndex = 10 Then
        RaiseEvent NextClick
    ElseIf mIndex = 11 Then
        RaiseEvent BottomClick
    End If
    If mGagal Then
        mGagal = False
        Exit Sub
    End If
    SetClick mIndex
End Sub

Sub SetClick(ByVal tIndex As Integer)
Dim v As Boolean
Dim i As Byte
    If mNoData = True And tIndex <> 0 Then
        Command1(0).Enabled = True
        For i = 1 To Command1.Count - 1
            Command1(i).Enabled = False
        Next
        Exit Sub
    End If
    mIndex = tIndex
    If mIndex <> 0 And mIndex <> 1 And mIndex <> 3 And mIndex <> 4 And mIndex <> 12 Then Exit Sub
    v = mIndex = 0 Or mIndex = 1
    Command1(0).Enabled = Not v
    Command1(1).Enabled = Not v
    Command1(2).Enabled = Not v
    Command1(3).Enabled = v
    Command1(4).Enabled = v
    Command1(5).Enabled = Not v
    'Command1(6).Enabled = Not v
    'Command1(7).Enabled = Not v
    'Command1(8).Enabled = Not v
    'Command1(9).Enabled = Not v
    'command1(10).Enabled = Not v
    'Command1(11).Enabled = Not v
End Sub

Sub SetIndex(ByVal tIndex As Integer)
    mIndex = tIndex
End Sub

Function GetIndex() As Integer
    GetIndex = mIndex
End Function

Function GetEnabled(ByVal tIndex As Integer) As Boolean
    GetEnabled = Command1(tIndex).Enabled
End Function

Function GetText() As String
    GetText = Text1.Text
End Function

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        mIndex = 12
        If mCekValid Then
            If Not Cek_Valid(mIndex, Parent.Tag) Then Exit Sub
        End If
        RaiseEvent TextEnter
        'SetClick mIndex
    End If
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    mBagSee1 = pBagSee1
    mBagAdd1 = pBagAdd1
    mBagEdit1 = pBagEdit1
    mBagDelete1 = pBagDelete1
    mBagPrint1 = pBagPrint1
    mBagSee2 = pBagSee2
    mBagAdd2 = pBagAdd2
    mBagEdit2 = pBagEdit2
    mBagDelete2 = pBagDelete2
    mBagPrint2 = pBagPrint2
    mCekValid = True
    pScript.AddCode "iNew = 0"
    pScript.AddCode "iEdit = 1"
    pScript.AddCode "iDelete = 2"
    pScript.AddCode "iSave = 3"
    pScript.AddCode "iCancel = 4"
    pScript.AddCode "iPrint = 5"
    pScript.AddCode "iList = 6"
    pScript.AddCode "iExit = 7"
    pScript.AddCode "iPrev = 8"
    pScript.AddCode "iTop = 9"
    pScript.AddCode "iNext = 10"
    pScript.AddCode "iBottom = 11"
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    RearrangePosition
    Height = Command1(0).Height
    Text1.Width = ScaleWidth - Text1.Left
End Sub

Sub SetVisible(ByVal tCommand As String, Optional ByVal tVal As Boolean = False)
Dim tCommandIndex As Integer
    tCommandIndex = pScript.Eval("i" & tCommand)
    Command1(tCommandIndex).Visible = tVal
    Command1(tCommandIndex).Tag = "0"
    RearrangePosition
End Sub

Function GetIndexByName(ByVal tCommand As String) As Integer
    GetIndexByName = pScript.Eval("i" & tCommand)
End Function
Sub SetEnabled(ByVal tCommand As String, ByVal tValue As Boolean)
    Command1(GetIndexByName(tCommand)).Enabled = tValue
End Sub

Sub SetNoData(ByVal tVal As Boolean)
Dim i As Byte
    mNoData = tVal
    If tVal Then SetClick -1
End Sub

Sub SetGagalSave()
    mGagal = True
End Sub

Sub RearrangePosition()
Dim i As Integer
Dim t As Single
    t = 0
    For i = 0 To Command1.Count - 1
        If Command1(i).Tag <> "0" Then
            Command1(i).Left = t
            t = t + Command1(i).Width + 10
        End If
    Next
    Text1.Left = t + 10
End Sub

Function Cek_Valid(ByVal tIndex As Integer, ByVal tTag As Integer) As Boolean
Dim v As Long
Dim tPass As Long
    If tTag < 31 Then
        v = 1
        tPass = 2 ^ tTag
        If tIndex = 6 Or tIndex = 8 Or tIndex = 9 Or tIndex = 10 Or tIndex = 11 Or tIndex = 12 Then 'Waktu GetResult
            v = mBagSee1 And tPass
        ElseIf tIndex = 0 Then
            v = mBagAdd1 And tPass
        ElseIf tIndex = 1 Then
            v = mBagEdit1 And tPass
        ElseIf tIndex = 2 Then
            v = mBagDelete1 And tPass
        ElseIf tIndex = 5 Then
            v = mBagPrint1 And tPass
        End If
    Else
        tPass = 2 ^ (tTag - 31)
        If tIndex = 6 Or tIndex = 8 Or tIndex = 9 Or tIndex = 10 Or tIndex = 11 Or tIndex = 12 Then 'Waktu GetResult
            v = mBagSee2 And tPass
        ElseIf tIndex = 0 Then
            v = mBagAdd2 And tPass
        ElseIf tIndex = 1 Then
            v = mBagEdit2 And tPass
        ElseIf tIndex = 2 Then
            v = mBagDelete2 And tPass
        ElseIf tIndex = 5 Then
            v = mBagPrint2 And tPass
        End If
    End If
    Cek_Valid = IIf(v = 0, False, True)
    If Not Cek_Valid Then MsgBox "TIDAK BERHAK !!!"
End Function




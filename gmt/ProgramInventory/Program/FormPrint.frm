VERSION 5.00
Begin VB.Form FormPrint 
   Caption         =   "PRINT"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox fTopMargin 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox fLeftMargin 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox fListPrinter 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton fCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton fNo 
      Caption         =   "&No"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton fYes 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Top Margin"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Left Margin"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label dMsg 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "FormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LPrinter As Boolean

Private Sub fCancel_Click()
    FormPreview.Picture1.Tag = vbCancel
    Unload Me
End Sub

Private Sub fLeftMargin_Validate(Cancel As Boolean)
    pLeftMargin = fLeftMargin
End Sub

Private Sub fListPrinter_GotFocus()
    If Not LPrinter Then
        For i = 0 To Printers.Count - 1
            fListPrinter.List(i) = Printers(i).DeviceName
            If fListPrinter.List(i) = Printer.DeviceName Then fListPrinter.ListIndex = i
        Next
        LPrinter = True
    End If
End Sub

Private Sub fListPrinter_Validate(Cancel As Boolean)
    For i = 0 To Printers.Count - 1
        If fListPrinter = Printers(i).DeviceName Then Set Printer = Printers(i)
    Next
    dMsg = "Langsung Print dengan " & Printer.DeviceName & "?"
End Sub

Private Sub fNo_Click()
    FormPreview.Picture1.Tag = vbNo
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If TypeName(ActiveControl) = "ComboBox" Then Exit Sub
    If KeyCode = 78 Then
        fNo.SetFocus
        fNo_Click
    ElseIf KeyCode = 89 Then
        fYes.SetFocus
        fYes_Click
    ElseIf KeyCode = 67 Then
        fCancel.SetFocus
        fCancel_Click
    End If
End Sub

Private Sub Form_Load()
    LPrinter = False
    dMsg = "Langsung Print dengan " & Printer.DeviceName & "?"
    fListPrinter = Printer.DeviceName
    a = "select top 1 Kiri, Atas from m_MarginPrinter where Nama='" & esc(pSettingName) & "'"
    query a
    If RS.RecordCount > 0 Then
        pLeftMargin = RS.Fields("Kiri").Value
        pTopMargin = RS.Fields("Atas").Value
    End If
    fLeftMargin = pLeftMargin
    fTopMargin = pTopMargin
End Sub

Private Sub fTopMargin_Validate(Cancel As Boolean)
    fTopMargin = pTopMargin
End Sub

Private Sub fYes_Click()
    FormPreview.Picture1.Tag = vbYes
    Unload Me
End Sub

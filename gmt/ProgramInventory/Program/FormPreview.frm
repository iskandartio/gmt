VERSION 5.00
Begin VB.Form FormPreview 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PREVIEW"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox fPagesPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox fPage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton fLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fNext 
      Caption         =   ">"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   6
         Left            =   0
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1935
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2655
         LargeChange     =   6
         Left            =   3360
         Max             =   5
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   15500
         Left            =   50
         ScaleHeight     =   15465
         ScaleWidth      =   11970
         TabIndex        =   1
         Top             =   50
         Width           =   12000
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label fTotalPage 
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FormPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mReportForm As String
Dim mParams As String
Dim mSQL As String
Dim mSQLAfterPrint As String
Dim mPage As Integer
Dim mPageSign() As Long
Dim mTotalPage As Integer
Dim mSetPageHeight As Boolean
Dim f As Form
Private Const WM_MOUSEWHEEL = 522

Sub SetPageHeight(ByVal tPageHeight As Single)
    mSetPageHeight = True
    Picture1.Height = tPageHeight
End Sub

Function GetSetPageHeight() As Boolean
    GetSetPageHeight = mSetPageHeight
End Function

Private Function CariSemuaPageSign() As Integer
    j = 1
    Do
        mPageSign(j) = f.PrintData(mPageSign(j - 1), 0, j, 0, True)
        j = j + 1
        If mTotalPage <> 0 Then Exit Do
    Loop
    CariSemuaPageSign = mTotalPage
End Function

Sub SetTotalPage(ByVal tTotalPage As Long)
    mTotalPage = tTotalPage
    mPage = mTotalPage - 1
    fTotalPage = "of " & mTotalPage & " Page(s)"
End Sub

Sub LoadMe(tf As Form, ByVal tReportForm As String, Optional ByVal tParams As String, Optional ByVal tSQL As String, Optional ByVal tSQLAfterPrint As String, Optional ByVal tPageHeight As Single = 0)
    Printer.Orientation = vbPRORPortrait
    VScroll1.Tag = False
    mReportForm = tReportForm
    mParams = tParams
    mSQL = tSQL
    mSQLAfterPrint = tSQLAfterPrint
    mPage = 1
    FormPrint.Show vbModal
    b = Picture1.Tag
    If b = vbNo Then
        Picture1.Cls
        If Not InitReport(Picture1, False) Then
            MsgBox "No Data"
            Unload Me
            Exit Sub
        End If
        PreviewReport Picture1
        Show
    ElseIf b = vbYes Then
        If Not InitReport(Picture1, False) Then
            MsgBox "No Data"
            Unload Me
            Exit Sub
        End If
        PrintReport
        If mSQLAfterPrint <> "" Then ExecMe mSQLAfterPrint
    End If
End Sub

Private Sub PrintReport()
    MousePointer = vbHourglass
    If mSQLAfterPrint <> "" Then ExecMe mSQLAfterPrint
    Printer.Line (0, 0)-(1, 0)
    f.SetObj Printer
    If mTotalPage = -1 Then
        PreviewReport Printer
        Printer.EndDoc
        MousePointer = vbDefault
        Exit Sub
    End If
    a = CariSemuaPageSign
    m = IIf(Picture1.Height < 8000, 2, 1)
    For i = 1 To a
        If i <> 1 And (i - 1) Mod m = 0 Then Printer.NewPage
        PreviewReport Printer, i
    Next
    Printer.EndDoc
    MousePointer = vbDefault
End Sub

Private Function InitReport(obj As Object, ByVal tSetData As Boolean) As Boolean
    Set f = Forms.Add("Report" & mReportForm)
    mTotalPage = 0
    If Not tSetData Then
        If mSQL <> "" Then
            f.LoadHeader mParams, obj, mSQL
        Else
            f.LoadHeader mParams, obj
        End If
        InitReport = RS.RecordCount > 0
    Else
        InitReport = True
    End If
    ReDim mPageSign(255)
    VScroll1.Tag = True
End Function

Sub PreviewReport(obj As Object, Optional ByVal tPage As Integer = 1, Optional ByVal tPrintedPage As Integer = 1)
    PaintMe obj, tPage, tPrintedPage
End Sub

Private Sub PrintHead(obj As Object, Optional ByVal tPlus As Single)
On Error Resume Next
    mSetPageHeight = False
    For i = 0 To f.dHead.Count - 1
        PaintHeader f.dHead(i), obj, f.dHead(i), tPlus
    Next
    For i = 0 To f.Lines.Count - 1
        PaintLine obj, f.Lines(i).x1, f.Lines(i).y1 + tPlus, f.Lines(i).x2, f.Lines(i).y2 + tPlus
    Next
    Dim tMax As Byte
    For i = 0 To f.dFoot.Count - 1
        PaintDetail f.dFoot(i), obj, f.dFoot(i), Picture1.Height - f.dFooterLine.y1 + f.dFoot(i).Top, tMax, False
    Next
    For i = 0 To f.LinesFoot.Count - 1
        PaintLine obj, f.LinesFoot(i).x1, Picture1.Height - f.dFooterLine.y1 + f.LinesFoot(i).y1, f.LinesFoot(i).x2, Picture1.Height - f.dFooterLine.y1 + f.LinesFoot(i).y2
    Next
End Sub

Private Sub PaintMe(obj As Object, Optional ByVal tPage As Integer = 1, Optional ByVal tPrintedPage As Single = 1)
On Error Resume Next
Dim j As Long
Dim iFirst As Long
Dim iLast As Long
    If mTotalPage = -1 Then
        PrintHead obj
        f.PrintData
        Exit Sub
    End If
    If tPage > mTotalPage And mTotalPage > 0 Then
        tPage = mTotalPage
    End If
    If tPage = 0 Then tPage = 1
    mPage = tPage - 1
    tPlus = IIf(tPrintedPage Mod 2 = 1 And Picture1.Height < 8000, 7750, 0)
    PrintHead obj, tPlus
    If mPageSign(mPage) > 0 Or mPage = 0 Then
        j = mPage + 1
        If mPageSign(mPage + 1) > 0 Then
            iLast = mPageSign(mPage + 1)
            j = j + 1
        End If
    Else
        j = 1
    End If
    While j <= mPage + 1 And mTotalPage = 0
        While mPageSign(j) > 0
            j = j + 1
        Wend
        mPageSign(j) = f.PrintData(mPageSign(j - 1), 0, j, 0, True)
        j = j + 1
    Wend
    If mPage > mTotalPage And mTotalPage > 0 Then mPage = mTotalPage - 1
    iFirst = mPageSign(mPage)
    iLast = mPageSign(mPage + 1) - 1
    f.PrintData iFirst, iLast, mPage + 1, tPlus
    fPage = mPage + 1
    fPagesPrint = mPage + 1
    PaintHeader "Page " & fPage & IIf(mTotalPage > 0, " of " & mTotalPage, ""), obj, f.dPageNumber, tPlus
End Sub

Private Sub fFirst_Click()
On Error Resume Next
    fPage = 1
    fPage_KeyDown 13, 0
End Sub

Private Sub fLast_Click()
    fPage = 255
    fPage_KeyDown 13, 0
End Sub

Private Sub fNext_Click()
On Error Resume Next
    fPage = fPage + 1
    fPage_KeyDown 13, 0
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Picture2.Height = ScaleHeight - Picture2.Top - 50
    Picture2.Width = ScaleWidth - 2 * Picture2.Left
    VScroll1.Left = Picture2.ScaleWidth - VScroll1.Width
    VScroll1.Height = Picture2.ScaleHeight - VScroll1.Width
    HScroll1.Top = Picture2.Top + Picture2.Height - HScroll1.Height - 650
    HScroll1.Width = Picture2.ScaleWidth - VScroll1.Width
    b = (Picture1.Height - Picture2.ScaleHeight) / 500 + 1
    VScroll1.Enabled = b >= 1
    VScroll1.Max = Int(b)
    b = (Picture1.Width - Picture2.ScaleWidth) / 500 + 1
    HScroll1.Enabled = b >= 1
    HScroll1.Max = Int(b)
End Sub

Private Sub fPage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Picture1.Cls
        PreviewReport Picture1, fPage
        mPage = fPage
    End If
    Picture2.SetFocus
End Sub

Private Sub fPage_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub fPrev_Click()
On Error Resume Next
    fPage = fPage - 1
    fPage_KeyDown 13, 0
End Sub

Private Sub fPrint_Click()
    MousePointer = vbHourglass
    c = fPagesPrint
    a = Split(c, ",")
    f.SetObj Printer
    If mSQLAfterPrint <> "" Then ExecMe mSQLAfterPrint
    Printer.Line (0, 0)-(1, 0)
    If mTotalPage = -1 Then
        PreviewReport Printer
        Printer.EndDoc
        MousePointer = vbDefault
        Exit Sub
    End If
    Dim k As Long
    k = 0
    m = IIf(Picture1.Height < 8000, 2, 1)
    For i = 0 To UBound(a)
        b = Split(a(i), "-")
        If UBound(b) = 0 Then
            If k > 0 And k Mod m = 0 Then Printer.NewPage
            PreviewReport Printer, b(0), k
            k = k + 1
        Else
            For j = b(0) To b(1)
                If k > 0 And (j - 1) Mod m = 0 Then Printer.NewPage
                PreviewReport Printer, j, k
                k = k + 1
            Next
        End If
    Next
    Printer.EndDoc
    If mSQLAfterPrint <> "" Then ExecMe mSQLAfterPrint
    f.SetObj Picture1
    fPagesPrint = c
    MousePointer = vbDefault
End Sub

Private Sub VScroll1_Change()
    If VScroll1.Max = 0 Then
        Picture1.Top = 50
        Exit Sub
    End If
    Picture1.Top = VScroll1.Value / VScroll1.Max * (Picture2.ScaleHeight - Picture1.Height)
    If VScroll1.Value = 0 Then
        Picture1.Top = Picture1.Top + 50
    ElseIf VScroll1.Value = VScroll1.Max Then
        Picture1.Top = Picture1.Top - 300
    End If
End Sub

Private Sub VScrollChange(ByVal v As Integer)
End Sub

Private Sub HScroll1_Change()
    If HScroll1.Max = 0 Then
        Picture1.Left = 50
        Exit Sub
    End If
    Picture1.Left = HScroll1.Value / HScroll1.Max * (Picture2.ScaleWidth - Picture1.Width)
    If HScroll1.Value = 0 Then
        Picture1.Left = Picture1.Left + 50
    ElseIf HScroll1.Value = HScroll1.Max Then
        Picture1.Left = Picture1.Left - 300
    End If
End Sub

Sub LoadFromData(tf As Form, ByVal tReportForm As String, x As XArrayDB, Optional ByVal tParams As String, Optional ByVal tSQLAfterPrint As String, Optional ByVal tPageHeight As Single = 0)
    VScroll1.Tag = False
    mReportForm = tReportForm
    mParams = tParams
    mSQLAfterPrint = tSQLAfterPrint
    mPage = 1
    b = MsgBox("Langsung Print dengan " & Printer.DeviceName & "?", vbYesNoCancel)
    If b = vbNo Then
        Picture1.Cls
        If Not InitReport(Picture1, True) Or x.UpperBound(1) = -1 Then
            MsgBox "No Data"
            Unload Me
            Exit Sub
        End If
        f.SetObj Picture1
        f.SetData x, tParams
        PreviewReport Picture1
        Show
    ElseIf b = vbYes Then
        If Not InitReport(Picture1, True) Or x.UpperBound(1) = -1 Then
            MsgBox "No Data"
            Unload Me
            Exit Sub
        End If
        f.SetData x, tParams
        PrintReport
        If mSQLAfterPrint <> "" Then ExecMe mSQLAfterPrint
    End If
End Sub

Private Sub Form_Load()
    AttachMessage Picture2.hwnd, Me
End Sub

Sub EventModule(ByVal iMsg As Long, ByVal wParam, ByVal lParam, ByVal hwnd As Long)
    If hwnd = Picture2.hwnd Then
        If iMsg = WM_MOUSEWHEEL Then
            If wParam < 0 Then
                If VScroll1.Value < VScroll1.Max Then
                    VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
                    VScroll1_Change
                End If
            Else
                If VScroll1.Value > VScroll1.Min Then
                    VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
                    VScroll1_Change
                End If
            End If
        End If
    End If
End Sub

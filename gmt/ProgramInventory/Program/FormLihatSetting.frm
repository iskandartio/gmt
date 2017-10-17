VERSION 5.00
Begin VB.Form FormLihatSetting 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LIHAT"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Tag             =   "38"
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   8160
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton fDelete 
         Caption         =   "&DELETE"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton fADD 
         Caption         =   "&ADD"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox fAddText 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2535
      End
      Begin VB.ListBox fListLihat 
         Height          =   5325
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.TextBox fLihat2 
      Height          =   6375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "FormLihatSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub


Private Sub fADD_Click()
On Error GoTo err
    i = 0
    Do
        f = False
        For j = 0 To fListLihat.ListCount - 1
            If fListLihat.ItemData(j) = i Then f = True
        Next
        If Not f Then Exit Do Else i = i + 1
    Loop
    a = "insert into m_lihat(lihat1, Tag, Pengupdate) values('" & fAddText.Text & "'," & i & ",'" & pUsr & "')"
    If ExecMe(a, False) = 0 Then GoTo err
    DoQuery
err:
End Sub

Private Sub fDelete_Click()
    a = "delete from m_lihat where lihat1='" & esc(fListLihat.Text) & "'"
    ExecMe a, False
    DoQuery
End Sub

Private Sub fListLihat_Click()
On Error Resume Next
    If fLihat2.Tag <> "" Then
        a = "update m_lihat set Waktu=" & cD(pServerDate) & ",Pengupdate='" & pUsr & "',lihat2='" & Replace(Replace(fLihat2, "'", "''"), vbCrLf, "#") & "' where lihat1='" & fLihat2.Tag & "'"
        b = ExecMe(a, False)
    End If
    fLihat2 = ""
    a = "select top 1 lihat1,lihat2 from m_lihat where lihat1='" & esc(fListLihat.Text) & "'"
    query a
    a = RS.Fields(1).Value
    fLihat2 = Replace(a, "#", vbCrLf)
    fLihat2.Tag = RS.Fields(0).Value
End Sub

Private Sub DoQuery()
On Error GoTo err
    a = "select lihat1, Tag from m_Lihat order by lihat1"
    query a
    fListLihat.Clear
    For i = 0 To RS.RecordCount - 1
        fListLihat.List(i) = RS.Fields(0).Value
        fListLihat.ItemData(i) = RS.Fields(1).Value
        RS.MoveNext
    Next
err:
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    DoQuery
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Left = ScaleWidth - Frame1.Width
    Frame1.Height = ScaleHeight - Frame1.Top
    fListLihat.Height = ScaleHeight - fListLihat.Top
    fLihat2.Width = ScaleWidth - Frame1.Width
    fLihat2.Height = ScaleHeight - fLihat2.Top - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If fLihat2.Tag <> "" Then
        a = "update m_lihat set Waktu=" & cD(pServerDate) & ",Pengupdate='" & pUsr & "', Lihat2='" & Replace(fLihat2, vbCrLf, "#") & "' where lihat1='" & fLihat2.Tag & "'"
        b = ExecMe(a, False)
    End If
End Sub





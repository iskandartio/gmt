VERSION 5.00
Begin VB.Form FormLihat 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin GMT.UsrList UsrList1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   5106
      _ExtentY        =   4895
   End
End
Attribute VB_Name = "FormLihat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub FormHelpKeyDown(ByVal tRetVal As String)
    UsrList1.TextIndex(UsrList1.GetActiveControl) = tRetVal
    UsrList1.FocusSelect
    UsrList1.TextKeyDown UsrList1.GetActiveControl
End Sub

Private Sub Form_Resize()
On Error Resume Next
    UsrList1.Width = ScaleWidth - UsrList1.Left * 2
    UsrList1.Height = ScaleHeight - UsrList1.Top
End Sub
Sub PascaLoading(ByVal tCommand As String)
    
End Sub

Sub LoadMe(Optional ByVal tLihat As String = "", Optional ByVal tWidth As Single = 0, Optional ByVal tHeight As Single = 0)
On Error GoTo err
    UsrList1.SetConn RS
    If tLihat = "" Then
        a = "select lihat2, lihat1 from m_Lihat where tag<31 and (" & pBag3 & "\(2^tag)) mod 2=1" & _
        " union all select lihat2, lihat1 from m_Lihat where tag>=31 and (" & pBag4 & "\(2^tag)) mod 2=1" & _
        " order by Lihat1"
        query a
        s = ""
        For i = 0 To RS.RecordCount - 1
            s = s & "|" & RS.Fields(0).Value
            RS.MoveNext
        Next
    Else
        a = "select lihat2 from " & tLihat & " left join m_lihat on " & tLihat & ".lihat1=m_lihat.lihat1 order by " & tLihat & ".lihat1"
        query a
        s = ""
        For i = 0 To RS.RecordCount - 1
            s = s & "|" & RS.Fields(0).Value
            RS.MoveNext
        Next
    End If
    UsrList1.LoadMe Mid(s, 2), Me
    If tWidth <> 0 Then Width = tWidth
    If tHeight <> 0 Then Height = tHeight
    Show
err:
End Sub

Private Sub UsrList1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove Button, Me
End Sub



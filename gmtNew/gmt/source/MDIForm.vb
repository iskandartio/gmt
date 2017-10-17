Imports System.Reflection

Public Class MDIForm

    Private Sub MDIForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim f As New frmLogin
        If f.ShowDialog(Me) <> Windows.Forms.DialogResult.Yes Then
            Me.Dispose()
        End If
    End Sub

    Private Sub CustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerToolStripMenuItem.Click
        Dim f As New frmMCustomer
        f.MdiParent = Me
        f.Show()

    End Sub
End Class
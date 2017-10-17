Public Class frmMLotMapping
    Inherits FormMain
    Dim a As New SQLmLotMapping
    Dim dt As DataTable
    Dim dtPenerima As DataTable
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If a.UpdateData(dt) Then
            MsgBox("Sukses")
        Else
            MsgBox("Failed")
        End If
    End Sub

    Private Sub frmMCustomer_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If dt Is Nothing Then
            dt = a.qData()
            dg.DataSource = dt
            ModControl.setGridVisible(dg, False, "rowid")
        End If

    End Sub

    Private Sub txtFind_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyDown

        If e.KeyCode = Keys.Enter Then
            dg.Find(txtFind.Text)
            txtFind.Focus()
        End If
    End Sub


End Class
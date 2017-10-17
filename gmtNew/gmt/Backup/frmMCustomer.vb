Public Class frmMCustomer
    Inherits FormMain
    Dim a As New SQLmCustomer
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

        End If

    End Sub
    Private Sub GetDetail(ByVal tRow As Integer)

        dtPenerima = a.qDataPenerima(dg.Rows(tRow).Cells("CustomerID").Value)
        dgPenerima.DataSource = dtPenerima
    End Sub

    Private Sub txtFind_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyDown

        If e.KeyCode = Keys.Enter Then
            dg.Find(txtFind.Text)
            txtFind.Focus()
        End If
    End Sub


    Private Sub dg_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.RowEnter
        GetDetail(e.RowIndex)
    End Sub

    Private Sub btnUpdatePenerima_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdatePenerima.Click
        If a.UpdatePenerima(dtPenerima) Then
            MsgBox("Sukses")
        Else
            MsgBox("Failed")
        End If
    End Sub
End Class
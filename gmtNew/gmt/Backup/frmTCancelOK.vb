Public Class frmTCancelOK
    Dim mdt As New DataTable
    Dim bs As New BindingSource
    Dim a As New SQLtInputStock
    Private Sub btnTipe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTipe.Click
        If lblTipe.Text = "DTY" Then
            lblTipe.Text = "PE"
            pTipe = "PE"
        Else
            lblTipe.Text = "DTY"
            pTipe = "DTY"
        End If
        Search()
    End Sub

    Private Sub Search()
        dg.AutoGenerateColumns = False
        mdt = a.qDataOKPacking(dtTgl.Value)
        If mdt.Rows.Count = 0 Then
            mdt.Rows.Add(mdt.NewRow)
            mdt.Rows(0)("NoBukti") = 0
        End If
        bs.DataSource = mdt
        SetDataSource(dg, mdt, bs, "", "@CANCEL")
        setGridVisible(dg, False, "Tanggal", "IDTrans")
        ModControl.setGridReadOnly(dg, True, "0")
        ModControl.setGridReadOnly(dg, False, "CANCEL")
        ModControl.SetBinding(Me, bs)
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Search()
    End Sub
    Private Sub frmTCancelOK_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pTipe = "DTY"
        dtTgl.Value = Now.Date
        
    End Sub
    Private Sub dg_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.CellContentClick
        If e.RowIndex = -1 Then Exit Sub
        If e.ColumnIndex = 0 Then
            a.CancelPacking(dg.Rows(e.RowIndex).Cells("IDTrans").Value)
            dg.Rows.RemoveAt(e.RowIndex)
        End If
        
    End Sub

End Class
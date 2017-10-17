Public Class frmRInputStockNotOK
    Dim mdt As New DataTable
    Dim bs As New BindingSource
    Dim a As New SQLtInputStock
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Search()
    End Sub

    Private Sub Search()
        dg.AutoGenerateColumns = False
        mdt = a.qDataNotOK()
        If mdt.Rows.Count = 0 Then
            mdt.Rows.Add(mdt.NewRow)
            mdt.Rows(0)("NoBukti") = 0
        End If
        For Each row As DataRow In mdt.Rows
            row("Tanggal") = ModFunction.convertToDate(row("Tanggal"))
            row("TglGudang") = ModFunction.convertToDate(row("TglGudang"))
        Next
        bs.DataSource = mdt
        SetDataSource(dg, mdt, bs, "@Status")
        setGridReadOnly(dg, True, "0")
        ModControl.SetBinding(Me, bs)
    End Sub

    Private Sub frmtInputStockOK_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pTipe = "DTY"

    End Sub

    Private Sub dg_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.CellContentClick
        Try
            If dg.Columns(e.ColumnIndex).Name = "OK" Then
                dg.Rows.Remove(dg.Rows(e.RowIndex))
            ElseIf dg.Columns(e.ColumnIndex).Name = "CANCEL" Then
                a.CancelGudang(dg.Rows(e.RowIndex).Cells("IDTrans").Value)
                dg.Rows.Remove(dg.Rows(e.RowIndex))
            End If
        Catch ex As Exception

        End Try

    End Sub

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

End Class
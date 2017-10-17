Public Class frmtInputStockOKPacking
    Dim mdt As New DataTable
    Dim bs As New BindingSource
    Dim a As New SQLtInputStock
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Search()
    End Sub

    Private Sub Search()
        dg.AutoGenerateColumns = False
        mdt = a.qDataPrintedPacking(dtTgl.Value)
        If mdt.Rows.Count = 0 Then
            mdt.Rows.Add(mdt.NewRow)
            mdt.Rows(0)("NoBukti") = 0
        End If
        bs.DataSource = mdt
        SetDataSource(dg, mdt, bs, "@Check@Status")
        setGridVisible(dg, False, "Tanggal")
        ModControl.SetBinding(Me, bs)
        ModControl.setUpdateOnly(dg)
        ModControl.setGridReadOnly(dg, True, "0")
        ModControl.setGridReadOnly(dg, False, "Check")
    End Sub

    Private Sub frmtInputStockOK_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pTipe = "DTY"
        dtTgl.Value = Now.Date
        dtTglGudang.Value = Now.Date
        
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

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If a.OKPacking(mdt, dtTglGudang.Value) Then
            MsgBox("Sukses")
        Else
            MsgBox("Failed")
        End If
    End Sub

    Private Sub btnPilih_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPilih.Click
        If btnPilih.Text = "Pilih Semua" Then
            For i As Integer = 0 To dg.Rows.Count - 1
                dg.Rows(i).Cells("Check").Value = 1
                btnPilih.Text = "Hapus Pilihan"
            Next
        Else
            For i As Integer = 0 To dg.Rows.Count - 1
                dg.Rows(i).Cells("Check").Value = 0
                btnPilih.Text = "Pilih Semua"
            Next
        End If
        

    End Sub
End Class

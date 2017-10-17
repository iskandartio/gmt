Public Class frmTSTT
    Dim mFirstLoad As Boolean = True
    Dim a As New SQLSTT
    Dim mNew As Boolean
    Dim mReadTime As DateTime
    Dim mdt As DataTable
    Dim mdtPembayaran As DataTable
    Dim mdtPelunasan As DataTable
    Private WithEvents ctrl As TextBox

    Private Sub _Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If mFirstLoad Then
            dbCustomerID.DataSource = New SQLmCustomer().qData(setParam("IsActive", 1))
            dbCustomerID.DisplayMember = "NamaCustomer"
            dbCustomerID.ValueMember = "CustomerID"
            dbCustomerID.SelectedIndex = -1
            ModControl.LockMe(dbPengupdate, True)
            ModControl.LockMe(dbWaktuUpdate, True)
            NewData()
            mFirstLoad = False
        End If
    End Sub

    Private Sub GenerateNumber()
        Dim dt As DataTable = New SQLPenomoran().qData("STT")
        dbNoSTT.Text = ("00000" & dbNoSTT.Tag).Substring(Len(dbNoSTT.Tag.ToString)) & dt.Rows(0)("Penomoran") & Format(dbTglSTT.Value, "MM/yy")
    End Sub

    Private Sub dbTgl_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbTglSTT.TextChanged
        If Not dbNoSTT.Text.EndsWith(Format(dbTglSTT.Value, "yy")) Then
            dbNoSTT.Tag = a.qNo(dbTglSTT.Value)
            GenerateNumber()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        BindToDataTable(Me)
        If a.Save(mdt, mdtPembayaran, mdtPelunasan) Then
            btnSave.Enabled = Not btnSave.Enabled
            btnCancel.Enabled = Not btnCancel.Enabled
            MsgBox("Sukses")
        Else
            MsgBox("Gagal")
        End If
    End Sub

    Private Sub DoQuery()
        Dim z As New List(Of Param)
        Dim t() As String
        t = Split(fQuickFind.Text, "/")
        z.Add(New Param("toint(left(NoSTT,5))", "=" & t(0), True))
        If t.Length > 1 Then
            z.Add(New Param("year(TglSTT)", "=" & 2000 + t(1), True))
        Else
            z.Add(New Param("year(TglSTT)", "=" & Year(Now.Date), True))
        End If
        mdt = a.qData(z)
        BindMe()
        If mdt.Rows.Count = 0 Then
            dgPembayaran.Columns.Clear()
            dgPelunasan.Columns.Clear()
            Exit Sub
        End If
        mdtPembayaran = a.qDataPembayaran(setParam("a.NoSTT", mdt.Rows(0)("NoSTT")))
        ModControl.EnableControl(dgPembayaran, False)
        mdtPelunasan = a.qDataPelunasan(setParam("a.NoSTT", mdt.Rows(0)("NoSTT")))
        ModControl.EnableControl(dgPelunasan, False)
        ModControl.setGridDataSource(dgPembayaran, mdtPembayaran)
        ModControl.setGridDataSource(dgPelunasan, mdtPelunasan)
        ModControl.setGridStyle(dgPembayaran)
        ModControl.setGridStyle(dgPelunasan)
        btnSave.Enabled = False
        btnCancel.Enabled = True
    End Sub
    Private Sub BindMe()
        ClearBinding(Me)
        SetBinding(Me, mdt)

    End Sub
    Private Sub fQuickFind_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles fQuickFind.KeyDown
        If e.KeyCode = Keys.Enter Then
            DoQuery()
        End If
    End Sub


    Private Sub btnPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrev.Click
        SetCurrent(-1)
    End Sub
    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        SetCurrent(1)
    End Sub
    Private Sub SetCurrent(ByVal Plus As Integer)
        Dim t() As String
        t = Split(fQuickFind.Text, "/")
        If t(0) + Plus < 1 Then Exit Sub
        t(0) = t(0) + Plus
        If t.Length > 1 Then
            fQuickFind.Text = t(0) & "/" & t(1)
        Else
            fQuickFind.Text = t(0)
        End If
        DoQuery()
    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If a.Cancel(dbNoSTT.Text) Then
            MsgBox("Sukses")
            ModControl.setDTReadOnly(mdt, False, "0")
            ModControl.EnableControl(dgPembayaran, True)
            ModControl.EnableControl(dgPelunasan, True)
            BindMe()
            btnSave.Enabled = True
            btnCancel.Enabled = False
        Else
            MsgBox("Gagal")
        End If
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        NewData()
    End Sub
    Private Sub NewData()
        mdt = a.qStructure()
        mdt.Rows.Add(mdt.NewRow)
        mdt.Rows(0)("TglSTT") = Now.Date
        mdtPembayaran = a.qStructurePembayaran
        ModControl.setGridDataSource(dgPembayaran, mdtPembayaran)
        ModControl.setGridStyle(dgPembayaran)
        ModControl.EnableControl(dgPembayaran, True)
        ModControl.setGridReadOnly(dgPembayaran, True, "NilaiRp")
        ModControl.EnableControl(dgPelunasan, True)
        dbNoSTT.Tag = a.qNo(dbTglSTT.Value)
        BindMe()
        btnCancel.Enabled = False
        btnSave.Enabled = True
        GenerateNumber()
    End Sub

    Private Sub dbCustomerID_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbCustomerID.SelectionChangeCommitted
        mdtPelunasan = a.qBelumLunas(setParam("a.CustomerID", "=" & dbCustomerID.SelectedValue))
        ModControl.setGridDataSource(dgPelunasan, mdtPelunasan)
        ModControl.setGridStyle(dgPelunasan, "0")

    End Sub

    Private Sub dgPembayaran_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgPembayaran.CellEndEdit
        Try
            dgPembayaran.Item("NilaiRp", e.RowIndex).Value = 0
            If dgPembayaran.Item("MataUang", e.RowIndex).Value.ToString.ToUpper = "RP" Then
                dgPembayaran.Item("Kurs", e.RowIndex).Value = 1
            End If
            dgPembayaran.Item("NilaiRp", 0).Value = dgPembayaran.Item("Kurs", 0).Value * dgPembayaran.Item("Nilai", 0).Value
        Catch ex As Exception

        End Try
        
    End Sub
End Class
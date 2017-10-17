Public Class frmTSC
    Dim mFirstLoad As Boolean = True
    Dim a As New SQLSC
    Dim mNew As Boolean
    Dim mReadTime As DateTime
    Dim bs As New BindingSource
    Dim mdt As DataTable
    Dim mdt2 As DataTable
    Private WithEvents ctrl As TextBox

    Private Sub frmTSC_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If mFirstLoad Then
            NewData()
            dbCustomerID.DataSource = New SQLmCustomer().qData(setParam("IsActive", 1))
            dbCustomerID.DisplayMember = "NamaCustomer"
            dbCustomerID.ValueMember = "CustomerID"
            dbCustomerID.SelectedIndex = -1
            dbMataUang.DataSource = New SQLmMataUang().qData
            dbMataUang.DisplayMember = "Curr"
            dbMataUang.ValueMember = "Curr"
            ModControl.LockMe(dbPengupdate, True)
            ModControl.LockMe(dbWaktuUpdate, True)
            ModControl.LockMe(dbNilaiKontrak, True)
            ModControl.LockMe(txtStatus, True)
            mFirstLoad = False
        End If
    End Sub
    Private Sub NewData()
        mdt = a.qStructure
        mdt.Rows.Add(mdt.NewRow)
        mdt.Rows(0)("TglSC") = Now.Date
        BindMe()
        dbNoSC.Tag = a.qNo(dbTglSC.Value)
        GenerateNumber()
        fQuickFind.Text = dbNoSC.Tag
        mReadTime = Nothing
    End Sub
    Private Sub GenerateNumber()
        Dim dt As DataTable = New SQLPenomoran().qData("SC")
        dbNoSC.Text = ("00000" & dbNoSC.Tag).Substring(Len(dbNoSC.Tag.ToString)) & dt.Rows(0)("Penomoran") & Format(dbTglSC.Value, "MM/yy")
    End Sub

    Private Sub dbTglSC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbTglSC.TextChanged
        If Not dbNoSC.Text.EndsWith(Format(dbTglSC.Value, "yy")) Then
            dbNoSC.Tag = a.qNo(dbTglSC.Value)
            GenerateNumber()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        bs.EndEdit()
        BindToDataTable(Me)
        If a.Save(mdt, mdt2, mReadTime) Then
            MsgBox("Sukses")
        Else
            MsgBox("Gagal")
        End If
    End Sub
    Private Sub BindMe()
        ClearBinding(Me)
        SetBinding(Me, mdt)
        mdt2 = a.qDataDetail(setParam("NoSC", dbNoSC.Text))
        dg.DataSource = mdt2
        ModControl.setGridAutoInvisible(dg)
        ModControl.setGridStyle(dg, 0)
    End Sub
    Private Sub DoQuery()
        Dim z As New List(Of Param)
        Dim t() As String
        t = Split(fQuickFind.Text, "/")
        z.Add(New Param("toint(left(NoSC,5))", "=" & t(0), True))
        If t.Length > 1 Then
            z.Add(New Param("year(TglSC)", "=" & 2000 + t(1), True))
        Else
            z.Add(New Param("year(TglSC)", "=" & Year(Now.Date), True))
        End If
        mdt = a.qData(z)
        bs.DataSource = mdt
        If Not mdt Is Nothing And mdt.Rows.Count <> 0 Then
            mReadTime = mdt.Rows(0)("WaktuUpdate")
            If mdt.Rows(0)("Status") = 0 Then
                txtStatus.Text = "BELUM DISETUJUI"
            ElseIf mdt.Rows(0)("Status") = 1 Then
                txtStatus.Text = "DISETUJUI"
            ElseIf mdt.Rows(0)("Status") = 2 Then
                txtStatus.Text = "CLOSED"
            End If
            btnApprove.Enabled = (mdt.Rows(0)("Status") = 0)
            btnSave.Enabled = (mdt.Rows(0)("Status") = 0)
            ModControl.LockMe(dg, (mdt.Rows(0)("Status") > 0))
        End If
        BindMe()
    End Sub

    Private Sub fQuickFind_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles fQuickFind.KeyDown
        If e.KeyCode = Keys.Enter Then
            DoQuery()

        End If
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        NewData()
    End Sub

    Private Sub dg_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.CellEndEdit
        ctrl = Nothing
        Dim col = e.ColumnIndex
        Dim tTipe As String = dg.Columns(dg.CurrentCell.ColumnIndex).Name
        If Not (tTipe = "Jenis" Or tTipe = "KodeBarang" Or tTipe = "Warna" Or tTipe = "NoWarna" Or tTipe = "Tube" Or tTipe = "Grade") Then
            HitungTotal()

            Exit Sub
        End If
        dg.CurrentCell.Value = setCurrentValue(New SQLmStock().qDataSingle(dg.Columns(col).Name, setParam(dg.Columns(col).Name, dg.CurrentCell.Value)))
    End Sub
    Private Sub HitungTotal()
        Try
            Dim tTotal As Double = 0
            For Each row As DataGridViewRow In dg.Rows
                tTotal = tTotal + row.Cells("Jumlah").Value * row.Cells("Harga").Value
            Next
            dbNilaiKontrak.Value = tTotal
        Catch ex As Exception
            dbNilaiKontrak.Value = 0
        End Try
    End Sub

    Private Sub dg_EditingControlShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dg.EditingControlShowing
        ctrl = e.Control
    End Sub

    Private Sub _KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ctrl.KeyDown, dg.KeyDown
        If e.KeyCode = Keys.F1 Then
            If dg.Rows.Count = 1 Then
                SendKeys.SendWait(" " & Chr(8))
            End If
            Dim tVal As String
            Dim tTipe As String
            tTipe = dg.Columns(dg.CurrentCell.ColumnIndex).Name
            If Not (tTipe = "Jenis" Or tTipe = "KodeBarang" Or tTipe = "Warna" Or tTipe = "NoWarna" Or tTipe = "Tube" Or tTipe = "Grade") Then
                Exit Sub
            End If
            If ctrl Is Nothing Then
                tVal = dg.CurrentCell.Value.ToString
            Else
                tVal = ctrl.Text
            End If
            Dim f As New FormPilihSingle(tTipe, tVal)
            Dim dlg As DialogResult
            dlg = f.ShowDialog(Me)
            If dlg = Windows.Forms.DialogResult.OK Then
                dg.EndEdit()
                dg.CurrentCell.Value = f.row(tTipe)
            End If
            f.Dispose()
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

    Private Sub dg_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.UserDeletedRow
        HitungTotal()
    End Sub

    Private Sub btnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        a.SetStatus(dbNoSC.Text, 1)
        DoQuery()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        dbKeterangan.Text = "test lagi deh"
        'a.SetStatus(dbNoSC.Text, 2)
        'DoQuery()
    End Sub
End Class
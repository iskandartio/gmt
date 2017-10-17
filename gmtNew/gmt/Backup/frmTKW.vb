Public Class frmTKW
    Dim mFirstLoad As Boolean = True
    Dim a As New SQLKW
    Dim mNew As Boolean
    Dim mReadTime As DateTime
    Dim mdt As DataTable
    Dim mdt2 As DataTable
    Private WithEvents ctrl As TextBox
    
    Private Sub _Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If mFirstLoad Then

            dbMataUang.DataSource = New SQLmMataUang().qData
            dbMataUang.DisplayMember = "Curr"
            dbMataUang.ValueMember = "Curr"
            dbMataUang.SelectedIndex = -1
            ModControl.LockMe(dbPengupdate, True)
            ModControl.LockMe(dbWaktuUpdate, True)
            ModControl.LockMe(txtNilai, True)
            dbNoKW.Tag = a.qNo(dbTglKW.Value)
            GenerateNumber()
            ModControl.setUpdateOnly(dg)
            ModControl.setUpdateOnly(dgDetail)

            mFirstLoad = False
        End If
    End Sub

    Private Sub GenerateNumber()
        Dim dt As DataTable = New SQLPenomoran().qData("KW")
        dbNoKW.Text = ("00000" & dbNoKW.Tag).Substring(Len(dbNoKW.Tag.ToString)) & dt.Rows(0)("Penomoran") & Format(dbTglKW.Value, "MM/yy")
    End Sub

    Private Sub dbTgl_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbTglKW.TextChanged
        If Not dbNoKW.Text.EndsWith(Format(dbTglKW.Value, "yy")) Then
            dbNoKW.Tag = a.qNo(dbTglKW.Value)
            GenerateNumber()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If a.Save(mdt, setParam("NoKW", dbNoKW.Text, "TglKW", dbTglKW.Value)) Then
            MsgBox("Sukses")
        Else
            MsgBox("Gagal")
        End If
    End Sub

    Private Sub DoQuery()
        Dim z As New List(Of Param)
        Dim t() As String
        t = Split(fQuickFind.Text, "/")
        z.Add(New Param("toint(left(NoKW,5))", "=" & t(0), True))
        If t.Length > 1 Then
            z.Add(New Param("year(TglKW)", "=" & 2000 + t(1), True))
        Else
            z.Add(New Param("year(TglKW)", "=" & Year(Now.Date), True))
        End If
        mdt = a.qData(z)
        ModControl.EnableControl(dg, False)
        ModControl.setDTReadOnly(mdt, True, "0")
        dg.DataSource = mdt
        ModControl.setGridStyle(dg, 0)
        ModControl.setGridAutoInvisible(dg)
        ModControl.setGridVisible(dg, False, "0")
        ModControl.setGridVisible(dg, True, "Pilih", "NoSJ", "TglSJ", "NamaPenerima", "Total")
        If Not mdt Is Nothing And mdt.Rows.Count <> 0 Then
            mReadTime = mdt.Rows(0)("WaktuUpdate")
            Dim dt As DataTable = New SQLmCustomer().qData(setParam("CustomerID", "=" & mdt.Rows(0)("CustomerID")))
            dbCustomerID.DataSource = dt
            dbCustomerID.DisplayMember = "NamaCustomer"
            dbCustomerID.ValueMember = "CustomerID"

            btnSave.Enabled = False
            btnCancel.Enabled = True
        End If
        BindMe()
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


    Private Sub dg_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.CellEndEdit
        HitungTotal()
    End Sub
    Private Sub HitungTotal()
        Dim tTotal As Double = 0
        Try

            For Each row As DataGridViewRow In dg.Rows
                If row.Cells("Pilih").Value Then
                    tTotal = tTotal + row.Cells("Total").Value
                End If

            Next
            txtNilai.Value = tTotal
        Catch ex As Exception
            txtNilai.Value = tTotal
        End Try
    End Sub

    Private Sub dg_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.CellValueChanged

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

    Private Sub dg_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.RowEnter
        GetDetail(e.RowIndex)
    End Sub

    

    Private Sub dbMataUang_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbMataUang.SelectionChangeCommitted
        dbCustomerID.DataSource = New SQLKW().qCustomer(dbMataUang.SelectedValue)
        dbCustomerID.DisplayMember = "NamaCustomer"
        dbCustomerID.ValueMember = "CustomerID"
        dbCustomerID.SelectedIndex = -1
    End Sub

    
    Private Sub GetDetail(ByVal tRow As Integer)
        mdt2 = a.qDetailSPP(dg.Rows(tRow).Cells("NoSJ").Value)
        dgDetail.DataSource = mdt2
        ModControl.setGridAutoInvisible(dgDetail)
        ModControl.setGridStyle(dgDetail, 0)
        ModControl.LockMe(dgDetail, True)
    End Sub

    Private Sub dbCustomerID_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbCustomerID.SelectionChangeCommitted
        mdt = a.qSPP(dbCustomerID.SelectedValue, dbMataUang.SelectedValue)
        ModControl.setDTReadOnly(mdt, True, "0")
        ModControl.setDTReadOnly(mdt, False, "Pilih")
        ModControl.setGridDataSource(dg, mdt)
        
        ModControl.setGridReadOnly(dg, False, "Pilih")
        ModControl.setGridStyle(dg, 0)
        

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If a.Cancel(dbNoKW.Text) Then
            MsgBox("Sukses")
            ModControl.setDTReadOnly(mdt, False, "0")
            ModControl.EnableControl(dg, True)
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
        ModControl.EnableControl(dg, True)
        ModControl.EnableControl(dgDetail, True)
        
        BindMe()
    End Sub

    Private Sub frmTKW_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
End Class
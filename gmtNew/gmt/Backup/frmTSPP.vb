Public Class frmTSPP
    Dim mFirstLoad As Boolean = True
    Dim a As New SQLSPP
    Dim mNew As Boolean
    Dim mReadTime As DateTime
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
            ModControl.LockMe(txtStatus, True)
            mFirstLoad = False
        End If
    End Sub
    Private Sub NewData()
        mdt = a.qStructure
        mdt.Rows.Add(mdt.NewRow)
        mdt.Rows(0)("TglSPP") = Now.Date
        mdt.Rows(0)("TglKirim") = Now.Date
        BindMe()
        dbNoSPP.Tag = a.qNo(dbTglSPP.Value)
        FillPenerima()
        GenerateNumber()
        fQuickFind.Text = dbNoSPP.Tag
        mReadTime = Nothing
        
    End Sub
    Private Sub GenerateNumber()
        Dim dt As DataTable = New SQLPenomoran().qData("SPP")
        dbNoSPP.Text = ("00000" & dbNoSPP.Tag).Substring(Len(dbNoSPP.Tag.ToString)) & dt.Rows(0)("Penomoran") & Format(dbTglSPP.Value, "MM/yy")
    End Sub

    Private Sub _TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbTglSPP.TextChanged
        If Not dbNoSPP.Text.EndsWith(Format(dbTglSPP.Value, "yy")) Then
            dbNoSPP.Tag = a.qNo(dbTglSPP.Value)
            GenerateNumber()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
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
        cmbNamaPenerima.DataBindings.Add("text", mdt, "NamaPenerima")
        mdt2 = a.qDataDetail(dbNoSPP.Text)
        dg.DataSource = mdt2
        ModControl.setGridAutoInvisible(dg)
        ModControl.setGridStyle(dg, 0)
    End Sub
    Private Sub DoQuery()
        Dim z As New List(Of Param)
        Dim t() As String
        t = Split(fQuickFind.Text, "/")
        If t.Length > 1 Then
            z.Add(New Param("year(TglSPP)", "=" & 2000 + t(1), True))
        Else
            z.Add(New Param("year(TglSPP)", "=" & Year(Now.Date), True))
        End If
        z.Add(New Param("toint(left(NoSPP,5))", "=" & t(0), True))
        mdt = a.qData(z)
        If Not mdt Is Nothing And mdt.Rows.Count <> 0 Then
            mReadTime = mdt.Rows(0)("WaktuUpdate")
            If mdt.Rows(0)("Status") = 0 Then
                txtStatus.Text = "BELUM DISETUJUI"
            ElseIf mdt.Rows(0)("Status") = 1 Then
                txtStatus.Text = "DISETUJUI"
            ElseIf mdt.Rows(0)("Status") = 2 Then
                txtStatus.Text = "SUDAH SJ"
            ElseIf mdt.Rows(0)("Status") = 3 Then
                txtStatus.Text = "SUDAH KW"
            End If
            btnApprove.Enabled = Not (mdt.Rows(0)("Status") > 1)
            btnSave.Enabled = Not (mdt.Rows(0)("Status") > 1)
            ModControl.LockMe(dg, (mdt.Rows(0)("Status") > 1))
            Panel1.Enabled = True
        Else
            ClearControlMe()
        End If
        BindMe()
    End Sub

    Private Sub ClearControlMe()

        Panel1.Enabled = False
        
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
        'Dim col = e.ColumnIndex
        'Dim tTipe As String = dg.Columns(dg.CurrentCell.ColumnIndex).Name
        'If Not (tTipe = "Jenis" Or tTipe = "KodeBarang" Or tTipe = "Warna" Or tTipe = "NoWarna" Or tTipe = "Tube" Or tTipe = "Grade") Then
        '    Exit Sub
        'End If

        'dg.CurrentCell.Value = setCurrentValue(New SQLmStock().qDataSingle(dg.Columns(col).Name, setParam(dg.Columns(col).Name, dg.CurrentCell.Value.ToString)))

    End Sub

    Private Sub dg_EditingControlShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dg.EditingControlShowing
        ctrl = e.Control

    End Sub

    Private Sub _KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ctrl.KeyDown, dg.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim f As FormPilih
            Dim dlg As DialogResult
            If dg.Rows.Count = 1 Then
                SendKeys.SendWait(" " & Chr(8))
            End If
            Dim tVal As String
            Dim tTipe As String
            tTipe = dg.Columns(dg.CurrentCell.ColumnIndex).Name
            If tTipe = "NoSC" Then
                dg.EndEdit()
                f = New FormPilihSC(" and b.CustomerID=" & dbCustomerID.SelectedValue)
                dlg = f.ShowDialog(Me)
                If dlg = Windows.Forms.DialogResult.OK Then
                    dg.EndEdit()
                    dg.CurrentRow.Cells("IDSC").Value = f.row("IDTrans")
                    dg.CurrentCell.Value = f.row("NoSC")
                    dg.CurrentRow.Cells("Jenis").Value = f.row("Jenis")
                    dg.CurrentRow.Cells("KodeBarang").Value = f.row("KodeBarang")
                    dg.CurrentRow.Cells("Warna").Value = f.row("Warna")
                    dg.CurrentRow.Cells("NoWarna").Value = f.row("NoWarna")
                    dg.CurrentRow.Cells("Tube").Value = f.row("Tube")
                    dg.CurrentRow.Cells("Grade").Value = f.row("Grade")
                    dg.CurrentRow.Cells("Harga").Value = f.row("Harga")
                    dg.AutoResizeColumns(DataGridViewAutoSizeColumnMode.DisplayedCells)
                End If
                f.Dispose()
            ElseIf (tTipe = "Jenis" Or tTipe = "KodeBarang" Or tTipe = "Warna" Or tTipe = "NoWarna" Or tTipe = "Tube" Or tTipe = "Grade") Then
                dg.IsOnPilih = True
                dg.EndEdit()
                f = New FormPilihStock(getParam(dg.CurrentCell.RowIndex), "Jenis, KodeBarang,Warna,NoWarna, Tube, Grade, SatBesar, JumlahBox, JumlahKG")
                dlg = f.ShowDialog(Me)
                dg.IsOnPilih = False
                If dlg = Windows.Forms.DialogResult.OK Then
                    dg.EndEdit()
                    dg.CurrentRow.Cells("Jenis").Value = f.row("Jenis")
                    dg.CurrentRow.Cells("KodeBarang").Value = f.row("KodeBarang")
                    dg.CurrentRow.Cells("Warna").Value = f.row("Warna")
                    dg.CurrentRow.Cells("NoWarna").Value = f.row("NoWarna")
                    dg.CurrentRow.Cells("Tube").Value = f.row("Tube")
                    dg.CurrentRow.Cells("Grade").Value = f.row("Grade")
                    dg.AutoResizeColumns(DataGridViewAutoSizeColumnMode.DisplayedCells)
                End If
                f.Dispose()
            End If
            If ctrl Is Nothing Then
                tVal = dg.CurrentCell.Value.ToString
            Else
                tVal = ctrl.Text
            End If


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



    Private Sub FillPenerima()
        Try
            Dim dt2 As DataTable
            dt2 = New SQLmCustomer().qDataPenerima(dbCustomerID.SelectedValue)
            If dt2.Rows.Count = 0 Then
                dt2 = New SQLmCustomer().qData(setParam("CustomerID", "=" & dbCustomerID.SelectedValue))
                cmbNamaPenerima.DataSource = dt2
                cmbNamaPenerima.DisplayMember = "NamaCustomer"
                cmbNamaPenerima.ValueMember = "CustomerID"
                cmbNamaPenerima.SelectedIndex = 0
                dbAlamatPenerima.Text = cmbNamaPenerima.SelectedItem("Alamat")
            Else
                cmbNamaPenerima.DataSource = dt2
                cmbNamaPenerima.DisplayMember = "Nama"
                cmbNamaPenerima.ValueMember = "PenerimaID"
                dbAlamatPenerima.Text = cmbNamaPenerima.SelectedItem("Alamat")
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub dbCustomerID_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbCustomerID.SelectionChangeCommitted
        FillPenerima()
    End Sub


    Private Sub cmbNamaPenerima_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbNamaPenerima.SelectionChangeCommitted
        dbAlamatPenerima.Text = cmbNamaPenerima.SelectedItem("Alamat")
    End Sub

    Private Sub dg_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.RowLeave
        If dg.Rows(e.RowIndex).IsNewRow OrElse dg.IsOnPilih Then Exit Sub
        Dim dt As DataTable = New SQLmStock().qStock(getParam(e.RowIndex))
        If dt.Rows.Count = 0 Then
            MsgBox("Nama Stock belum diinput")
            Exit Sub
        End If
        dg.Rows(e.RowIndex).Cells("IDStock").Value = dt.Rows(0)("IDStock")

    End Sub

    Private Function getParam(ByVal row As Integer) As List(Of Param)
        Return setParam("Jenis", "=" & Trim(dg.Rows(row).Cells("Jenis").Value.ToString), _
                                                            "KodeBarang", "=" & Trim(dg.Rows(row).Cells("KodeBarang").Value.ToString), _
                                                            "Warna", "=" & Trim(dg.Rows(row).Cells("Warna").Value.ToString), _
                                                            "NoWarna", "=" & Trim(dg.Rows(row).Cells("NoWarna").Value.ToString), _
                                                            "Tube", "=" & Trim(dg.Rows(row).Cells("Tube").Value.ToString), _
                                                            "Grade", "=" & Trim(dg.Rows(row).Cells("Grade").Value.ToString) _
                                                            )
    End Function
    
    Private Sub btnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        a.SetStatus(dbNoSPP.Text, 1)
        DoQuery()
    End Sub


End Class
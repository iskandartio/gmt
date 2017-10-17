Public Class frmRInputStockAll
    Dim mdt As DataTable
    Private WithEvents bs As New BindingSource
    Dim mdtDetail As DataTable
    Dim bsDetail As New BindingSource
    Dim a As New SQLtInputStock()
    Private WithEvents ctrl As TextBox
    Private WithEvents prt As New Printing.PrintDocument
    Dim dtPrint As DataTable


    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Search()
    End Sub

    Private Sub Search()
        dg.AutoGenerateColumns = False
        mdt = a.qDataAll(dtTgl.Value)
        If mdt.Rows.Count = 0 Then
            NewClick()
        End If
        bs.DataSource = mdt
        SetDataSource(dg, mdt, bs, "@Status")
        setGridVisible(dg, False, "Tanggal")
        ModControl.SetBinding(Me, bs)
    End Sub
    Private Sub GetDetail(ByVal tNoBukti As Integer)
        Try
            mdtDetail = a.qDetail(tNoBukti)
            SetDataSource(dgDetail, mdtDetail, bsDetail, "")
            ModControl.setGridVisible(dgDetail, False, "NoBukti")
            Dim dbStock As New SQLmStock
            If mdt.Rows(bs.Position).RowState = DataRowState.Added Then
                a.mCurrentIDStock = dbStock.qStock(setParam("Jenis", "=" & dbJenis.Text _
                                                        , "KodeBarang", "=" & dbKodeBarang.Text _
                                                        , "Warna", "=" & dbWarna.Text _
                                                        , "NoWarna", "=" & dbNoWarna.Text _
                                                        , "Tube", "=" & dbTube.Text _
                                                        , "Grade", "=" & dbGrade.Text _
                                                        , "SatBesar", "=" & dbSatBesar.Text _
                                                        , "SatKecil", "=" & dbSatKecil.Text _
                                                        ), "", "IDStock").Rows(0)(0)
            Else
                a.mCurrentIDStock = mdt.Rows(bs.Position)("IDStock")
            End If
        Catch ex As Exception

        End Try


    End Sub

    Private Sub frmtInputStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ActiveControl.Name = "txtKg" Then

        End If
    End Sub

    Private Sub frmtInputStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pTipe = "DTY"
        dg.AllowUserToDeleteRows = False
        ModControl.EnableControl(dg, False)

        dtTgl.Value = Now.Date
        Search()
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

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        AddRow()
    End Sub

    Private Sub AddRow()
        Dim r As DataRow
        r = mdtDetail.NewRow
        r("NoBukti") = dbNoBukti.Text
        r("NoUrut") = txtNoUrut.Value
        r("Kg") = txtKg.Value
        r("Cns") = txtCns.Value
        mdtDetail.Rows.Add(r)
        txtNoUrut.Value = txtNoUrut.Value + 1
    End Sub

    Private Sub dg_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dg.EditingControlShowing
        ctrl = e.Control
    End Sub

    Private Sub dg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dbJenis.KeyDown, dbKodeBarang.KeyDown, dbWarna.KeyDown, dbNoWarna.KeyDown, dbTube.KeyDown, dbGrade.KeyDown
        If e.KeyCode = Keys.F1 Then
            dg.EndEdit()
            Dim dlg As DialogResult
            Dim f As FormPilih
            f = New FormPilihStock(setParam("NoWarna", dbNoWarna.Text, "KodeBarang", dbKodeBarang.Text, "Tube", dbTube.Text, "Grade", dbGrade.Text, "Warna", dbWarna.Text, "Jenis", dbJenis.Text) _
                                   , "Jenis, KodeBarang,Warna,NoWarna, Tube, Grade, SatBesar, SatKecil, JumlahBox, JumlahKG")
            dlg = f.ShowDialog(Me)
            If dlg = Windows.Forms.DialogResult.OK Then
                dbJenis.Text = f.row("Jenis")
                dbKodeBarang.Text = f.row("KodeBarang")
                dbWarna.Text = f.row("Warna")
                dbNoWarna.Text = f.row("NoWarna")
                dbTube.Text = f.row("Tube")
                dbGrade.Text = f.row("Grade")
                dbSatBesar.Text = f.row("SatBesar")
                dbSatKecil.Text = f.row("SatKecil")
            End If
            f.Dispose()
        End If
    End Sub
    Private Function getParam(ByVal row As Integer) As List(Of Param)
        Return setParam("Jenis", Trim(dg.Rows(row).Cells("Jenis").Value.ToString), _
                                                            "KodeBarang", Trim(dg.Rows(row).Cells("KodeBarang").Value.ToString), _
                                                            "Warna", Trim(dg.Rows(row).Cells("Warna").Value.ToString), _
                                                            "NoWarna", Trim(dg.Rows(row).Cells("NoWarna").Value.ToString), _
                                                            "Tube", Trim(dg.Rows(row).Cells("Tube").Value.ToString), _
                                                            "Grade", Trim(dg.Rows(row).Cells("Grade").Value.ToString) _
                                                            )
    End Function


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim tTotalKG As Double
        Dim tTotal As Integer
        Dim tTotalCones As Integer
        tTotalKG = 0
        tTotal = 0
        For Each row As DataGridViewRow In dgDetail.Rows
            If row.IsNewRow Then Exit For
            tTotalKG = tTotalKG + row.Cells("Kg").Value
            tTotalCones = tTotalCones + row.Cells("Cns").Value
            tTotal = tTotal + 1
        Next
        bs.EndEdit()
        mdt.Rows(bs.Position)("Tanggal") = convertToInteger(dtTgl.Value)
        If mdt.Rows(bs.Position)("SatBesar") = mdt.Rows(bs.Position)("SatKecil") Then
            mdt.Rows(bs.Position)("n2") = mdt.Rows(bs.Position)("n1")
            mdt.Rows(bs.Position)("nCones") = mdt.Rows(bs.Position)("n1")
            mdtDetail.Rows.Clear()
        Else
            mdt.Rows(bs.Position)("n1") = tTotal
            mdt.Rows(bs.Position)("n2") = tTotalKG
            mdt.Rows(bs.Position)("nCones") = tTotalCones
        End If
        mdt.Rows(bs.Position)("Status") = 0
        mdt.Rows(bs.Position)("IDStock") = New SQLmStock().qStock(setParam("Jenis", "=" & dbJenis.Text _
                                                    , "KodeBarang", "=" & dbKodeBarang.Text _
                                                    , "Warna", "=" & dbWarna.Text _
                                                    , "NoWarna", "=" & dbNoWarna.Text _
                                                    , "Tube", "=" & dbTube.Text _
                                                    , "Grade", "=" & dbGrade.Text _
                                                    , "SatBesar", "=" & dbSatBesar.Text _
                                                    , "SatKecil", "=" & dbSatKecil.Text _
                                                    ), "", "IDStock").Rows(0)(0)

        If a.Save(mdt.Rows(bs.Position), mdtDetail) Then
            mdt.AcceptChanges()
            MsgBox("Success")
        Else
            MsgBox("Failed")
        End If
    End Sub

    Private Sub txtKg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtKg.KeyDown
        If e.KeyCode = Keys.Decimal Then
            If My.Application.Culture.NumberFormat.CurrencyDecimalSeparator = "," Then
                e.Handled = True
                SendKeys.SendWait(Chr(8) & ",")
            End If
        End If

        If e.KeyCode = Keys.Enter Then
            AddRow()
            txtKg.SelectAll()
            e.Handled = True

        End If

    End Sub



    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        NewClick()
    End Sub
    Private Sub NewClick()
        ModControl.ClearControl(Panel1)
        ModControl.ClearControl(Panel2)
        txtCns.Text = "24"
        Dim r As DataRow
        r = mdt.NewRow
        r("NoBukti") = a.qNextNoBukti()
        r("n1") = 0
        r("n2") = 0
        mdt.Rows.Add(r)
        bs.MoveLast()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        ModControl.ClearControl(Panel1)
        dbNoBukti.Text = a.qNextNoBukti()
    End Sub



    Private Sub bs_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles bs.CurrentChanged
        Try
            If mdt Is Nothing Then Exit Sub
            If Not mdtDetail Is Nothing Then mdtDetail.Clear()
            If bs.Position >= mdt.Rows.Count Then Exit Sub
            GetDetail(mdt.Rows(bs.Position)("NoBukti"))
            
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


End Class


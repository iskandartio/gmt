Public Class frmMStock
    Inherits FormMain
    Dim a As New SQLmStock
    Dim dt As DataTable
    Dim mType As String
    Private Sub cmdStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStock.Click
        Dim z As New List(Of Param)
        z.Add(New Param("KodeBarang", fKode.Text, IIf(fKode.Text.StartsWith("="), True, False)))
        z.Add(New Param("Jenis", fJenis.Text, IIf(fJenis.Text.StartsWith("="), True, False)))
        z.Add(New Param("Warna", fWarna.Text, IIf(fWarna.Text.StartsWith("="), True, False)))
        z.Add(New Param("NoWarna", fNoWarna.Text, IIf(fNoWarna.Text.StartsWith("="), True, False)))
        z.Add(New Param("Tube", fTube.Text, IIf(fTube.Text.StartsWith("="), True, False)))
        z.Add(New Param("Grade", fGrade.Text, IIf(fGrade.Text.StartsWith("="), True, False)))
        dt = a.qStock(z)
        dg.Columns.Clear()
        dg.DataSource = dt
        ModControl.setGridAutoInvisible(dg)
        ModControl.setGridStyle(dg, 0)
        mType = "Stock"
    End Sub

    Private Sub cmdJenis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJenis.Click
        dt = a.qJenis()
        dg.Columns.Clear()
        dg.DataSource = dt
        dg.Columns("Jenis").ReadOnly = True
        mType = "Jenis"
    End Sub

    Private Sub btnKodeBarang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKodeBarang.Click
        dt = a.qKodeBarang()
        dg.Columns.Clear()
        dg.DataSource = dt
        dg.Columns("KodeBarang").ReadOnly = True
        mType = "KodeBarang"
    End Sub
    Private Sub btnWarna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWarna.Click
        dt = a.qWarna()
        dg.Columns.Clear()
        dg.DataSource = dt
        dg.Columns("Warna").ReadOnly = True
        mType = "Warna"
    End Sub
    Private Sub btnNoWarna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNoWarna.Click
        dt = a.qNoWarna()
        dg.Columns.Clear()
        dg.DataSource = dt
        dg.Columns("NoWarna").ReadOnly = True
        mType = "NoWarna"
    End Sub
    Private Sub btnTube_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTube.Click
        dt = a.qTube()
        dg.Columns.Clear()
        dg.DataSource = dt
        dg.Columns("Tube").ReadOnly = True
        mType = "Tube"
    End Sub
    Private Sub btnGrade_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGrade.Click
        dt = a.qGrade()
        dg.Columns.Clear()
        dg.DataSource = dt
        dg.Columns("Grade").ReadOnly = True
        mType = "Grade"
    End Sub
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim v As Boolean
        If mType = "Stock" Then
            v = a.UpdateStock(dt)
        ElseIf mType = "Jenis" Then
            v = a.UpdateJenis(dt)
        ElseIf mType = "KodeBarang" Then
            v = a.UpdateKodeBarang(dt)
        ElseIf mType = "Warna" Then
            v = a.UpdateWarna(dt)
        ElseIf mType = "NoWarna" Then
            v = a.UpdateNoWarna(dt)
        ElseIf mType = "Tube" Then
            v = a.UpdateTube(dt)
        ElseIf mType = "Grade" Then
            v = a.UpdateGrade(dt)
        End If
        If v Then MsgBox("Sukses") Else MsgBox("Gagal")
    End Sub
    Private Sub txtFind_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyDown

        If e.KeyCode = Keys.Enter Then
            dg.Find(txtFind.Text)
            txtFind.Focus()
        End If
    End Sub
    Private Sub btnUbahNama_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUbahNama.Click
        If Not dg.CurrentCell.ReadOnly Then Exit Sub
        Dim s As String = InputBox("Nama Baru:", "", dg.CurrentCell.Value)
        If s <> "" Then
            Dim v As Boolean
            If mType = "Jenis" Then
                v = a.UpdateJenisMain(dg.CurrentRow.Cells("Jenis").Value, s)
            ElseIf mType = "KodeBarang" Then
                v = a.UpdateKodeBarangMain(dg.CurrentRow.Cells("KodeBarang").Value, s)
            ElseIf mType = "Warna" Then
                v = a.UpdateWarnaMain(dg.CurrentRow.Cells("Warna").Value, s)
            ElseIf mType = "NoWarna" Then
                v = a.UpdateNoWarnaMain(dg.CurrentRow.Cells("NoWarna").Value, s)
            ElseIf mType = "Tube" Then
                v = a.UpdateTubeMain(dg.CurrentRow.Cells("Tube").Value, s)
            ElseIf mType = "Grade" Then
                v = a.UpdateGradeMain(dg.CurrentRow.Cells("Grade").Value, s)
            End If
            If v Then
                MsgBox("Sukses")
                If mType = "Jenis" Then
                    cmdJenis_Click(btnJenis, e)
                ElseIf mType = "KodeBarang" Then
                    btnKodeBarang_Click(btnKodeBarang, e)
                ElseIf mType = "Warna" Then
                    btnWarna_Click(btnWarna, e)
                ElseIf mType = "NoWarna" Then
                    btnNoWarna_Click(btnNoWarna, e)
                ElseIf mType = "Tube" Then
                    btnTube_Click(btnTube, e)
                ElseIf mType = "Grade" Then
                    btnGrade_Click(btnGrade, e)
                End If
            Else
                MsgBox("Gagal")
            End If
        End If
    End Sub
End Class
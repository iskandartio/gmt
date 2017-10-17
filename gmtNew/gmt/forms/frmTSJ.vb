Public Class frmTSJ
    Dim mFirstLoad As Boolean
    Dim a As New SQLSJ
    Dim mdt As DataTable
    Dim mdt2 As DataTable
    Dim mReadTime As DateTime
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
    Private Sub frmTSJ_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If Not mFirstLoad Then
            fQuickFind.Text = a.qLastNoSJ()
            dg.AllowUserToAddRows = False
            dg.AllowUserToDeleteRows = False
            DoQuery()
            ModControl.LockMe(txtStatus, True)
            ModControl.LockMe(dbWaktuUpdate, True)
            ModControl.LockMe(dbPengupdate, True)
            mFirstLoad = True
        End If
    End Sub
    Private Sub BindMe()
        ClearBinding(Me)
        SetBinding(Me, mdt)
        If mdt.Rows.Count > 0 Then
            mdt.Rows(0)("TglSJ") = Now.Date
            mdt2 = a.qDataDetail(dbNoSPP.Text)
            dg.DataSource = mdt2
            ModControl.setGridAutoInvisible(dg)
            ModControl.setGridStyle(dg, 0)
        End If
        
    End Sub
    Private Sub DoQuery()
        Dim z As New List(Of Param)
        Dim t() As String
        t = Split(fQuickFind.Text, "/")
        z.Add(New Param("toint(left(NoSPP,5))", "=" & t(0), True))
        If t.Length > 1 Then
            z.Add(New Param("year(TglSPP)", "=" & 2000 + t(1), True))
        Else
            z.Add(New Param("year(TglSPP)", "=" & Year(Now.Date), True))
        End If
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
            btnCancel.Enabled = (mdt.Rows(0)("Status") = 2)
            btnSave.Enabled = Not (mdt.Rows(0)("Status") > 1)
            ModControl.LockMe(dg, (mdt.Rows(0)("Status") > 1))
        End If
        BindMe()
    End Sub

    

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        BindToDataTable(Me)
        If a.Save(mdt, mdt2, mReadTime) Then
            MsgBox("Sukses")
            DoQuery()
        Else
            MsgBox("Gagal")
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If a.Cancel(dbNoSPP.Text) Then
            MsgBox("Sukses")
            DoQuery()
        Else
            MsgBox("Gagal")
        End If
    End Sub
End Class
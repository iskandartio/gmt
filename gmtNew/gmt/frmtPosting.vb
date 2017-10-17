Public Class frmtPosting
    Dim a As New SQLtPosting()
    Private Sub btnPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPost.Click
        refreshData()
    End Sub
    Private Sub refreshData()
        Dim a As New SQLtPosting
        a.Post(dtTgl.Value)
    End Sub

    Private Sub frmtPosting_Load(sender As Object, e As EventArgs) Handles Me.Load
        dtTgl.Value = DateAdd(DateInterval.Day, -1, Now.Date)
    End Sub

    Private Sub btnMutasi_Click(sender As Object, e As EventArgs) Handles btnMutasi.Click
        Dim dtPrint As New DataTable
        dtPrint = a.qMutasi(dtTgl.Value)
        If dtPrint.Rows.Count = 0 Then
            MsgBox("No Data")
            Exit Sub
        End If
        Dim b As New rptMutasi
        b.LoadMe(dtPrint, Format(dtTgl.Value, "dd-MMM-yyyy"))

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        a.generateMutasi()
    End Sub
End Class
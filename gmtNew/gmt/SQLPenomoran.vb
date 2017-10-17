Public Class SQLPenomoran
    Dim a As New db
    Public Function qData(ByVal tNamaForm As String) As DataTable
        Dim s As String
        s = "select * from vmPenomoran where NamaForm='" & esc(tNamaForm) & "'"
        Return a.doQuery(s)
    End Function
End Class

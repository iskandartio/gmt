Public Class SQLmMataUang
    Dim a As New db
    Public Function qData() As DataTable
        Return a.doQuery("select Curr from mMataUang")
    End Function
End Class

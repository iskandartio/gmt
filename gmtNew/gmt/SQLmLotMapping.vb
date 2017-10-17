Public Class SQLmLotMapping
    Dim a As New db
    Private Function esc(ByVal s As String) As String
        Return Replace(s, "'", "''")
    End Function

    Public Function qData(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select * from m_lot_mapping where 1=1" & setFilter(z) & " order by Ket, Warna, KodeBarang, Lot"
        Return a.doQuery(s, z)
    End Function
    Public Function UpdateData(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            Dim dv As New DataView(dt, "", "", DataViewRowState.Deleted)
            For Each dvrow As DataRowView In dv
                tArr.Clear()
                tArr.Add(New Param("rowid", dvrow("rowid")))
                s = "delete from m_lot_mapping where rowid=@rowid"
                a.execMe(s, tArr)
            Next
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "KodeBarang", "Lot", "Warna", "Ket", "rowid")
                If row.RowState = DataRowState.Modified Then
                    s = "update m_lot_mapping set Lot=@Lot, Ket=@Ket, KodeBarang=@KodeBarang, Warna=@Warna" & _
                    " where rowid=@rowid"
                    a.execMe(s, tArr)
                ElseIf row.RowState = DataRowState.Added Then
                    s = "insert into m_lot_mapping(KodeBarang, Lot, Warna, Ket)" & _
                    " values(@KodeBarang, @Lot, @Warna, @Ket)"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            mdt.AcceptChanges()

            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try

    End Function

End Class

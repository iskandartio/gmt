Public Class SQLSTT
    Dim a As New db
    Public Function qNo(ByVal tTgl As Date) As Integer
        Dim s As String
        Dim r As Object
        s = "select max(toint(left(NoSTT,5))) from vtSTT where year(TglSTT)=" & Year(tTgl)
        r = a.doQueryScalar(s)
        If IsDBNull(r) Then r = 1 Else r = r + 1
        Return r
    End Function
    Public Function qData(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select a.NoSTT, a.TglSTT, a.CustomerID, a.WaktuUpdate, a.Pengupdate from vtstt a where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qStructure(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select a.NoSTT, a.TglSTT, a.CustomerID, a.WaktuUpdate, a.Pengupdate from tstt a where 1=0"
        Return a.doQuery(s)
    End Function
    Public Function qDataPembayaran(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select a.CaraBayar, a.TglBayar, a.MataUang, a.Kurs, a.Nilai, a.Kurs*a.Nilai NilaiRp from tSTTPembayaran a where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qStructurePembayaran(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select a.CaraBayar, a.TglBayar, a.MataUang, a.Nilai, a.Kurs, a.Kurs*a.Nilai as NilaiRP from tSTTPembayaran a where 1=0"
        Return a.doQuery(s)
    End Function
    Public Function qDataPelunasan(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "SELECT b.NoSPP, b.TglSJ, b.Total, a.Nilai FROM tSTTPelunasan a" & _
" left join tSPP b on a.NoSPP=b.NoSPP where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qBelumLunas(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "SELECT a.NoSPP, a.MataUang, a.Total, ifnull(sum(b.Nilai),0) Pelunasan, a.Total-ifnull(sum(b.Nilai),0) Sisa,0 Nilai FROM vtspp a" & _
" left join tsttpelunasan b on a.NoSPP=b.NoSPP" & _
" where a.status = 2" & setFilter(z) & _
" group by a.NoSPP, a.MataUang having round(ifnull(sum(b.Nilai),0),0)<round(a.Total,0)"
        Return a.doQuery(s)
    End Function
    Public Function Save(ByVal mdt As DataTable, ByVal mdtPembayaran As DataTable, ByVal mdtPelunasan As DataTable) As Boolean
        Dim s As String
        Dim dt As New DataTable
        Dim p As New List(Of Param)
        Dim tNoSTT As String
        Try
            a.BeginTransaction()
            tNoSTT = mdt.Rows(0)("NoSTT")
            For Each row As DataRow In mdt.Rows
                addArrayAll(p, row, "NoSTT", "TglSTT", "CustomerID")
                p.Add(New Param("@Pengupdate", gUser))
                s = "insert into tSTT(NoSTT, TglSTT, CustomerID, Tipe, Pengupdate, WaktuUpdate)" & _
                    " select @NoSTT, @TglSTT, @CustomerID, Tipe, @Pengupdate, now() from mCurrentTipe"
                a.execMe(s, p)
                
            Next
            
            For Each row As DataRow In mdtPembayaran.Rows
                addArrayAll(p, row, "CaraBayar", "Nilai", "TglBayar", "Kurs", "MataUang")
                p.Add(New Param("@Pengupdate", gUser))
                p.Add(New Param("@NoSTT", tNoSTT))
                s = "insert into tSTTPembayaran(NoSTT, TglBayar, Nilai, CaraBayar, Kurs, MataUang, Pengupdate, WaktuUpdate)" & _
                    " select @NoSTT, @TglBayar, @Nilai, @CaraBayar, @Kurs, @MataUang, @Pengupdate, now() from mCurrentTipe"
                a.execMe(s, p)

            Next

            For Each row As DataRow In mdtPelunasan.Rows
                addArrayAll(p, row, "NoSPP", "Nilai")
                p.Add(New Param("@Pengupdate", gUser))
                p.Add(New Param("@NoSTT", tNoSTT))
                s = "insert into tSTTPelunasan(NoSTT, NoSPP, Nilai, Pengupdate, WaktuUpdate)" & _
                " select @NoSTT, @NoSPP, @Nilai, @Pengupdate, now() from mCurrentTipe"
                a.execMe(s, p)
            Next
            mdt.AcceptChanges()
            mdtPembayaran.AcceptChanges()
            mdtPelunasan.AcceptChanges()
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            a.RollbackTransaction()
            Return False
        End Try

    End Function
    Public Function Cancel(ByVal tNoSTT As String) As Boolean
        Dim s As String
        Try
            a.BeginTransaction()
            s = "delete from tSTT where NoSTT='" & esc(tNoSTT) & "'"
            a.execMe(s)
            s = "delete from tSTTPembayaran where NoSTT='" & esc(tNoSTT) & "'"
            a.execMe(s)
            s = "delete from tSTTPelunasan where NoSTT='" & esc(tNoSTT) & "'"
            a.execMe(s)
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            Return False
        End Try
    End Function
End Class

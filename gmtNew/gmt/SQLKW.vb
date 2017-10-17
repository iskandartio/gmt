Public Class SQLKW
    Dim a As New db
    Public Function qStructure() As DataTable
        Dim s As String
        s = "select tobool(0) Pilih, a.NoSJ, a.TglSJ, a.NamaPenerima, a.Total, a.WaktuUpdate, a.Pengupdate, a.CustomerID, a.MataUang, a.NoKW, a.TglKW from vtspp a where 1=0"
        Return a.doQuery(s)
    End Function
    Public Function qCustomer(ByVal tMataUang As String) As DataTable
        Dim s As String
        s = "select distinct b.* from vtspp a left join mCustomer b on a.CustomerID=b.CustomerID where a.status=2"
        Return a.doQuery(s)
    End Function
    Public Function qData(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select tobool(1) Pilih, a.NoSJ, a.TglSJ, a.NamaPenerima, a.Total, a.WaktuUpdate, a.Pengupdate, a.CustomerID, a.MataUang, a.NoKW, a.TglKW from vtspp a where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qSPP(ByVal tCustomerID As String, ByVal tMataUang As String) As DataTable
        Dim s As String
        s = "select tobool(0) Pilih, a.NoSJ, a.TglSJ, a.NamaPenerima, a.Total from vtspp a where a.CustomerID='" & tCustomerID & "' and a.MataUang='" & tMataUang & "' and a.status=2"
        Return a.doQuery(s)
    End Function
    Public Function qDetailSPP(ByVal tNoSJ As String) As DataTable
        Dim s As String
        s = "select a.IDTrans, a.IDStock, b.Jenis, b.KodeBarang, b.Warna, b.NoWarna, b.Tube, b.Grade, a.JumlahBox, a.JumlahKG, a.Harga, a.JumlahKG*a.Harga TotalNilai, a.KeteranganSPPDetail from tSPPDetail a " & _
        " left join mStock b on a.IDStock=b.IDStock" & _
        " left join tSPP c on c.NoSPP=a.NoSPP" & _
        " where c.NoSJ='" & esc(tNoSJ) & "'"
        Return a.doQuery(s)
    End Function
    Public Function qNo(ByVal tTgl As Date) As Integer
        Dim s As String
        Dim r As Object
        s = "select max(toint(left(NoKW,5))) from vtSPP where year(TglKW)=" & Year(tTgl)
        r = a.doQueryScalar(s)
        If IsDBNull(r) Then r = 1 Else r = r + 1
        Return r
    End Function
    Public Function Save(ByVal mdt As DataTable, ByVal z As List(Of Param)) As Boolean
        Dim s As String
        Try
            a.BeginTransaction()
            z.Add(New Param("@Pengupdate", gUser))
            For Each row As DataRow In mdt.Rows
                If row("Pilih") Then

                    s = "update tSPP set Pengupdate=@Pengupdate, WaktuUpdate=now(), status=status+1, NoKW=@NoKW, TglKW=@TglKW where NoSJ='" & row("NoSJ") & "'"
                    a.execMe(s, z)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            Return False
        End Try

    End Function
    Public Function Cancel(ByVal tNoKW As String) As Boolean
        Dim s As String
        Try
            s = "update tSPP set status=status-1, NoKW=NULL where NoKW='" & esc(tNoKW) & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class

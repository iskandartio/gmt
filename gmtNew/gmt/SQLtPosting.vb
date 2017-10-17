Public Class SQLtPosting
    Dim a As New dbOLEDB("Access")
    Public Sub generateMutasi()
        Dim s As String
        Dim kgs As String()
        Dim dt As DataTable = a.doQuery("select IdStock, Kg from mutasi26062016")
        Dim result As New Dictionary(Of Integer, String)
        Dim param As New List(Of Param)
        For Each row As DataRow In dt.Rows
            If (result.ContainsKey(row("IdStock"))) Then
                result(row("IdStock")) = result(row("IdStock")) & "_" & row("Kg")
            Else
                result.Add(row("IdStock"), row("Kg"))
            End If
        Next
        s = "delete from mutasi"
        a.execMe(s)
        Dim totalKg As Double
        For Each kvp As KeyValuePair(Of Integer, String) In result
            kgs = kvp.Value.Split("_")
            totalKg = 0
            For Each kg As String In kgs
                totalKg = totalKg + kg
            Next

            param.Clear()
            param.Add(New Param("@IdStock", kvp.Key))
            param.Add(New Param("@akhirBox", kgs.Length))
            param.Add(New Param("@akhirKg", totalKg))
            param.Add(New Param("@kgs", kvp.Value))
            s = "insert into mutasi(tgl, IdStock, akhirBox, akhirKg, kgs) values(160626, @IdStock, @akhirBox, @akhirKg, @kgs)"
            a.execMe(s, param)
        Next
    End Sub
    Public Function qMutasi(ByVal tDate As Date) As DataTable
        Dim s As String
        Dim b As New db
        s = "SELECT b.Warna, b.NoWarna, b.KodeBarang, b.Grade, a.inBox, a.inKg, a.outBox, a.outKg, a.akhirBox, a.akhirKg FROM mutasi a LEFT JOIN mStock b ON a.IdStock = b.IdStock" & _
            " where (a.inKg<>0 or a.outKg<>0 or a.akhirKg<>0)" & _
            " order by b.IdWarna, b.IdKodeBarang, b.IdGrade"
        Return b.doQuery(s)

    End Function
    Public Function qData(ByVal tDate As Date) As Dictionary(Of String, DataTable)
        Dim sqlDef As String
        Dim s As String
        Dim result As New Dictionary(Of String, DataTable)
        Dim b As New db
        Dim lastDate As Object
        lastDate = b.doQueryScalar("select max(tgl) from mutasi")
        If IsDBNull(lastDate) OrElse DateAdd(DateInterval.Day, 1, lastDate) > tDate Then
            s = "truncate table mutasi"
            b.execMe(s)
            s = "delete from mutasi_hist where tgl>=" & convertToInteger(tDate)
            b.execMe(s)
            lastDate = b.doQueryScalar("select max(tgl) from mutasi_hist")
            If (IsDBNull(lastDate)) Then
                lastDate = New Date(2015, 12, 31)
            Else
                s = "insert into mutasi select * from mutasi_hist where tgl=" & convertToInteger(lastDate)
                b.execMe(s)

            End If

        End If
        lastDate = DateAdd(DateInterval.Day, 1, lastDate)
        sqlDef = "select Tanggal as tgl, IDStock, n1 as box, n2 as kg from t_InputStock~ where Tanggal between " & convertToInteger(lastDate) & " and " & convertToInteger(tDate)
        s = sqlDef.Replace("~", "DTY")
        Dim dt As New DataTable
        dt = a.doQuery(s)
        result.Add("inputStock", dt)

        sqlDef = "select TanggalDetail as tgl, IDStock, JumlahBox as box, JumlahKG as kg from t_SPPDetail~ where TanggalDetail between " & convertToInteger(lastDate) & " and " & convertToInteger(tDate)
        s = sqlDef.Replace("~", "DTY")
        dt = New DataTable
        dt = a.doQuery(s)
        result.Add("outputStock", dt)

        sqlDef = "select IdWarna, IdKodeBarang, IdGrade, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, IDStock from m_stock~"
        s = sqlDef.Replace("~", "DTY")
        dt = New DataTable
        dt = a.doQuery(s)
        result.Add("mStock", dt)

        Return result
    End Function

    Public Sub Save(tbl As DataTable, table_name As String)
        Dim s As String
        Dim b As New db
        If table_name = "mStock" Then
            b.execMe("truncate table mStock")
            For Each row As DataRow In tbl.Rows
                Dim arr As New List(Of Param)
                addArrayAll(arr, row, "IdWarna", "IdKodeBarang", "IdGrade", "Jenis", "KodeBarang", "Warna", "NoWarna", "Tube", "Grade", "SatBesar", "SatKecil", "IDStock")
                s = "insert into " & table_name & " (IdWarna, IdKodeBarang, IdGrade, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, IDStock)" & _
                    " values(@IdWarna, @IdKodeBarang, @IdGrade, @Jenis, @KodeBarang, @Warna, @NoWarna, @Tube, @Grade, @SatBesar, @SatKecil, @IDStock)"
                b.execMe(s, arr)
            Next

        Else
            For Each row As DataRow In tbl.Rows
                Dim arr As New List(Of Param)
                addArrayAll(arr, row, "tgl", "IDStock", "box", "kg")
                s = "insert into " & table_name & " (tgl, IDStock, box, kg)" & _
                        " values(@tgl, @IDStock, @box, @kg)"
                b.execMe(s, arr)
            Next

        End If
        
    End Sub
    Public Sub Post(tDate As Date)
        Dim data As New Dictionary(Of String, DataTable)
        Dim b As New db
        Dim s As String
        data = qData(tDate)

        Truncate()
        Save(data("inputStock"), "inputStock")
        Save(data("outputStock"), "outputStock")
        Save(data("mStock"), "mStock")
        b.BeginTransaction()
        b.execMe("DROP TEMPORARY TABLE IF EXISTS a")
        s = "create temporary table a" & _
" select @tgl, IDStock, sum(case when tgl=@tgl then inBox else 0 end) inBox, sum(case when tgl=@tgl then inKg else 0  end) inKg" & _
" , sum(case when tgl=@tgl then outBox else 0 end) outBox, sum(case when tgl=@tgl then outKg else 0 end) outKg" & _
" , sum(akhirBox)+sum(inBox)-sum(outBox) akhirBox" & _
" , sum(akhirKg)+sum(inKg)-sum(outKg) akhirKg" & _
"   from (" & _
" select tgl, IDStock, akhirBox, akhirKg, 0 inBox, 0.0 inKg, 0 outBox, 0.0 outKg from mutasi" & _
" union all select tgl, IDStock, 0,0,box, kg,  0,0 from inputStock" & _
" union all select tgl, IDStock, 0,0,0,0,box, kg from outputStock) a" & _
" group by IDStock"
        s = s.Replace("@tgl", convertToInteger(tDate))
        b.execMe(s)
        s = "insert into mutasi_hist select * from mutasi"
        b.execMe(s)
        s = "truncate table mutasi"
        b.execMe(s)
        s = "insert into mutasi select * from a"
        b.execMe(s)
        b.CommitTransaction()

    End Sub
    Public Sub Truncate()
        Dim b As New db
        b.execMe("truncate table inputStock")
        b.execMe("truncate table outputStock")
    End Sub
End Class

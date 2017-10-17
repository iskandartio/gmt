Public Class SQLtInputStock
    Dim a As New dbOLEDB("Access")
    Public mCurrentIDStock As Integer
    Public Function qPrint() As DataTable
        Dim s As String
        s = "SELECT m_Stock~.IDStock, t_InputStock~.Lot, m_Stock~.Jenis, m_Stock~.Warna, m_Stock~.KodeBarang, m_Stock~.NoWarna, m_Stock~.Tube, m_Stock~.Grade, m_Stock~.SatBesar, m_Stock~.SatKecil, t_InputStock~.Tanggal, t_InputStockDetail~.Cns, t_InputStockDetail~.Kg, t_InputStockDetail~.NoUrut, t_InputStockDetail~.NoBukti" & _
" FROM t_InputStockDetail~ INNER JOIN (t_InputStock~ INNER JOIN m_Stock~ ON t_InputStock~.IdStock = m_Stock~.IdStock) ON t_InputStockDetail~.NoBukti = t_InputStock~.NoBukti where t_InputStock~.Status=0 order by t_InputStock~.NoBukti, m_Stock~.IDStock, t_InputStockDetail~.NoUrut"
        Return a.doQuery(s)
    End Function
    Public Function Printed() As Boolean
        Dim s As String
        s = "update t_InputStock~ set Status=1 where Status=0"
        Return a.execMe(s)
    End Function
    Public Function OK() As Boolean
        Dim s As String
        s = "update t_InputStock~ set Status=2 where Status=1"
        Return a.execMe(s)
    End Function
    Public Function OKGudang(ByVal tID As Integer) As Boolean
        Dim s As String
        Try
            a.BeginTransaction()
            Dim TanggalGudang As Integer
            TanggalGudang = ModFunction.convertToInteger(Now)
            s = "update t_InputStock~ set Status=3, TanggalGudang=" & TanggalGudang & " where IDTrans = " & tID
            a.execMe(s)
            s = "UPDATE t_InputStock~ as a INNER JOIN m_Stock~ as b ON a.IdStock = b.IdStock SET b.JumlahBox = b.JumlahBox+a.n1, b.JumlahKG=b.JumlahKG+a.n2 where a.IDTrans=" & tID
            a.execMe(s)
            a.CommitTransaction()
        Catch ex As Exception
            a.RollbackTransaction()
        End Try
    End Function
    Public Function OKGudangAll(ByVal tTgl As Date) As Boolean
        Dim s As String
        Try
            a.BeginTransaction()
            Dim TanggalGudang As Integer
            TanggalGudang = ModFunction.convertToInteger(Now)
            s = "update t_InputStock~ set Status=3, TanggalGudang=" & TanggalGudang & " where Tanggal = " & ModFunction.convertToInteger(tTgl)
            a.execMe(s)
            s = "UPDATE t_InputStock~ as a INNER JOIN m_Stock~ as b ON a.IdStock = b.IdStock SET b.JumlahBox = b.JumlahBox+a.n1, b.JumlahKG=b.JumlahKG+a.n2 where a.Tanggal = " & ModFunction.convertToInteger(tTgl)
            a.execMe(s)
            a.CommitTransaction()
        Catch ex As Exception
            a.RollbackTransaction()
        End Try
    End Function
    Public Function CancelGudang(ByVal tID As Integer) As Boolean
        Dim s As String
        s = "update t_InputStock~ set Status=0 where IDTrans=" & tID
        Return a.execMe(s)
    End Function
    Public Function qNextNoUrut2(ByVal tNoBukti As String) As Integer
        Dim s As String
        s = "select iif(isnull(max(a.NoUrut)), 1, max(a.NoUrut)+1) from (t_InputStockDetail~ as a" & _
" left join t_InputStock~ as b on a.NoBukti=b.NoBukti)" & _
" left join m_Stock~ as c on c.IDStock=b.IDStock" & _
" where b.NoBukti=" & tNoBukti
        Return a.doQueryScalar(s)
    End Function
    Public Function qNextNoUrut(ByVal tKodeBarang As String, ByVal tWarna As String, ByVal tNoWarna As String) As Integer
        Dim s As String
        s = "select iif(isnull(max(a.NoUrut)), 1, max(a.NoUrut)+1) from (t_InputStockDetail~ as a" & _
" left join t_InputStock~ as b on a.NoBukti=b.NoBukti)" & _
" left join m_Stock~ as c on c.IDStock=b.IDStock" & _
" where c.KodeBarang='" & tKodeBarang & _
"' and c.Warna='" & tWarna & _
"' and c.NoWarna='" & tNoWarna & "'"
        Return a.doQueryScalar(s)
    End Function
    Public Function qNextNoBukti() As Integer
        Dim s As String
        s = "select iif(isnull(max(a.NoBukti)), 1, max(a.NoBukti)+1) from t_InputStock~ as a"
        Return a.doQueryScalar(s)
    End Function
    Public Function qData(ByVal tDate As Date) As DataTable
        Dim s As String
        s = "select Status, Shift, Tanggal, NoBukti, Lot, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, m_Stock~.IdStock, n1,nCones, n2, Keterangan, PrintedCode, IDTrans from t_InputStock~ left join m_Stock~ on t_InputStock~.IdStock=m_Stock~.IdStock where Tanggal=" & convertToInteger(tDate) & " and Status<2 order by NoBukti"
        Return a.doQuery(s)
    End Function
    Public Function qDataPrinted(ByVal tDate As Date) As DataTable
        Dim s As String
        s = "select 'OK' as [OK], 'CANCEL' as [CANCEL], Shift, Tanggal, NoBukti, Lot, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, SatBesar, SatKecil, m_Stock~.IdStock, n1,n2, Keterangan, PrintedCode, IDTrans from t_InputStock~ left join m_Stock~ on t_InputStock~.IdStock=m_Stock~.IdStock where Tanggal=" & convertToInteger(tDate) & " and Status=2 order by NoBukti"
        Return a.doQuery(s)
    End Function
    Public Function qRekap(ByVal tDate As Date, ByVal tDateAkhir As Date) As DataTable
        Dim s As String
        s = "SELECT b.KodeBarang,  mid(Tanggal,5,2)&'-'&mid(Tanggal,3,2)&'-'&2000+left(Tanggal,2) as Tanggal, a.Lot, b.Warna, b.NoWarna, b.Grade, b.Tube, a.nCones, a.n1, a.n2  FROM t_InputStock~ as a INNER JOIN m_Stock~ as b ON a.IdStock = b.IdStock where Tanggal between " & convertToInteger(tDate) & " and " & convertToInteger(tDateAkhir) & " order by b.IdKodeBarang, a.Lot, b.WarnaDasar, b.NoWarna, b.IdTube, b.IdGrade"
        Return a.doQuery(s)
    End Function
    Public Function qSummary(ByVal tDate As Date, ByVal tDateAkhir As Date) As DataTable
        Dim s As String
        s = "SELECT b.IdWarna, b.Warna, b.IdKodeBarang, b.KodeBarang, b.IdGrade, b.Grade, sum(a.n1) as n1, sum(a.n2) as n2 FROM t_InputStock~ as a INNER JOIN m_Stock~ as b ON a.IdStock = b.IdStock where Tanggal between " & convertToInteger(tDate) & " and " & convertToInteger(tDateAkhir) & " group by b.IdWarna, b.Warna, b.IdKodeBarang, b.KodeBarang, b.IdGrade, b.Grade order by b.IdWarna, b.IdKodeBarang, b.IdGrade"
        Return a.doQuery(s)
    End Function
    Public Function qDetail(ByVal tNoBukti As Integer) As DataTable
        Dim s As String
        s = "select * from t_InputStockDetail~ where NoBukti=" & tNoBukti & " order by NoUrut"
        Return a.doQuery(s)
    End Function
    Public Function Save(ByVal mRow As DataRow, ByVal mdtDetail As DataTable) As Boolean
        Dim s As String
        Try
            a.BeginTransaction()
            Dim arr As New List(Of Param)
            addArrayAll(arr, mRow, "NoBukti", "Shift", "Tanggal", "Lot", "IDStock", "Status", "n1", "n2", "nCones", "Keterangan", "IDTrans")
            arr.Add(New Param("@PrintedCode", CLng(Rnd() * 100000000)))
            If mRow.RowState = DataRowState.Added Then
                s = "insert into t_InputStock~(NoBukti, Shift, Tanggal, Lot, IDStock, Status, n1,n2,nCones,PrintedCode, Keterangan)" & _
                " values(@NoBukti, @Shift, @Tanggal, @Lot, @IDStock, @Status, @n1,@n2,@nCones, @PrintedCode, @Keterangan)"
                a.execMe(s, arr)
            ElseIf mRow.RowState = DataRowState.Modified Then
                If mRow("n1") = 0 Then
                    s = "delete from t_InputStock~ where IDTrans=" & mRow("IDTrans")
                    a.execMe(s, arr)
                    s = "delete from t_InputStockDetail~ where IDTrans=" & mRow("NoBukti")
                    a.execMe(s, arr)
                Else
                    s = "update t_InputStock~ set NoBukti=@NoBukti, Shift=@Shift, Tanggal=@Tanggal, Lot=@Lot, IDStock=@IDStock, Status=@Status, n1=@n1, n2=@n2, nCones=@nCones, PrintedCode=@PrintedCode, Keterangan=@Keterangan where IDTrans=@IDTrans"
                    a.execMe(s, arr)
                End If
            End If
            Dim dt As DataTable
            dt = mdtDetail.GetChanges
            If Not dt Is Nothing Then
                Dim dr As New DataView(dt, "", "", DataViewRowState.Deleted)
                For Each row As DataRowView In dr
                    s = "delete from t_InputStockDetail~ where IDTrans=" & row("IDTrans")
                    a.execMe(s)
                Next
                For Each row As DataRow In dt.Rows
                    If row.RowState = DataRowState.Deleted Then Continue For
                    addArrayAll(arr, row, "NoBukti", "Cns", "Kg", "NoUrut", "IDTrans")
                    If row.RowState = DataRowState.Added Then
                        s = "insert into t_InputStockDetail~(NoBukti, Cns, Kg, NoUrut) " & _
                        " values(@NoBukti, @Cns, @Kg, @NoUrut)"
                        a.execMe(s, arr)
                    ElseIf row.RowState = DataRowState.Modified Then
                        s = "update t_InputStockDetail~ set NoBukti=@NoBukti, Cns=@Cns, Kg=@Kg, NoUrut=@NoUrut where IDTrans=@IDTrans"
                        a.execMe(s, arr)
                    End If
                Next
            End If
            mdtDetail.AcceptChanges()
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            Return False
        End Try


    End Function
End Class

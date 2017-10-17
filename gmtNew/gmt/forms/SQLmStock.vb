Public Class SQLmStock
    Dim a As New dbOLEDB("Access")

    Public Function esc(ByVal tStr As String) As String
        Return Replace(tStr, "'", "''")
    End Function
    Public Function qDataSingle(ByVal tTipe As String, Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select distinct " & tTipe & " from vmstock where 1=1" & setFilter(z, True) & " order by ID" & tTipe
        Return a.doQuery(s)
    End Function
    Public Function qDataJenis(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select distinct Jenis from vmstock where 1=1" & setFilter(z) & " order by IDJenis"
        Return a.doQuery(s)
    End Function
    Public Function qDataKodeBarang(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select distinct KodeBarang from vmstock where 1=1" & setFilter(z) & " order by IDKodeBarang"
        Return a.doQuery(s)
    End Function
    Public Function qDataWarna(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select distinct Warna from vmstock where 1=1" & setFilter(z) & " order by IDWarna"
        Return a.doQuery(s)
    End Function
    Public Function qDataNoWarna(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select distinct NoWarna from vmstock where 1=1" & setFilter(z) & " order by IDNoWarna"
        Return a.doQuery(s)
    End Function
    Public Function qDataTube(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select distinct Tube from vmstock where 1=1" & setFilter(z) & " order by IDTube"
        Return a.doQuery(s)
    End Function
    Public Function qDataGrade(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select distinct Grade from vmstock where 1=1" & setFilter(z) & " order by IDGrade"
        Return a.doQuery(s)
    End Function
    Public Function qStock(ByVal z As List(Of Param), Optional ByVal tDefaultFilter As String = "", Optional ByVal tFields As String = "") As DataTable
        Dim s As String
        s = "select " & IIf(tFields = "", "*", tFields) & " from m_stock~ where 1=1" & setFilter(z) & tDefaultFilter
        Return a.doQuery(s)
    End Function
    Public Function qJenis() As DataTable
        Dim s As String
        s = "select distinct IDJenis, Jenis from vmstock where Jenis<>'' order by IDJenis"
        Return a.doQuery(s)
    End Function
    Public Function qKodeBarang() As DataTable
        Dim s As String
        s = "select distinct IDKodeBarang, D, F, NamaID, KonversiKG, KodeBarang from vmstock where KodeBarang<>'' order by IDKodeBarang, D, F, NamaID"
        Return a.doQuery(s)
    End Function
    Public Function qWarna() As DataTable
        Dim s As String
        s = "select distinct IDWarna, Warna from vmstock where Warna<>'' order by IDWarna"
        Return a.doQuery(s)
    End Function
    Public Function qNoWarna() As DataTable
        Dim s As String
        s = "select distinct IDNoWarna, NoWarna, KetWarna from vmstock where NoWarna<>'' order by IDNoWarna"
        Return a.doQuery(s)
    End Function
    Public Function qTube() As DataTable
        Dim s As String
        s = "select distinct IDTube, Tube from vmstock where Tube<>'' order by IDTube"
        Return a.doQuery(s)
    End Function
    Public Function qGrade() As DataTable
        Dim s As String
        s = "select distinct IDGrade, Grade from vmstock where Grade<>'' order by IDGrade"
        Return a.doQuery(s)
    End Function
    Public Function UpdateStock(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim dt2 As New DataTable
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            Dim dv As New DataView(dt, "", "", DataViewRowState.Deleted)
            For Each dvrow As DataRowView In dv
                tArr.Add(New Param("IDStock", "IDStock"))
                s = "delete from mstock where IDStock=@IDStock"
                If a.execMe(s, tArr) = 0 Then Throw New Exception(dvrow("IDStock") & " : Data terpakai")
            Next
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                s = "select IDJenis from vmstock where Jenis='" & esc(row("Jenis")) & "'"
                row("IDJenis") = a.doQueryScalar(s, tArr)
                s = "select IDWarna from vmstock where Warna='" & esc(row("Warna")) & "'"
                row("IDWarna") = a.doQueryScalar(s, tArr)
                s = "select IDTube from vmstock where Tube='" & esc(row("Tube")) & "'"
                row("IDTube") = a.doQueryScalar(s, tArr)
                s = "select IDGrade from vmstock where Grade='" & esc(row("Grade")) & "'"
                row("IDGrade") = a.doQueryScalar(s, tArr)
                s = "select IDKodeBarang, D, F, NamaID, KonversiKG from vmstock where KodeBarang='" & esc(row("KodeBarang")) & "' limit 1"
                dt2 = a.doQuery(s, tArr)
                If Not dt2 Is Nothing Then
                    row("IDKodeBarang") = dt2.Rows(0)("IDKodeBarang")
                    row("D") = dt2.Rows(0)("D")
                    row("F") = dt2.Rows(0)("F")
                    row("NamaID") = dt2.Rows(0)("NamaID")
                    row("KonversiKG") = dt2.Rows(0)("KonversiKG")
                    row("NamaID") = dt2.Rows(0)("NamaID")
                End If
                s = "select IDNoWarna, KetWarna, IDNoWarna from vmstock where NoWarna='" & esc(row("NoWarna")) & "'"
                dt2 = a.doQuery(s, tArr)
                If Not dt2 Is Nothing Then
                    row("IDNoWarna") = dt2.Rows(0)("IDNoWarna")
                    row("KetWarna") = dt2.Rows(0)("KetWarna")
                    row("IDNoWarna") = dt2.Rows(0)("IDNoWarna")
                    row("KetWarna") = dt2.Rows(0)("KetWarna")
                End If
                addArrayAll(tArr, row, "IDStock", "nDet", "SatBesar", "JumlahBox", "JumlahKG", "IDJenis", "Jenis", "IDWarna", "Warna", "IDKodeBarang", "KodeBarang", "NoWarna", "Tube", "Grade", "SatKecil", "D", "F", "IDNoWarna", "IdTube", "IdGrade", "KetWarna", "KonversiKG", "NamaBarang", "IsActive", "IdKodeBarang", "NamaId")
                If row.RowState = DataRowState.Modified Then

                    s = "update vmstock set IDStock=@IDStock, nDet=@nDet, SatBesar=@SatBesar, JumlahBox=@JumlahBox, JumlahKG=@JumlahKG, Jenis=@Jenis, Warna=@Warna, KodeBarang=@KodeBarang, NoWarna=@NoWarna, Tube=@Tube, Grade=@Grade, SatKecil=@SatKecil, IdJenis=@IdJenis, IdWarna=@IdWarna, D=@D, F=@F, IDNoWarna=@IDNoWarna, IdTube=@IdTube, IdGrade=@IdGrade, KetWarna=@KetWarna, KonversiKG=@KonversiKG, NamaBarang=@NamaBarang, IsActive=@IsActive, IdKodeBarang=@IdKodeBarang, NamaId=@NamaId" & _
                    " where IDStock=@IDStock"
                    a.execMe(s, tArr)
                ElseIf row.RowState = DataRowState.Added Then
                    s = "insert into vmstock(IDStock, nDet, SatBesar, JumlahBox, JumlahKG, Jenis, Warna, KodeBarang, NoWarna, Tube, Grade, SatKecil, IdJenis, IdWarna, D, F, IDNoWarna, IdTube, IdGrade, KetWarna, KonversiKG, NamaBarang, IsActive, IdKodeBarang, NamaId)" & _
                    " values(@IDStock, @nDet, @SatBesar, @JumlahBox, @JumlahKG, @Jenis, @Warna, @KodeBarang, @NoWarna, @Tube, @Grade, @SatKecil, @IdJenis, @IdWarna, @D, @F, @IDNoWarna, @IdTube, @IdGrade, @KetWarna, @KonversiKG, @NamaBarang, @IsActive, @IdKodeBarang, @NamaId)"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function


    Public Function UpdateJenisMain(ByVal tAwal As String, ByVal tAkhir As String) As Boolean
        Try
            Dim s As String
            s = "update vmstock set Jenis='" & esc(tAkhir) & "' where Jenis='" & esc(tAwal) & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function UpdateJenis(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "IdJenis", "Jenis")
                If row.RowState = DataRowState.Modified Then
                    s = "update vmstock set IdJenis=@IdJenis" & _
                    " where Jenis=@Jenis"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function UpdateKodeBarangMain(ByVal tAwal As String, ByVal tAkhir As String) As Boolean
        Try
            Dim s As String
            s = "update vmstock set KodeBarang='" & esc(tAkhir) & "' where KodeBarang='" & esc(tAwal) & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function UpdateKodeBarang(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "IdKodeBarang", "KodeBarang", "D", "F", "KonversiKG", "NamaID")
                If row.RowState = DataRowState.Modified Then
                    s = "update vmstock set IdKodeBarang=@IdKodeBarang, D=@D, F=@F, KonversiKG=@KonversiKG, NamaID=@NamaID" & _
                    " where KodeBarang=@KodeBarang"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function UpdateWarnaMain(ByVal tAwal As String, ByVal tAkhir As String) As Boolean
        Try
            Dim s As String
            s = "update vmstock set Warna='" & esc(tAkhir) & "' where Warna='" & esc(tAwal) & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function UpdateWarna(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "IdWarna", "Warna")
                If row.RowState = DataRowState.Modified Then
                    s = "update vmstock set IdWarna=@IdWarna" & _
                    " where Warna=@Warna"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function UpdateNoWarnaMain(ByVal tAwal As String, ByVal tAkhir As String) As Boolean
        Try
            Dim s As String
            s = "update vmstock set NoWarna='" & esc(tAkhir) & "' where NoWarna='" & esc(tAwal) & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function UpdateNoWarna(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "IDNoWarna", "NoWarna", "KetWarna")
                If row.RowState = DataRowState.Modified Then
                    s = "update vmstock set IDNoWarna=@IDNoWarna, KetWarna=@KetWarna" & _
                    " where NoWarna=@NoWarna"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function UpdateTubeMain(ByVal tAwal As String, ByVal tAkhir As String) As Boolean
        Try
            Dim s As String
            s = "update vmstock set Tube='" & esc(tAkhir) & "' where Tube='" & esc(tAwal) & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function UpdateTube(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "IdTube", "Tube")
                If row.RowState = DataRowState.Modified Then
                    s = "update vmstock set IdTube=@IdTube" & _
                    " where Tube=@Tube"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function UpdateGradeMain(ByVal tAwal As String, ByVal tAkhir As String) As Boolean
        Try
            Dim s As String
            s = "update vmstock set Grade='" & esc(tAkhir) & "' where Grade='" & esc(tAwal) & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function UpdateGrade(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "IdGrade", "Grade")
                If row.RowState = DataRowState.Modified Then
                    s = "update vmstock set IdGrade=@IdGrade" & _
                    " where Grade=@Grade"
                    a.execMe(s, tArr)
                End If
            Next
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

End Class

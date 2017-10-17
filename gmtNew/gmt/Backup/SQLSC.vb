Public Class SQLSC
    Dim a As New db
    Public Function qStructure() As DataTable
        Return a.doQuery("select * from tSC where 1=0", Nothing, False, True)
    End Function
    Public Function setFilter(Optional ByVal z As List(Of Param) = Nothing, Optional ByVal tStartWith As Boolean = False) As String
        Dim tFilter As String = ""
        If z Is Nothing Then Return ""
        For Each p As Param In z
            If p.Exact Then
                If Not IsDBNull(p.ParamValue) AndAlso p.ParamValue <> "" Then tFilter &= " and " & p.ParamName & " ='" & esc(p.ParamValue.ToString.Substring(1)) & "'"
            Else
                If Not IsDBNull(p.ParamValue) AndAlso p.ParamValue <> "" Then tFilter &= " and " & p.ParamName & " like '" & IIf(tStartWith, "", "%") & esc(p.ParamValue) & "%'"
            End If
        Next
        Return tFilter
    End Function
    Public Function qData(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select * from vtSC where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qDataDetail(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select IDTrans, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, Jumlah, Satuan, Harga, Keterangan, StatusDetail  from tSCDetail where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qCustomerSC() As DataTable
        Dim s As String = "select b.* from vtsc a inner join mCustomer b on a.CustomerID=b.CustomerID and Status=0"
        Return a.doQuery(s)
    End Function
    Public Function qDataDetailLengkap(Optional ByVal z As List(Of Param) = Nothing, Optional ByVal tDefaultFilter As String = "") As DataTable
        Dim s As String
        s = "select a.* from tscdetail a left join vtsc b on a.NoSC=b.NoSC where 1=1" & setFilter(z) & tDefaultFilter
        Return a.doQuery(s)
    End Function
    Public Function qNo(ByVal tTgl As Date) As Integer
        Dim s As String
        Dim r As Object
        s = "select max(toint(left(NoSC,5))) from vtsc where year(TglSC)=" & Year(tTgl)
        r = a.doQueryScalar(s)
        If IsDBNull(r) Then r = 1 Else r = r + 1
        Return r
    End Function
    Public Function SetStatus(ByVal tNoSC As String, ByVal st As Integer) As Boolean
        Try
            Dim s As String
            s = "update tSC set status=" & st & " where NoSC='" & tNoSC & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function Save(ByVal mdt As DataTable, ByVal mdt2 As DataTable, ByVal tReadTime As DateTime, Optional ByVal tID As Integer = 0) As Boolean
        Dim tIDHeader As String
        Dim s As String = ""
        Dim z As New List(Of Param)
        Try
            a.BeginTransaction()
            tIDHeader = mdt.Rows(0)("NoSC")
            
            'Header 
            addArrayAll(z, mdt.Rows(0), "NoSC", "TglSC", "CustomerID", "MataUang", "LamaKontrak", "WaktuPembayaran", "NilaiKontrak", "Keterangan", "Status", "DP", "NamaMarketing", "NamaCustomerSC", "IDTrans")
            z.Add(New Param("Pengupdate", gUser))
            If tReadTime.Ticks = 0 Then
                s = "insert into tSC(Tipe, NoSC, TglSC, CustomerID, MataUang, LamaKontrak, WaktuPembayaran, NilaiKontrak, Keterangan, Status, Pengupdate, WaktuUpdate, DP, NamaMarketing, NamaCustomerSC) " & _
                " select Tipe, @NoSC, @TglSC, @CustomerID, @MataUang, @LamaKontrak, @WaktuPembayaran, @NilaiKontrak, @Keterangan, 0, @Pengupdate, now(), @DP, @NamaMarketing, @NamaCustomerSC from mCurrentTipe"
            Else
                s = "update tSC set NoSC=@NoSC, TglSC=@TglSC, CustomerID=@CustomerID, MataUang=@MataUang, LamaKontrak=@LamaKontrak, WaktuPembayaran=@WaktuPembayaran, NilaiKontrak=@NilaiKontrak, Keterangan=@Keterangan, Status=@Status, Pengupdate=@Pengupdate, WaktuUpdate=now(), DP=@DP, NamaMarketing=@NamaMarketing, NamaCustomerSC=@NamaCustomerSC where IDTrans=@IDTrans"
            End If
            a.execMe(s, z)
            

            Dim dt As DataTable = mdt2.GetChanges

            If Not dt Is Nothing Then
                Dim dv As New DataView(dt, "", "", DataViewRowState.Deleted)
                For Each dvrow As DataRowView In dv
                    z.Clear()
                    z.Add(New Param("IDTrans", dvrow("IDTrans")))
                    s = "delete from tSCDetail where IDTrans=@IDTrans"
                    If a.execMe(s, z) = 0 Then
                        Throw New Exception(dvrow("IDTrans") & " : Data terpakai")
                    End If

                Next

                For Each row As DataRow In dt.Rows
                    If row.RowState = DataRowState.Deleted Then Continue For
                    addArrayAll(z, row, "IDTrans", "Jenis", "KodeBarang", "Warna", "NoWarna", "Tube", "Grade", "Satuan", "Jumlah", "Harga", "Keterangan")
                    z.Add(New Param("Pengupdate", gUser))
                    z.Add(New Param("NoSC", tIDHeader))
                    If row.RowState = DataRowState.Added Then
                        s = "insert into tSCDetail(NoSC, Jenis, KodeBarang, Warna, NoWarna, Tube, Grade, Satuan, Jumlah, Harga, Keterangan, Pengupdate, WaktuUpDate)" & _
                        " select @NoSC, @Jenis, @KodeBarang, @Warna, @NoWarna, @Tube, @Grade, @Satuan, @Jumlah, @Harga, @Keterangan, @Pengupdate, now()"
                    Else
                        s = "update tSCDetail set NoSC=@NoSC, Jenis=@Jenis, KodeBarang=@KodeBarang, Warna=@Warna, NoWarna=@NoWarna, Tube=@Tube, Grade=@Grade, Satuan=@Satuan, Jumlah=@Jumlah, Harga=@Harga, Keterangan=@Keterangan, Pengupdate=@Pengupdate, WaktuUpdate=now() where IDTrans=@IDTrans"
                    End If
                    a.execMe(s, z)
                Next
            End If

            mdt.AcceptChanges()
            mdt2.AcceptChanges()
            a.CommitTransaction()
            Return True

        Catch ex As Exception
            a.RollbackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try

    End Function
End Class

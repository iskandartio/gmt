Public Class SQLSPP
    Dim a As New db
    Public Function qLastNoSJ() As Object
        Dim s As String
        s = "select min(toint(left(NoSPP,5))) from vtspp where status=1"
        Return a.doQueryScalar(s)
    End Function
    Public Function qStructure() As DataTable
        Return a.doQuery("select * from tSPP where 1=0", Nothing, False, True)
    End Function
    Public Function qData(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select * from vtSPP where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qDataDetail(ByVal tNoSPP As String) As DataTable
        Dim s As String
        s = "select a.IDTrans, a.IDSC, a.IDStock, c.NoSC, b.Jenis, b.KodeBarang, b.Warna, b.NoWarna, b.Tube, b.Grade, a.JumlahBox, a.Harga, a.KeteranganSPPDetail from tSPPDetail a " & _
        " left join mStock b on a.IDStock=b.IDStock" & _
        " left join tSCDetail c on a.IDSC=c.IDTrans" & _
        " where a.NoSPP='" & esc(tNoSPP) & "'"
        Return a.doQuery(s)
    End Function
    Public Function qNo(ByVal tTgl As Date) As Integer
        Dim s As String
        Dim r As Object
        s = "select max(toint(left(NoSPP,5))) from vtSPP where year(TglSPP)=" & Year(tTgl)
        r = a.doQueryScalar(s)
        If IsDBNull(r) Then r = 1 Else r = r + 1
        Return r
    End Function
    Public Function Save(ByVal mdt As DataTable, ByVal mdt2 As DataTable, ByVal tReadTime As DateTime, Optional ByVal tID As Integer = 0) As Boolean
        Dim tIDHeader As String
        Dim s As String = ""
        Dim z As New List(Of Param)
        Try
            a.BeginTransaction()
            tIDHeader = mdt.Rows(0)("NoSPP")
            If Not New SQLSPP().CheckTime(tReadTime, tIDHeader) Then
                Throw New Exception("Data harus di refresh")
            End If
            'Header 
            addArrayAll(z, mdt.Rows(0), "NoSPP", "TglSPP", "CustomerID", "TglKirim", "NamaPenerima", "AlamatPenerima", "MataUang", "WaktuPembayaran", "KeteranganSPP", "Status", "IDTrans")
            z.Add(New Param("Pengupdate", gUser))
            If tReadTime.Ticks = 0 Then
                s = "insert into tSPP(Tipe,NoSPP, TglSPP, CustomerID, TglKirim, NamaPenerima, AlamatPenerima, MataUang, WaktuPembayaran, KeteranganSPP, Status, Pengupdate, WaktuUpdate) " & _
                " select Tipe, @NoSPP, @TglSPP, @CustomerID, @TglKirim, @NamaPenerima, @AlamatPenerima, @MataUang, @WaktuPembayaran, @KeteranganSPP, 0, @Pengupdate, now() from mCurrentTipe"
            Else
                s = "update tSPP set NoSPP=@NoSPP, TglSPP=@TglSPP, CustomerID=@CustomerID, TglKirim=@TglKirim, NamaPenerima=@NamaPenerima, AlamatPenerima=@AlamatPenerima, MataUang=@MataUang, WaktuPembayaran=@WaktuPembayaran, KeteranganSPP=@KeteranganSPP, Status=@Status, Pengupdate=@Pengupdate, WaktuUpdate=now() where IDTrans=@IDTrans"
            End If
            a.execMe(s, z)


            Dim dt As DataTable = mdt2.GetChanges

            If Not dt Is Nothing Then
                Dim dv As New DataView(dt, "", "", DataViewRowState.Deleted)
                For Each dvrow As DataRowView In dv
                    z.Clear()
                    z.Add(New Param("IDTrans", dvrow("IDTrans")))
                    s = "delete from tSPPDetail where IDTrans=@IDTrans"
                    If a.execMe(s, z) = 0 Then
                        Throw New Exception(dvrow("IDTrans") & " : Data terpakai")
                    End If

                Next

                For Each row As DataRow In dt.Rows
                    If row.RowState = DataRowState.Deleted Then Continue For
                    addArrayAll(z, row, "IDTrans", "JumlahBox", "Harga", "KeteranganSPPDetail", "IDSC", "IDStock")
                    z.Add(New Param("Pengupdate", gUser))
                    z.Add(New Param("NoSPP", tIDHeader))
                    If row.RowState = DataRowState.Added Then
                        s = "insert into tSPPDetail(NoSPP, JumlahBox, Harga, KeteranganSPPDetail, IDSC, IDStock, Pengupdate, WaktuUpdate)" & _
                        " select @NoSPP, @JumlahBox, @Harga, @KeteranganSPPDetail, @IDSC, @IDStock, @Pengupdate, now()"
                    Else
                        s = "update tSPPDetail set NoSPP=@NoSPP, JumlahBox=@JumlahBox, Harga=@Harga, KeteranganSPPDetail=@KeteranganSPPDetail, IDSC=@IDSC, IDStock=@IDStock, Pengupdate=@Pengupdate, WaktuUpdate=now() where IDTrans=@IDTrans"
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
    Public Function SetStatus(ByVal tNoSPP As String, ByVal st As Integer) As Boolean
        Try
            Dim s As String
            s = "update tSPP set status=" & st & " where NoSPP='" & tNoSPP & "'"
            a.execMe(s)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function CheckTime(ByVal tReadTime As DateTime, ByVal tNoSPP As String) As Boolean
        Dim s As String
        Dim p As New List(Of Param)
        Dim dt As DataTable
        s = "select * from vtspp where NoSPP='" & tNoSPP & "' and WaktuUpdate>@WaktuUpdate"
        p.Add(New Param("WaktuUpdate", tReadTime))
        dt = a.doQuery(s, p)
        If dt.Rows.Count > 0 Then
            Return False
        End If
        Return True
    End Function
End Class

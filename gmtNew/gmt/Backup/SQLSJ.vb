Public Class SQLSJ
    Dim a As New db
    Public Function qLastNoSJ() As Object
        Dim s As String
        s = "select min(toint(left(NoSPP,5))) from vtspp where status=1"
        Return a.doQueryScalar(s)
    End Function
    Public Function qData(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        s = "select IDTrans, NoSPP, TglSPP, TglKirim, NamaPenerima, AlamatPenerima, KeteranganSPP, Pengupdate, WaktuUpdate, TglSJ, NamaSopir, NamaAngkutan, NoKendaraan, KeteranganSJ,Status from vtspp where 1=1" & setFilter(z)
        Return a.doQuery(s)
    End Function
    Public Function qDataDetail(ByVal tNoSPP As String) As DataTable
        Dim s As String
        s = "select a.IDTrans, a.IDStock, b.Jenis, b.KodeBarang, b.Warna, b.NoWarna, b.Tube, b.Grade, a.JumlahBox, a.JumlahKG, a.KeteranganSPPDetail from tSPPDetail a " & _
        " left join mStock b on a.IDStock=b.IDStock" & _
        " where a.NoSPP='" & esc(tNoSPP) & "'"
        Return a.doQuery(s)
    End Function

    Public Function Save(ByVal mdt As DataTable, ByVal mdt2 As DataTable, ByVal tReadTime As DateTime, Optional ByVal tID As Integer = 0) As Boolean
        Dim tIDHeader As String
        Dim s As String = ""
        Dim tSJ As String
        Dim tTotal As Double
        Dim z As New List(Of Param)
        Try
            a.BeginTransaction()
            tIDHeader = mdt.Rows(0)("NoSPP")
            If Not New SQLSPP().CheckTime(tReadTime, tIDHeader) Then
                Throw New Exception("Data harus di refresh")

            End If
            

            Dim dt As DataTable = mdt2

            If Not dt Is Nothing Then
                For Each row As DataRow In dt.Rows
                    If row.RowState = DataRowState.Deleted Then Continue For
                    addArrayAll(z, row, "IDTrans", "JumlahKG", "IDStock")
                    z.Add(New Param("Pengupdate", gUser))
                    z.Add(New Param("NoSPP", tIDHeader))
                    s = "update mStock a inner join tSPPDetail b on a.IDStock=b.IDStock set a.JumlahKG=a.JumlahKG-@JumlahKG where IDTrans=@IDTrans"
                    a.execMe(s, z)
                    s = "update tSPPDetail set JumlahKG=@JumlahKG where IDTrans=@IDTrans"
                    a.execMe(s, z)
                Next
            End If

            tSJ = Left(tIDHeader, 5) & New SQLPenomoran().qData("SJ").Rows(0)("Penomoran") & Format(mdt.Rows(0)("TglSPP"), "MM/yy")
            'Header 
            addArrayAll(z, mdt.Rows(0), "IDTrans", "TglSJ", "NamaSopir", "NamaAngkutan", "NoKendaraan", "KeteranganSJ")
            z.Add(New Param("Pengupdate", gUser))
            z.Add(New Param("NoSJ", tSJ))
            s = "select sum(JumlahKG*Harga) from tSPPDetail where NoSPP='" & tIDHeader & "'"
            tTotal = a.doQueryScalar(s)
            z.Add(New Param("Total", tTotal))
            s = "update tSPP set Total=@Total, Status=2, NoSJ=@NoSJ, TglSJ=@TglSJ, NamaSopir=@NamaSopir, NamaAngkutan=@NamaAngkutan, NoKendaraan=@NoKendaraan, KeteranganSJ=@KeteranganSJ where IDTrans=@IDTrans"
            a.execMe(s, z)

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
    Public Function Cancel(ByVal tNoSPP As String) As Boolean
        Dim s As String
        Try
            a.BeginTransaction()
            s = "update mStock a inner join tSPPDetail b on a.IDStock=b.IDStock set a.JumlahKG=a.JumlahKG+b.JumlahKG where NoSPP='" & tNoSPP & "'"
            a.execMe(s)
            s = "update tSPP set Status=1 where NoSPP='" & tNoSPP & "'"
            a.execMe(s)
            a.CommitTransaction()
            Return True
        Catch ex As Exception
            a.RollbackTransaction()
            Return False
        End Try

    End Function
End Class

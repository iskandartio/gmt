Imports System.Data.SqlClient
Imports System.Reflection

Public Class SQLmCustomer
    Dim a As New db
    Private Function esc(ByVal s As String) As String
        Return Replace(s, "'", "''")
    End Function

    Public Function qData(Optional ByVal z As List(Of Param) = Nothing) As DataTable
        Dim s As String
        
        s = "select * from mCustomer where 1=1" & setFilter(z) & " order by NamaCustomer"
        Return a.doQuery(s, z)
    End Function
    Public Function qDataPenerima(ByVal tCustomerID As Integer) As DataTable
        Dim s As String
        s = "select * from mPenerima where CustomerID='" & tCustomerID & "' order by Nama"
        Return a.doQuery(s)
    End Function
    Public Function UpdateData(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            Dim dv As New DataView(dt, "", "", DataViewRowState.Deleted)
            For Each dvrow As DataRowView In dv
                tArr.Add(New param("CustomerID", "CustomerID"))
                s = "delete from mCustomer where CustomerID=@CustomerID"
                a.ExecMe(s, tArr)
            Next
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "CustomerID", "NamaCustomer", "Alamat", "Kota", "Telepon", "Fax", "ContactPerson", "WaktuPembayaran", "MataUang", "AlamatPendek", "IsActive", "LimitKredit")
                If row.RowState = DataRowState.Modified Then
                    s = "update mCustomer set CustomerID=@CustomerID, NamaCustomer=@NamaCustomer, Alamat=@Alamat, Kota=@Kota, Telepon=@Telepon, Fax=@Fax, ContactPerson=@ContactPerson, WaktuPembayaran=@WaktuPembayaran, MataUang=@MataUang, AlamatPendek=@AlamatPendek, IsActive=@IsActive, LimitKredit=@LimitKredit" & _
                    " where CustomerID=@CustomerID"
                    a.ExecMe(s, tArr)
                ElseIf row.RowState = DataRowState.Added Then
                    s = "insert into mCustomer(CustomerID, NamaCustomer, Alamat, Kota, Telepon, Fax, ContactPerson, WaktuPembayaran, MataUang, AlamatPendek, IsActive, LimitKredit)" & _
                    " values(@CustomerID, @NamaCustomer, @Alamat, @Kota, @Telepon, @Fax, @ContactPerson, @WaktuPembayaran, @MataUang, @AlamatPendek, @IsActive, @LimitKredit)"
                    a.ExecMe(s, tArr)
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
    Public Function UpdatePenerima(ByVal mdt As DataTable) As Boolean
        Dim s As String
        Dim dt As DataTable = mdt.GetChanges
        Dim tArr As New List(Of Param)
        Try
            a.BeginTransaction()
            Dim dv As New DataView(dt, "", "", DataViewRowState.Deleted)
            For Each dvrow As DataRowView In dv
                tArr.Add(New Param("PenerimaID", dvrow("PenerimaID")))
                s = "delete from mPenerima where PenerimaID=@PenerimaID"
                a.execMe(s, tArr)
            Next
            tArr.Clear()
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then Continue For
                addArrayAll(tArr, row, "CustomerID", "Nama", "Alamat", "AlamatPendek", "Telepon", "Fax", "ContactPerson", "NoUrut", "Kota", "PenerimaID")
                If row.RowState = DataRowState.Modified Then
                    s = "update mPenerima set CustomerID=@CustomerID, Nama=@Nama, Alamat=@Alamat, AlamatPendek=@AlamatPendek, Telepon=@Telepon, Fax=@Fax, ContactPerson=@ContactPerson, NoUrut=@NoUrut, Kota=@Kota" & _
                    " where PenerimaID=@PenerimaID"
                    a.execMe(s, tArr)
                ElseIf row.RowState = DataRowState.Added Then
                    s = "insert into mPenerima(CustomerID, Nama, Alamat, AlamatPendek, Telepon, Fax, ContactPerson, NoUrut, Kota)" & _
                    " values(@CustomerID, @Nama, @Alamat, @AlamatPendek, @Telepon, @Fax, @ContactPerson, @NoUrut, @Kota)"
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

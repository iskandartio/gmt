﻿Public Class dbOLEDB
    Private conn As New OleDb.OleDbConnection
    Private cmd As New OleDb.OleDbCommand
    Private mConnectionString As String = ""

    Public Property ConnectionString() As String
        Get
            Return mConnectionString
        End Get
        Set(ByVal value As String)
            mConnectionString = value
        End Set
    End Property

    Public Sub BeginTransaction()

        If conn.State = ConnectionState.Closed Then ConnectDB()
        cmd.Transaction = conn.BeginTransaction()
    End Sub

    Public Sub CommitTransaction()
        cmd.Transaction.Commit()
        conn.Dispose()
    End Sub

    Public Sub RollbackTransaction()
        cmd.Transaction.Rollback()
    End Sub

    Public Sub ConnectDB()
        conn.ConnectionString = ConnectionString
        conn.Open()
        cmd.CommandTimeout = 0
        cmd.Connection = conn
    End Sub

    Public Function doQuery(ByVal tSQL As String, Optional ByVal arr As List(Of Param) = Nothing, Optional ByVal tReadonly As Boolean = False, Optional ByVal tSetDefault As Boolean = True) As DataTable
        Dim dt As New DataTable
        If conn.State = ConnectionState.Closed Then ConnectDB()
        cmd.Parameters.Clear()
        If Not arr Is Nothing Then
            For Each a As Param In arr
                cmd.Parameters.AddWithValue(a.ParamName, a.ParamValue)
            Next
        End If
        tSQL = Replace(tSQL, "~", pTipe)
        cmd.CommandText = tSQL
        dt.BeginLoadData()

        dt.Load(cmd.ExecuteReader)
        If tSetDefault Then
            Me.setDefault(dt, tReadonly)
        End If
        dt.AcceptChanges()
        dt.EndLoadData()
        If cmd.Transaction Is Nothing Then
            'conn.Dispose()
        End If
        Return dt
    End Function

    Public Function doQueryReader(ByVal tSQL As String, Optional ByVal arr As List(Of Param) = Nothing, Optional ByVal tReadonly As Boolean = False) As IDataReader
        Dim rdr As IDataReader
        If conn.State = ConnectionState.Closed Then ConnectDB()
        cmd.Parameters.Clear()
        If Not arr Is Nothing Then
            For Each a As Param In arr
                cmd.Parameters.AddWithValue(a.ParamName, a.ParamValue)
            Next
        End If
        cmd.CommandText = tSQL
        rdr = cmd.ExecuteReader
        If cmd.Transaction Is Nothing Then
            conn.Dispose()
        End If
        Return rdr
    End Function
    Public Function doQueryScalar(ByVal tSQL As String, Optional ByVal arr As List(Of Param) = Nothing, Optional ByVal tReadonly As Boolean = False) As Object
        Dim t As Object
        If conn.State = ConnectionState.Closed Then ConnectDB()
        cmd.Parameters.Clear()
        If Not arr Is Nothing Then
            For Each a As Param In arr
                cmd.Parameters.AddWithValue(a.ParamName, a.ParamValue)
            Next
        End If
        tSQL = Replace(tSQL, "~", pTipe)
        cmd.CommandText = tSQL
        t = cmd.ExecuteScalar
        If t Is Nothing Then Return DBNull.Value Else Return t
    End Function

    Private Sub setDefault(ByRef dt As DataTable, ByVal tReadOnly As Boolean)
        Dim col As DataColumnCollection = dt.Columns
        Dim tType As String
        For i As Integer = 0 To col.Count - 1
            col(i).ReadOnly = tReadOnly
            tType = col(i).DataType.Name
            If col(i).AllowDBNull AndAlso tType <> "DateTime" Then Continue For
            If tType = "DateTime" Then
                col(i).DefaultValue = #1/1/1900#
            ElseIf tType = "String" Then
                col(i).DefaultValue = ""
            ElseIf tType.Contains("Int") OrElse tType = "Double" OrElse tType = "Decimal" OrElse tType = "Single" Then
                If Not col(i).AutoIncrement Then col(i).DefaultValue = 0
            ElseIf tType = "Boolean" Then
                col(i).DefaultValue = False
            Else
                MsgBox(col(i).DataType.Name)
            End If
            If Not col(i).AutoIncrement Then
                For j As Integer = 0 To dt.Rows.Count - 1
                    If IsDBNull(dt.Rows(j)(i)) Then
                        If Not tReadOnly Then dt.Rows(j)(i) = col(i).DefaultValue
                    End If
                Next
            End If
        Next
        dt.AcceptChanges()
    End Sub

    Public Function doQuery2(ByVal tSQL As String, Optional ByVal arr As List(Of Param) = Nothing, Optional ByVal tReadonly As Boolean = False) As List(Of DataTable)
        Dim rdr As OleDb.OleDbDataReader
        Dim dt As DataTable
        Dim listDT As New List(Of DataTable)
        If conn.State = ConnectionState.Closed Then ConnectDB()
        cmd.Parameters.Clear()
        If Not arr Is Nothing Then
            For Each a As Param In arr
                cmd.Parameters.AddWithValue(a.ParamName, a.ParamValue)
            Next
        End If
        cmd.CommandText = tSQL
        rdr = cmd.ExecuteReader
        While Not rdr.IsClosed
            dt = New DataTable
            dt.Load(rdr)
            For i As Integer = 0 To dt.Columns.Count - 1
                dt.Columns(i).ReadOnly = tReadonly
                If dt.Columns(i).DataType.Name = "DateTime" Then
                    dt.Columns(i).DefaultValue = #1/1/1900#
                End If
            Next
            listDT.Add(dt)
        End While
        Return listDT
    End Function

    Public Function execMe(ByVal tSQL As String, Optional ByVal arr As List(Of Param) = Nothing, Optional ByVal tReturnID As Boolean = False) As Integer
        If conn.State = ConnectionState.Closed Then ConnectDB()
        cmd.Parameters.Clear()
        If Not arr Is Nothing Then
            For Each a As Param In arr
            
                If a.ParamValue.GetType.Name.Contains("Int") Then
                    tSQL = Replace(tSQL, a.ParamName, a.ParamValue)
                ElseIf a.ParamValue.GetType.Name.Contains("Double") Then
                    tSQL = Replace(tSQL, a.ParamName, a.ParamValue.ToString.Replace(",", "."))
                Else
                    cmd.Parameters.AddWithValue(a.ParamName, a.ParamValue)
                End If

            Next
        End If
        cmd.CommandTimeout = 0
        cmd.CommandType = CommandType.Text
        tSQL = Replace(tSQL, "~", pTipe)
        cmd.CommandText = tSQL
        Dim i As Integer = cmd.ExecuteNonQuery()
        If tReturnID Then
            Return Convert.ToInt32(doQuery("select @@IDENTITY").Rows(0)(0))
        Else
            Return i
        End If
    End Function

    Public Sub New(Optional ByVal s As String = "")
        Dim s2 As String
        Dim i As Integer
        Dim i2 As Integer
        Dim EncryptedPassword As String

        's2 = Encrypt("joTEsAKa6810")
        's2 = Encrypt("xaMEfIKa1587")
        's2 = Encrypt("123456")
        s2 = Encrypt("gmt")
        If s = "" Then s = My.Application.Info.AssemblyName

        s2 = System.Configuration.ConfigurationManager.ConnectionStrings(My.Application.Info.AssemblyName & ".My.MySettings." & s & "ConnectionString").ConnectionString
        i = s2.IndexOf("Password=")
        i2 = s2.IndexOf(";", i)
        EncryptedPassword = s2.Substring(i, i2 - i)
        ConnectionString = Replace(s2, EncryptedPassword, "Password=" & ModFunction.Decrypt(EncryptedPassword.Substring(9)))
        'ConnectionString = Replace(s2, EncryptedPassword, "Password=" & EncryptedPassword.Substring(9))

    End Sub


End Class

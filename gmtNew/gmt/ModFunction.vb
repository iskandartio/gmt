
Public Module ModFunction
    Public Function convertToInteger(ByVal tDate As Date) As Integer
        Return Format(tDate, "yyMMdd")
    End Function
    Public Function convertToDate(ByVal tDate As Integer) As String
        Return Format(DateSerial(2000 + tDate / 10000, (tDate Mod 10000) \ 100, tDate Mod 100), "dd MMM yyyy")
    End Function
    Public Function convertToDate(ByVal tDate As String) As String
        If tDate = "" Then Return ""
        Return Format(DateSerial(2000 + tDate / 10000, (tDate Mod 10000) \ 100, tDate Mod 100), "dd MMM yyyy")
    End Function

    Public Function setFilter(Optional ByVal z As List(Of Param) = Nothing, Optional ByVal tStartWith As Boolean = False) As String
        Dim tFilter As String = ""
        If z Is Nothing Then Return ""
        For Each p As Param In z
            If TypeOf (p.ParamValue) Is String Then
                If p.Exact Then
                    If p.ParamValue = "=" Then p.ParamValue = DBNull.Value
                Else
                    If p.ParamValue = "" Then p.ParamValue = DBNull.Value
                End If

            End If
            If p.Exact Then
                If Not IsDBNull(p.ParamValue) Then tFilter &= " and " & p.ParamName & " ='" & esc(p.ParamValue.ToString.Substring(1)) & "'"
            Else
                If Not IsDBNull(p.ParamValue) Then tFilter &= " and " & p.ParamName & " like '" & IIf(tStartWith, "", "%") & esc(p.ParamValue) & "%'"
            End If
        Next
        Return tFilter
    End Function
    Public Function setCurrentValue(ByVal dt As DataTable) As Object
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)(0)
        End If
        Return ""
    End Function
    Public Function setParam(ByVal ParamArray a As Object()) As List(Of Param)
        Dim z As New List(Of Param)
        Dim i As Integer = 0
        z.Clear()
        While i < a.Length
            If Left(a(i + 1), 1) = "=" Then
                z.Add(New Param(a(i), a(i + 1), True))
            Else
                z.Add(New Param(a(i), a(i + 1)))
            End If
            i = i + 2
        End While
        Return z
    End Function

    Public Function ToNPWP(ByVal tNPWP As String) As String
        Dim sb As New System.Text.StringBuilder
        sb.Append(tNPWP)
        If sb.Length > 14 Then
            sb.Insert(2, ".", 1).ToString()
            sb.Insert(6, ".", 1).ToString()
            sb.Insert(10, ".", 1).ToString()
            sb.Insert(12, "-", 1).ToString()
            sb.Insert(16, ".", 1).ToString()
        End If
        Return sb.ToString
    End Function
    Public Function ToNPWPRev(ByVal tNPWP As String) As String
        Return Replace(Replace(tNPWP, ".", ""), "-", "")
    End Function
    Public Function Encrypt(ByVal tText As String) As String
        Dim b() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(tText)
        Dim tKey As String = "PANIN LIFE ARMS"
        Dim c() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(tKey)
        Dim i As Integer
        Dim j As Integer = 0
        Dim k As Integer
        Dim m As Integer = 10
        For i = 0 To b.Length - 1
            k = (90 + b(i) - c(j) + m) Mod 90 + 32
            b(i) = k
            m = k
            j = j + 1
            If j = c.Length Then j = 0
        Next
        Return System.Text.ASCIIEncoding.ASCII.GetString(b)
    End Function
    Public Function Decrypt(ByVal tText As String) As String
        Dim b() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(tText)
        Dim tKey As String = "PANIN LIFE ARMS"
        Dim c() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(tKey)
        Dim i As Integer
        Dim j As Integer = 0
        Dim k As Integer
        Dim m As Integer = 10
        For i = 0 To b.Length - 1
            k = (90 + b(i) + c(j) - 32 - m) Mod 90
            If k < 32 Then k = k + 90
            m = b(i)
            b(i) = k
            j = j + 1
            If j = c.Length Then j = 0
        Next
        Return System.Text.ASCIIEncoding.ASCII.GetString(b)
    End Function
    Public Sub CopyDataToClipBoard(ByVal dt As DataTable, Optional ByVal tSuppressMessage As Boolean = False, Optional ByVal tCopyHeader As Boolean = True, Optional ByVal tCopyForExcel As Boolean = False)
        Dim s As New System.Text.StringBuilder
        If tCopyHeader Then
            For j As Integer = 0 To dt.Columns.Count - 1
                s.Append(dt.Columns(j).ColumnName & vbTab)
            Next
            s.AppendLine()
        End If

        For i As Integer = 0 To dt.Rows.Count - 1
            For j As Integer = 0 To dt.Columns.Count - 1
                If dt.Columns(j).DataType.Name = "String" Then
                    s.Append(Replace(dt.Rows(i)(j).ToString, vbCrLf, " ") & vbTab)
                Else
                    If tCopyForExcel And dt.Columns(j).DataType.Name = "DateTime" Then
                        If Not IsDBNull(dt.Rows(i)(j)) Then
                            s.Append(Format(dt.Rows(i)(j), "MM-dd-yyyy") & vbTab)
                        Else
                            s.Append(vbTab)
                        End If

                    Else
                        s.Append(dt.Rows(i)(j) & vbTab)
                    End If
                End If

            Next
            s.AppendLine()
        Next
        My.Computer.Clipboard.SetText(s.ToString)
        If Not tSuppressMessage Then MsgBox("Copied to Clipboard")
    End Sub

    Public Sub PasteToExcel(ByVal sht As Object, ByVal dt As DataTable, Optional ByVal tCopyHeader As Boolean = True)
        For i As Integer = 0 To dt.Columns.Count - 1
            If dt.Columns(i).DataType.Name = "String" Then
                sht.Columns(i + 1).NumberFormat = "@"
            ElseIf dt.Columns(i).DataType.Name = "DateTime" Then
                sht.Columns(i + 1).NumberFormat = "dd-mm-yyyy"
                sht.columns(i + 1).HorizontalAlignment = 4
            End If
        Next
        ModFunction.CopyDataToClipBoard(dt, True, tCopyHeader, True)
        sht.Paste()
        For i As Integer = 0 To dt.Columns.Count - 1
            If dt.Columns(i).DataType.Name = "Decimal" OrElse dt.Columns(i).DataType.Name = "Double" Then
                sht.Columns(i + 1).NumberFormat = "#,##0.00"
                sht.Columns(i + 1).Autofit()
            ElseIf dt.Columns(i).DataType.Name = "DateTime" Then
                sht.Columns(i + 1).Autofit()
            End If
        Next
    End Sub
End Module


Public Module ModFunction
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
    Public Sub CopyDataToClipBoard(ByVal dt As DataTable)
        Dim s As New System.Text.StringBuilder
        For j As Integer = 0 To dt.Columns.Count - 1
            s.Append(dt.Columns(j).ColumnName & vbTab)
        Next
        s.AppendLine()
        For i As Integer = 0 To dt.Rows.Count - 1
            For j As Integer = 0 To dt.Columns.Count - 1
                s.Append(dt.Rows(i)(j) & vbTab)
            Next
            s.AppendLine()
        Next
        My.Computer.Clipboard.SetText(s.ToString)
        MsgBox("Copied to Clipboard")
    End Sub
End Module

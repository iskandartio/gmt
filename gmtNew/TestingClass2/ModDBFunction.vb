Public Module ModDBFunction
    Public Function esc(ByVal s As String) As String
        Return Replace(s, "'", "''")
    End Function
    Public Sub addArrayAll(ByRef arr As ArrayList, ByVal row As DataRow, ByVal ParamArray args As String())
        arr.Clear()
        For i As Integer = 0 To args.GetUpperBound(0)
            arr.Add(New Param("@" + args(i), row(args(i))))
        Next        
    End Sub

End Module

Public Class Param
    Private mParam As String
    Public Property ParamName() As String
        Get
            Return mParam
        End Get
        Set(ByVal value As String)
            mParam = value
        End Set
    End Property

    Private mParamValue As Object
    Public Property ParamValue() As Object
        Get
            Return mParamValue
        End Get
        Set(ByVal value As Object)
            mParamValue = value
        End Set
    End Property

    Public Sub New(ByVal tParam As String, ByVal tParamValue As Object)
        mParam = tParam
        mParamValue = tParamValue
    End Sub

End Class
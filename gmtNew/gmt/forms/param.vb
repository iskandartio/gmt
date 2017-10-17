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

    Private mExact As Object
    Public Property Exact() As Object
        Get
            Return mExact
        End Get
        Set(ByVal value As Object)
            mExact = value
        End Set
    End Property
    Public Sub New(ByVal tParam As String, ByVal tParamValue As Object, Optional ByVal tExact As Boolean = False)
        mParam = tParam
        mParamValue = tParamValue
        mExact = tExact
    End Sub

End Class
Imports System.Reflection

Public Class FormFunctions
    Public Shared Function GetFormByName(ByVal FormName As String) As Form
        'first try: in case the full namespace has been provided (as it should  ) 
        Dim T As Type = Type.GetType(FormName, False)
        'if not found, search for it 
        If T Is Nothing Then T = FindType(FormName)
        'if still not found, throw exception 
        If T Is Nothing Then Throw New Exception(FormName + " could not be found")
        Return CType(Activator.CreateInstance(T), Form)
    End Function
#Region "Assemblies and types"
    Public Shared Function GetAllAssemblies() As ArrayList
        Dim al As New ArrayList
        Dim a As [Assembly] = [Assembly].GetEntryAssembly()
        FillAssemblies(a, al)
        Return al
    End Function

    Private Shared Sub FillAssemblies(ByVal a As [Assembly], ByVal al As ArrayList)
        If Not al.Contains(a) Then
            al.Add(a)
            Dim an As AssemblyName
            For Each an In a.GetReferencedAssemblies()
                If Not an.Name.StartsWith("System") Then FillAssemblies([Assembly].Load(an), al)
            Next
        End If
    End Sub

    Public Shared Function GetAllTypes() As ArrayList
        Dim a As [Assembly], t As Type, al As New ArrayList
        For Each a In GetAllAssemblies()
            For Each t In a.GetTypes
                If Not al.Contains(t) Then al.Add(t)
            Next
        Next
        Return al
    End Function

    Public Shared Function FindType(ByVal Name As String) As Type
        Dim T As Type
        For Each T In GetAllTypes()
            If T.Name = Name Then Return T
        Next
        Return Nothing
    End Function
#End Region
End Class

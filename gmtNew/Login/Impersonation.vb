Imports System
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Security.Permissions

<Assembly:SecurityPermissionAttribute(SecurityAction.RequestMinimum,UnmanagedCode := true)>
Public Class Impersonation

    <DllImport("C:\\WINDOWS\\System32\\advapi32.dll")> _
    Public Shared Function LogonUser(ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, ByRef phToken As Integer) As Boolean
    End Function

    <DllImport("C:\\WINDOWS\\System32\\Kernel32.dll")> _
    Public Shared Function GetLastError() As Integer
    End Function
End Class
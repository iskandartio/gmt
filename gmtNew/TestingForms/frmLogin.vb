Public Class frmLogin
    Inherits FormMain
    Dim a As New db
    Public gUser As String
    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        gUser = a.doQueryScalar("select UID from mUser where UID='" & txtUserID.Text & "' and Pwd=password('" & txtPassword.Text & "')")
        If Len(gUser) < 1 Then
            MsgBox("User ID atau Password Salah")
            Exit Sub
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Yes
        End If
    End Sub


End Class

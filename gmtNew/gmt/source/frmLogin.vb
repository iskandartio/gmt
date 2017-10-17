Public Class frmLogin
    Inherits FormMain
    Dim a As New db

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        If a.doQueryScalar("select * from mUser where UID='" & txtUserID.Text & "' and Pwd=password('" & txtPassword.Text & "')") Is Nothing Then
            MsgBox("User ID atau Password Salah")
            Exit Sub
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Yes
        End If
    End Sub


End Class

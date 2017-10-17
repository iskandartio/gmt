Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class frmLogin
    Inherits System.Windows.Forms.Form

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Login()
    End Sub
    Private Sub txtPassword_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            Login()
        End If
    End Sub

    Private Sub Login()
        Dim token1 As Int32
        Dim loggedOn As Boolean = Impersonation.LogonUser(txtUserID.Text, "Paninlife", txtPassword.Text, 3, 0, token1)
        If loggedOn Then
            Shell("MainMenu.exe " & Chr(1) & Chr(2) & Chr(3) & txtUserID.Text & "@" & My.Settings.cobaConnectionString, AppWinStyle.MaximizedFocus)
            Me.Dispose()
        End If
    End Sub
End Class
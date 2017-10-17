Imports System.Windows.Forms

Public Class FormMain
    Inherits System.Windows.Forms.Form
    
    Private Sub FormMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If TypeName(ActiveControl) = "StringTextBoxNoKeyPreview" Then Exit Sub
        If e.KeyCode = Keys.Decimal Then
            e.SuppressKeyPress = True
            SendKeys.Send(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator)
        ElseIf e.KeyCode = Keys.Enter Then
            SendKeys.Send(vbTab)
        End If
    End Sub

    Private Sub FormMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.KeyPreview = True
    End Sub

End Class


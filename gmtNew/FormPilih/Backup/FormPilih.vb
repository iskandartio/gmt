Imports System.Windows.Forms

Namespace PaninProject
    Public Class FormPilih
#Region "Events"
        Private Sub FormPilih_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
            If e.KeyCode = Keys.Escape Then
                CloseForm()
            End If
        End Sub

        Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
            CloseForm()
        End Sub
        Private Sub CloseForm()
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Dispose()
        End Sub
        Private Sub dg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dg.KeyDown
            If e.KeyCode = Keys.Enter Then
                Choose()
            End If
        End Sub
        Private Sub btnChoose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChoose.Click
            Choose()
        End Sub
#End Region

#Region "Function"
        Overridable Sub Choose()

        End Sub
#End Region

    End Class
End Namespace

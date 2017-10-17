Public Class NumericTextBox
    Inherits System.Windows.Forms.TextBox
    Private mNumberFormat As String

    Public Property Value() As Object
        Get
            Return CDec(Me.Text)
        End Get
        Set(ByVal value As Object)
            If IsDBNull(value) Then
                value = 0
            End If
            Me.Text = Format(value, Me.NumberFormat)
        End Set
    End Property
    Public Property NumberFormat() As String
        Get
            Return mNumberFormat
        End Get
        Set(ByVal value As String)
            mNumberFormat = value
        End Set
    End Property

    Private Sub NumericTextBox_BindingContextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.BindingContextChanged
        If Me.DataBindings.Count > 0 Then
            Me.DataBindings(0).DataSourceNullValue = 0
        End If
    End Sub

    Private Sub _GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.Text = CDec(Me.Text)
    End Sub

    
End Class

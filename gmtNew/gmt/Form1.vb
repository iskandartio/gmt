Public Class Form1
    Dim dt2 As DataTable
    Dim bs As New BindingSource
    Dim a As New db
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dt As DataTable
        dt2 = a.doQuery("select NamaCustomer from tSC a left join mCustomer b on a.CustomerID=b.CustomerID where a.NoSC='" & TextBox1.Text & "'")

        If ComboBox1.DataSource Is Nothing Then
            dt = a.doQuery("select NamaCustomer, Alamat from mCustomer")
            ComboBox1.LoadMe(dt, "NamaCustomer", "NamaCustomer@Alamat", "200@200")

        End If
        bs.DataSource = dt2
        ComboBox1.DataBindings.Clear()
        ComboBox1.DataBindings.Add("SelectedValue", bs, "NamaCustomer")

    End Sub

    Private Sub ComboBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Enter
        If ComboBox1.DataSource Is Nothing Then
            Dim dt As DataTable
            dt = a.doQuery("select NamaCustomer, Alamat from mCustomer")
            ComboBox1.LoadMe(dt, "NamaCustomer", "NamaCustomer@Alamat", "200@200")

        End If
    End Sub
End Class
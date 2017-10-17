Public Class FormPilihSingle
    Private dt As DataTable
    Public mTipe As String

    Private Sub FindMe(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, btnFind.Click
        Find()
    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        ModControl.ClearControl(grpFilter)
    End Sub
    Private Sub _KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyDown
        If e.KeyCode = Keys.Down OrElse e.KeyCode = Keys.Up Then
            dg.Focus()
        End If
    End Sub

    Private Sub Find()
        Dim z As New List(Of Param)
        z.Add(New Param(mTipe, txtCode.Text))
        dt = New SQLmStock().qDataSingle(mTipe, z)
        dg.DataSource = dt
        If dt.Rows.Count = 0 Then
            MsgBox("No Record Found")
        End If
        ModControl.setGridAutoInvisible(dg)
        ModControl.setGridStyle(dg)
    End Sub

    Overrides Sub Choose()
        Me.DialogResult = Windows.Forms.DialogResult.OK
        row = dt.Rows(dg.CurrentCell.RowIndex)
        Me.Hide()
    End Sub

    Public Sub New(ByVal tTipe As String, ByVal tDefaultCode As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        txtCode.Text = tDefaultCode
        mTipe = tTipe
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
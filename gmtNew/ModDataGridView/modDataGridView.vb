Public Class modDataGridView
    'this tells this class that is a datagridview
    Inherits System.Windows.Forms.DataGridView

    'this function do the half of job 
    'when you hit a key (like return) it transforms like you 
    'have been hit the tab key
    'this doent work when you are editing the cell
    Private mStr As String
    Private mLastSearch As String
    Public mFalseKey As Boolean
    Private mIsOnPilih As Boolean
    Public Property IsOnPilih() As Boolean
        Get
            Return mIsOnPilih
        End Get
        Set(ByVal value As Boolean)
            mIsOnPilih = value
        End Set
    End Property

    Public Property LastSearch() As String
        Get
            Return mLastSearch
        End Get
        Set(ByVal value As String)
            mLastSearch = value
        End Set
    End Property

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        mFalseKey = False
        If keyData = Keys.Return Then
            keyData = Keys.Tab
            mFalseKey = True
            With msg
                .WParam = Keys.Tab
            End With
        ElseIf keyData = Keys.F3 Then
            Find(mLastSearch)
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    'this second function works when you are editing the cell
    'does the same as the function above changing return for tab
    Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean

        If keyData = Keys.Return Then
            keyData = Keys.Tab
        End If

        Return MyBase.ProcessDialogKey(keyData)
    End Function

    
    Private Sub modDataGridView_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Control And e.KeyCode = Keys.V Then
            Dim s() As String, j As Integer, tCol As Integer
            s = Split(Clipboard.GetText, vbCrLf)
            j = Me.CurrentCell.RowIndex
            tCol = Me.CurrentCell.ColumnIndex
            For i As Integer = 0 To s.Length - 1
                If Me.CurrentCell.ReadOnly Then Exit Sub
                Me.CurrentCell = Me.Rows(j).Cells(tCol)
                Me.BeginEdit(False)
                Me.CurrentCell.Value = s(i)
                Me.EndEdit()
                j = j + 1
                If j = Me.Rows.Count - 1 Then Exit For
            Next
        End If
    End Sub

    Private Sub modDataGridView_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If Not Me.CurrentCell.ReadOnly Then Exit Sub
        mStr += e.KeyChar
        Timer1.Enabled = False
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If mStr = "" Then Exit Sub
        mLastSearch = mStr
        Find(mStr)
        Timer1.Enabled = False
        mStr = ""
    End Sub

    Public Sub Find(ByVal s As String)
        If s = "" Then Exit Sub
        Dim colIndex As Integer, i As Integer, Found As Boolean
        colIndex = Me.CurrentCell.ColumnIndex
        i = Me.CurrentRow.Index
        If i >= Me.RowCount - 1 Then i = -1
        Found = False
        Do Until i + 1 = Me.CurrentRow.Index Or Found
            i = i + 1
            If Not Me.Rows(i).Cells(colIndex).Value Is Nothing AndAlso Me.Rows(i).Cells(colIndex).Value.ToString.IndexOf(s, StringComparison.OrdinalIgnoreCase) <> -1 Then
                Me.CurrentCell = Me.Rows(i).Cells(colIndex)
                Found = True
            End If
            If i + 1 = Me.RowCount Then i = -1
        Loop
    End Sub
End Class



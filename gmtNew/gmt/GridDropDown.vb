Public Class GridDropDown
    Implements Popup.IPopupUserControl
    Dim mCtrl As Control
    Dim mdt As DataTable
    Dim dt As DataTable
    Dim mSQL As String
    Dim mField As String
    Dim mSearchType As String
    Dim mFirst As Boolean = False
    Public mFromServer As Boolean = False

    Public Function AcceptPopupClosing() As Boolean Implements Popup.IPopupUserControl.AcceptPopupClosing
        Return True
    End Function
    Public Sub SetDB(ByVal tSQL As String, ByVal tField As String, Optional ByVal tSearchType As String = "like", Optional ByVal tText As String = "")
        Dim s As String
        Dim a As New db
        mSQL = tSQL
        mField = tField
        mSearchType = tSearchType
        If tText <> "" Then
            tSQL = tSQL & " where " & mField & " like '" & IIf(tSearchType = "like", "%", "") & esc(tText) & IIf(tSearchType <> "exact", "%", "") & "'"
            If mFromServer Then mdt = a.doQuery(tSQL)
        Else
            mdt = a.doQuery(tSQL)
        End If

        If dt Is Nothing Then dt = a.doQuery(tSQL & " where 1=0")
        mdt.CaseSensitive = False
    End Sub
    Public Sub RefreshData(ByVal tText As String)
        If mFromServer Then
            SetDB(mSQL, mField, mSearchType, tText)

            dg.DataSource = mdt
        Else
            mFirst = True
            dg.DataSource = Nothing
            dt.Clear()
            Dim datarows As DataRow()
            datarows = Nothing
            If mSearchType = "like" Then
                datarows = mdt.Select(mField & " like '%" & esc(tText) & "%'")
            ElseIf mSearchType = "StartWith" Then
                datarows = mdt.Select(mField & " like '" & esc(tText) & "%'")
            ElseIf mSearchType = "exact" Then
                datarows = mdt.Select(mField & " = '" & esc(tText) & "'")
            End If
            For Each row As DataRow In datarows
                dt.ImportRow(row)
            Next
            mFirst = True
            dg.DataSource = dt

            End If
    End Sub
    Private Sub dg_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dg.SelectionChanged
        If mCtrl Is Nothing Then Exit Sub
        If mFirst Then
            mFirst = False
            Exit Sub
        End If
        If dg.DataSource.rows.count = 0 Then Exit Sub
        If IsDBNull(dg.CurrentCell.Value) Then Exit Sub
        mCtrl.Text = dg.CurrentCell.Value
    End Sub
    Public Sub SetControl(ByRef ctrl As Control)
        mCtrl = ctrl
    End Sub
End Class

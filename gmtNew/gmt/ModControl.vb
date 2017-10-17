Public Module ModControl
    Public Sub putGridAll(ByRef dg As DataGridView, ByVal row As DataRow, ByVal ParamArray args As String())
        For i As Integer = 0 To args.GetUpperBound(0)
            dg.CurrentRow.Cells(args(i)).Value = row(args(i))
        Next
    End Sub
    Public Sub KeepRowState(ByRef dt As DataTable, ByVal tPos As Integer)
        Dim rowState As DataRowState
        If tPos = -1 Then Exit Sub
        rowState = dt.Rows(tPos).RowState
        dt.AcceptChanges()
        If rowState = DataRowState.Added Then
            dt.Rows(tPos).SetAdded()
        Else
            dt.Rows(tPos).SetModified()
        End If
    End Sub
    Public Sub setGridVisible(ByRef dg As DataGridView, ByVal tVal As Boolean, ByVal ParamArray args As String())
        For i As Integer = 0 To args.GetUpperBound(0)
            If args(0) = "0" Then
                For j As Integer = 0 To dg.Columns.Count - 1
                    dg.Columns(j).Visible = tVal
                Next
                Exit Sub
            End If
            dg.Columns(args(i)).Visible = tVal
        Next

    End Sub
    Public Sub setGridAutoInvisible(ByRef dg As DataGridView, ByVal ParamArray headers As String())
        Dim j As Integer = 0
        For i As Integer = 0 To dg.ColumnCount - 1
            If UCase(Left(dg.Columns(i).Name, 2)) = "ID" Then
                dg.Columns(i).Visible = False
            Else
                If j < headers.Length Then
                    dg.Columns(i).HeaderText = headers(j)

                    j = j + 1
                End If
            End If
        Next

    End Sub
    Public Sub setGridDataSource(ByRef dg As DataGridView, ByVal dt As DataTable)
        dg.Columns.Clear()
        dg.DataSource = dt
    End Sub
    Sub CopyGrid(ByRef GridView As DataGridView)
        Dim i As Integer, j As Integer, s As String, s1 As String, c As DataGridViewColumn
        If GridView.SelectedColumns.Count = 0 Then
            If MsgBox("Copy all values in grid? You may select spesific column by clicking on column header", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            Else
                For Each c In GridView.Columns
                    c.Selected = True
                Next
            End If
        End If
        Clipboard.Clear()
        s1 = ""

        For j = 0 To GridView.RowCount - 1
            s = ""
            For i = 0 To GridView.SelectedColumns.Count - 1
                s = s & vbTab & GridView(j, i).Value.ToString
            Next
            s1 = s1 & vbCrLf & Mid(s, 2)
        Next
        Clipboard.SetText(Mid(s1, 3))
    End Sub

    Sub FormKeyDown(ByRef e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Decimal Then
            e.SuppressKeyPress = True
            SendKeys.Send(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator)
        ElseIf e.KeyCode = Keys.Enter Then
            SendKeys.Send(vbTab)
        End If
    End Sub

    Sub LockMe(ByVal tObj As Object, ByVal tVal As Boolean)
        If Not tObj.tabstop And Not TypeOf tObj Is ModComboBox Then Exit Sub
        If TypeOf tObj Is TabControl OrElse TypeOf tObj Is Label Then Exit Sub
        If TypeOf tObj Is modComboBox Then
            tObj.enabled = Not tVal
            If Not tVal Then
                tObj.BackColor = Color.White
            Else
                tObj.BackColor = tObj.DisabledBackColor
            End If
        ElseIf TypeOf tObj Is TextBox Then
            tObj.ReadOnly = tVal
            If tVal Then
                tObj.BackColor = Color.Gainsboro
            Else
                tObj.BackColor = Color.White
            End If
        ElseIf TypeOf tObj Is DateTimePicker Then

            tObj.Enabled = Not tVal
            If tVal Then
                tObj.BackColor = Color.White

            Else
                tObj.BackColor = tObj.DisabledBackColor
            End If
        ElseIf TypeOf tObj Is Label Then
            Exit Sub
        ElseIf TypeOf tObj Is Button Then
            tObj.Enabled = Not tVal
            Exit Sub
        ElseIf TypeOf tObj Is CheckBox Then
            tObj.AutoCheck = Not tVal
        ElseIf TypeOf tObj Is DataGridView Then
            Dim dg As DataGridView
            dg = tObj
            dg.ReadOnly = tVal
            For i As Integer = 0 To dg.Columns.Count - 1
                If Not dg.DataSource.Columns(i).ReadOnly Then
                    dg.Columns(i).ReadOnly = tVal
                End If
            Next
            dg.AllowUserToDeleteRows = Not tVal
            If tVal Then
                dg.DefaultCellStyle.BackColor = Color.Gainsboro
            Else
                dg.DefaultCellStyle.BackColor = Color.White
            End If
        End If
        Dim f As Form
        f = tObj.FindForm()
        If tVal Then
            'tObj.foreColor = Color.Blue
            tObj.font = New Font(f.Font, FontStyle.Bold)
        Else
            'tObj.foreColor = Color.Black
            tObj.font = New Font(f.Font, FontStyle.Regular)
        End If
    End Sub
    Public Sub setGridReadOnly(ByRef dg As DataGridView, ByVal tVal As Boolean, ByVal ParamArray fields As String())
        For i As Integer = 0 To fields.GetUpperBound(0)
            If fields(0) = "0" Then
                For j As Integer = 0 To dg.Columns.Count - 1
                    dg.Columns(j).ReadOnly = tVal
                Next
                Exit Sub
            End If
            dg.Columns(fields(i)).ReadOnly = tVal
        Next
    End Sub
    Public Sub setDTReadOnly(ByRef dt As DataTable, ByVal tVal As Boolean, ByVal ParamArray fields As String())
        For i As Integer = 0 To fields.GetUpperBound(0)
            If fields(0) = "0" Then
                For j As Integer = 0 To dt.Columns.Count - 1
                    dt.Columns(j).ReadOnly = tVal
                Next
                Exit Sub
            End If
            dt.Columns(fields(i)).ReadOnly = tVal
        Next
    End Sub
    Public Sub setGridWidth(ByRef dg As DataGridView, ByVal ParamArray args As Single())
        For i As Integer = 0 To args.GetUpperBound(0)
            dg.Columns(i).Width = args(i)
        Next
    End Sub
    Public Sub setUpdateOnly(ByVal ParamArray dg As DataGridView())
        For i As Integer = 0 To dg.GetUpperBound(0)
            dg(i).AllowUserToAddRows = False
            dg(i).AllowUserToDeleteRows = False
        Next
    End Sub
    Public Sub setGridStyle(ByRef dg As DataGridView, ByVal ParamArray widths As Single())
        Dim j As Integer
        For i As Integer = 0 To dg.Columns.Count - 1
            If dg.Columns(i).CellType.Equals(GetType(DataGridViewCheckBoxCell)) Then
                dg.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ElseIf dg.Columns(i).ValueType Is GetType(Double) Or dg.Columns(i).ValueType Is GetType(Decimal) Then
                dg.Columns(i).DefaultCellStyle.Format = "#,##0.00"
                dg.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            ElseIf dg.Columns(i).ValueType.Name.Contains("Int") Then
                dg.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            ElseIf dg.Columns(i).ValueType Is GetType(DateTime) Then
                dg.Columns(i).DefaultCellStyle.Format = "dd MMM yyyy"
            End If
            If dg.Columns(i).Visible Then
                If j < widths.Length Then

                    dg.Columns(i).Width = widths(j)
                    j = j + 1
                End If
            End If
        Next
        If widths.Length > 0 Then
            If widths(0) = 0 Then

                dg.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
            End If
        End If

    End Sub
    Public Sub SetDataSource(ByRef dg As DataGridView, ByVal dt As DataTable, ByRef bs As BindingSource, Optional ByVal tCheckBoxColumns As String = "", Optional ByVal tButtonColumns As String = "")
        If Not dg.AutoGenerateColumns Then
            dg.Columns.Clear()
            Dim cols As New List(Of DataGridViewColumn)
            For Each col As DataColumn In dt.Columns
                If tCheckBoxColumns.Contains("@" & col.ColumnName) Then
                    Dim chk As New DataGridViewCheckBoxColumn
                    chk.DataPropertyName = col.ColumnName
                    chk.Name = col.ColumnName
                    cols.Add(chk)
                ElseIf tButtonColumns.Contains("@" & col.ColumnName) Then
                    Dim btn As New DataGridViewButtonColumn
                    btn.DataPropertyName = col.ColumnName
                    btn.Name = col.ColumnName
                    cols.Add(btn)
                Else
                    Dim txt As New DataGridViewTextBoxColumn
                    txt.DataPropertyName = col.ColumnName
                    txt.Name = col.ColumnName
                    cols.Add(txt)
                End If
            Next
            dg.Columns.AddRange(cols.ToArray)
        End If
        bs.DataSource = dt
        dg.DataSource = bs
        ModControl.setGridStyle(dg, 0)
        ModControl.setGridAutoInvisible(dg)
    End Sub
    Private Sub ClearMe(ByVal tObj As Object)
        If TypeOf tObj Is NumericTextBox Then
            tObj.value = 0
        ElseIf TypeOf tObj Is ComboBox OrElse _
            TypeOf tObj Is TextBox Then
            tObj.Text = ""
        ElseIf TypeOf tObj Is DateTimePicker Then
            tObj.value = tObj.mindate
            If Year(tObj.value) < 1900 Then tObj.value = DateSerial(1900, 1, 1)
        ElseIf TypeOf tObj Is DataGridView Then
            Dim dg As DataGridView
            dg = tObj
            dg.Columns.Clear()
        ElseIf TypeOf tObj Is CheckBox Then
            tObj.checked = False
        End If
    End Sub
    

    Public Sub EnableControl(ByRef tControl As Control, ByVal tVal As Boolean)
        Dim xControl As Object
        LockMe(tControl, Not tVal)
        For Each xControl In tControl.Controls
            EnableControl(xControl, tVal)
        Next
    End Sub
    Public Sub BindToDataTable(ByRef tControl As Control)
        Dim xControl As Object
        For i As Integer = 0 To tControl.DataBindings.Count - 1
            tControl.DataBindings(i).WriteValue()
        Next
        For Each xControl In tControl.Controls
            BindToDataTable(xControl)
        Next
    End Sub
    

    Public Sub ClearControl(ByRef tControl As Control)
        Dim xControl As Object
        For Each xControl In tControl.Controls
            ClearControl(xControl)
            ClearMe(xControl)
        Next
    End Sub
    Public Sub ClearBinding(ByRef tControl As Control)
        Dim xControl As Object
        For Each xControl In tControl.Controls
            ClearBinding(xControl)
            ClearBindingMe(xControl)
        Next
    End Sub
    Public Sub ClearBindingMe(ByVal tObj As Object)
        If TypeOf tObj Is ComboBox Then
            Dim o As ComboBox = tObj
            o.DataBindings.Clear()
        ElseIf TypeOf tObj Is TextBox Then
            Dim o As TextBox = tObj
            o.DataBindings.Clear()
        ElseIf TypeOf tObj Is DateTimePicker Then
            Dim o As DateTimePicker = tObj
            o.DataBindings.Clear()
        ElseIf TypeOf tObj Is CheckBox Then
            Dim o As CheckBox = tObj
            o.DataBindings.Clear()
        ElseIf TypeOf tObj Is NumericTextBox Then
            Dim o As NumericTextBox = tObj
            o.DataBindings.Clear()
        End If
    End Sub
    Public Sub updateRowState(ByRef mdt As DataTable, ByVal tPos As Integer)
        Dim rowState As DataRowState

        rowState = mdt.Rows(tPos).RowState
        mdt.AcceptChanges()
        If rowState = DataRowState.Added Then
            mdt.Rows(tPos).SetAdded()
        Else
            mdt.Rows(tPos).SetModified()
        End If
    End Sub
    Public Sub SetBinding(ByRef tControl As Control, ByVal dt As DataTable)
        Dim xControl As Object
        For Each xControl In tControl.Controls
            SetBinding(xControl, dt)
            SetBindingMe(xControl, dt)
        Next
    End Sub
    Public Sub SetBindingMe(ByVal tObj As Object, ByVal dt As DataTable)
        If Not tObj.Name.StartsWith("db") Then Exit Sub
        tObj.DataBindings.Clear()
        If TypeOf tObj Is ComboBox Then
            Dim o As ComboBox = tObj
            o.DataBindings.Add("SelectedValue", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is NumericTextBox Then
            Dim o As NumericTextBox = tObj
            o.DataBindings.Add("value", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is TextBox Then
            Dim o As TextBox = tObj
            o.DataBindings.Add("text", dt, o.Name.Substring(2))

        ElseIf TypeOf tObj Is DateTimePicker Then
            Dim o As DateTimePicker = tObj
            o.DataBindings.Add("value", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is CheckBox Then
            Dim o As CheckBox = tObj
            o.DataBindings.Add("checked", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is NumericTextBox Then
            Dim o As NumericTextBox = tObj
            o.DataBindings.Add("value", dt, o.Name.Substring(2))
        End If
        If tObj.Name = "dbPengupdate" OrElse tObj.Name = "dbWaktuUpdate" Then
            LockMe(tObj, True)
            Exit Sub
        End If
        LockMe(tObj, dt.Columns(tObj.Name.Substring(2)).ReadOnly)
    End Sub

    Public Sub SetBinding(ByRef tControl As Control, ByVal dt As BindingSource)
        Dim xControl As Object
        For Each xControl In tControl.Controls
            SetBinding(xControl, dt)
            SetBindingMe(xControl, dt)
        Next
    End Sub
    Public Sub SetBindingMe(ByVal tObj As Object, ByVal dt As BindingSource)
        If Not tObj.Name.StartsWith("db") Then Exit Sub
        tObj.DataBindings.Clear()
        If TypeOf tObj Is ComboBox Then
            Dim o As ComboBox = tObj
            o.DataBindings.Add("SelectedValue", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is NumericTextBox Then
            Dim o As NumericTextBox = tObj
            o.DataBindings.Add("value", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is TextBox Then
            Dim o As TextBox = tObj
            o.DataBindings.Add("text", dt, o.Name.Substring(2))

        ElseIf TypeOf tObj Is DateTimePicker Then
            Dim o As DateTimePicker = tObj
            o.DataBindings.Add("value", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is CheckBox Then
            Dim o As CheckBox = tObj
            o.DataBindings.Add("checked", dt, o.Name.Substring(2))
        ElseIf TypeOf tObj Is NumericTextBox Then
            Dim o As NumericTextBox = tObj
            o.DataBindings.Add("value", dt, o.Name.Substring(2))
        End If
        If tObj.Name = "dbPengupdate" OrElse tObj.Name = "dbWaktuUpdate" Then
            LockMe(tObj, True)
            Exit Sub
        End If
    End Sub

End Module

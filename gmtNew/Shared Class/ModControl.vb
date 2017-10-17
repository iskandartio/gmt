Module ModControl
    Public Sub setGridVisible(ByRef dg As DataGridView, ByVal tVal As Boolean, ByVal ParamArray args As String())
        For i As Integer = 0 To args.GetUpperBound(0)
            dg.Columns(args(i)).Visible = tVal
        Next
    End Sub
    Public Sub setGridAutoInvisible(ByRef dg As DataGridView, ByVal ParamArray args As String())
        Dim j As Integer = 0
        For i As Integer = 0 To dg.ColumnCount - 1
            If Right(dg.Columns(i).Name, 2) = "ID" Then
                dg.Columns(i).Visible = False
            Else
                If j < args.Length Then
                    dg.Columns(i).HeaderText = args(j)
                    j = j + 1
                End If
            End If
        Next
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
        If Not tObj.tabstop Then Exit Sub
        If TypeOf tObj Is ComboBox Then
            tObj.enabled = Not tVal
            If tVal Then
                tObj.BackColor = Color.White
            Else
                tObj.BackColor = Color.Gainsboro
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
                tObj.BackColor = Color.Gainsboro
            End If
        ElseIf TypeOf tObj Is Label Then
            Exit Sub
        ElseIf TypeOf tObj Is Button Then
            Exit Sub
        ElseIf TypeOf tObj Is CheckBox Then
            tObj.enabled = Not tVal
        ElseIf TypeOf tObj Is DataGridView Then
            Dim dg As DataGridView
            dg = tObj
            dg.ReadOnly = tVal
            dg.AllowUserToDeleteRows = Not tVal
            If tVal Then
                dg.DefaultCellStyle.BackColor = Color.Gainsboro
            Else
                dg.DefaultCellStyle.BackColor = Color.White
            End If

        End If
        Dim f As Form
        f = tObj.FindForm()
        'If tVal Then
        '    tObj.font = New Font(f.Font, FontStyle.Bold)
        'Else
        '    tObj.font = New Font(f.Font, FontStyle.Regular)
        'End If
    End Sub
    Public Sub setGridReadOnly(ByRef dg As DataGridView, ByVal ParamArray args As String())
        For i As Integer = 0 To args.GetUpperBound(0)
            dg.Columns(args(i)).ReadOnly = True
        Next
    End Sub
    Public Sub setGridWidth(ByRef dg As DataGridView, ByVal ParamArray args As Single())
        For i As Integer = 0 To args.GetUpperBound(0)
            dg.Columns(i).Width = args(i)
        Next
    End Sub
    Public Sub setGridStyle(ByRef dg As DataGridView, ByVal ParamArray args As Single())
        Dim j As Integer
        For i As Integer = 0 To dg.Columns.Count - 1
            If dg.Columns(i).ValueType Is GetType(Double) Then
                dg.Columns(i).DefaultCellStyle.Format = "#,##0.00"
                dg.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            ElseIf dg.Columns(i).ValueType Is GetType(Integer) Then
                dg.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            ElseIf dg.Columns(i).ValueType Is GetType(DateTime) Then
                dg.Columns(i).DefaultCellStyle.Format = "dd MMM yyyy"
            End If
            If dg.Columns(i).Visible Then
                If j < args.Length Then
                    dg.Columns(i).Width = args(j)
                    j = j + 1
                End If
            End If
        Next

    End Sub
    Private Sub ClearMe(ByVal tObj As Object)
        If TypeOf tObj Is ComboBox OrElse _
            TypeOf tObj Is TextBox Then
            tObj.Text = ""
        ElseIf TypeOf tObj Is DateTimePicker Then
            tObj.value = tObj.mindate
            If Year(tObj.value) < 1900 Then tObj.value = DateSerial(1900, 1, 1)
        ElseIf TypeOf tObj Is DataGridView Then
            Dim dg As DataGridView
            dg = tObj
            dg.Rows.Clear()
        ElseIf TypeOf tObj Is CheckBox Then
            tObj.checked = False
        End If
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

    Public Sub EnableControl(ByRef tControl As Control, ByVal tVal As Boolean)
        Dim xControl As Object
        For Each xControl In tControl.Controls
            EnableControl(xControl, tVal)
            LockMe(xControl, Not tVal)
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
End Module

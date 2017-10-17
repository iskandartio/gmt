Public Class rptRekapInputStock
    Private WithEvents prt As New Printing.PrintDocument
    Private WithEvents PrintPreviewDialog1 As New PrintPreviewDialog

    Private dtPrint As DataTable
    Dim mCurrent As Integer
    Dim mUrut As Integer
    Dim mConesCount As Integer
    Dim mBoxCount As Integer
    Dim mKgCount As Double
    Dim mLast As String

    Sub LoadMe(ByVal dt As DataTable, ByVal tTgl As String)
        dtPrint = dt
        mCurrent = 0
        prt.DefaultPageSettings.PaperSize = New Printing.PaperSize("a", prt.DefaultPageSettings.PaperSize.Width, prt.DefaultPageSettings.PaperSize.Height)
        dTgl.Text = tTgl
        If MsgBox("Langsung Print?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            PrintPreviewDialog1.Document = prt
            PrintPreviewDialog1.Width = Screen.PrimaryScreen.Bounds.Width
            PrintPreviewDialog1.Height = Screen.PrimaryScreen.Bounds.Height
            PrintPreviewDialog1.PrintPreviewControl.Zoom = 1.0
            PrintPreviewDialog1.ShowDialog()
        Else
            Dim a As New SQLtInputStock
            prt.Print()
        End If

    End Sub
    Private Sub prt_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prt.PrintPage
        Dim urut As Integer
        Dim f As New StringFormat
        Dim ConesCount As Integer
        Dim BoxCount As Integer
        Dim KgCount As Double
        Dim GrandConesCount As Integer
        Dim GrandBoxCount As Integer
        Dim GrandKgCount As Double
        Dim x As Single
        Dim l As Single
        Dim y As Single
        
        Dim Last As String
        
        Dim row As DataRow
        For Each c As Control In PageHeader.Controls
            DrawMe(c, dtPrint.Rows(0), e)
        Next
        Last = ""

        y = GroupHeader.Top
        urut = mUrut
        ConesCount = mConesCount
        BoxCount = mBoxCount
        KgCount = mKgCount
        Last = mLast
        For i As Integer = mCurrent To dtPrint.Rows.Count - 1

            row = dtPrint.Rows(i)
            If i = mCurrent Then
                For Each c As Control In GroupHeader.Controls
                    DrawMe(c, row, e, 0, y)
                Next
            End If

            If Last <> row("KodeBarang") & "!@#" & row("Lot") Then
                If urut > 1 Then
                    PrintFooter(e, urut, 0, y, ConesCount, BoxCount, KgCount)
                    y = y + 40
                    'y = y + GroupFooter.Top - Detail.Top + 30
                    'mCurrent = i
                    'e.HasMorePages = True
                    'Exit Sub
                End If
                urut = 1
                For Each c As Control In GroupHeader.Controls
                    DrawMe(c, row, e, 0, y)
                Next
                'y = y + Detail.Top - GroupHeader.Top - GroupHeader.Height + 20
                Last = row("KodeBarang") & "!@#" & row("Lot")
                x = 0
                ConesCount = 0
                BoxCount = 0
                KgCount = 0
            End If

            For Each c As Control In Detail.Controls
                'If urut < 6 Then
                '    GetPosition(c, f, l)
                '    e.Graphics.DrawString(c.Text, c.Font, New SolidBrush(c.ForeColor), l + x, c.Top + y - 20, f)
                'End If
                If c.Text = "No" Then
                    DrawMe(c, row, e, x, y, urut)
                Else
                    DrawMe(c, row, e, x, y)
                End If
            Next
            urut = urut + 1
            'x = x + 130
            'If x = 130 * 5 Then
            '    x = 0
            '    y = y + 20
            'End If
            y = y + 20
            ConesCount = ConesCount + row("nCones")
            BoxCount = BoxCount + row("n1")
            KgCount = KgCount + row("n2")
            GrandConesCount = GrandConesCount + row("nCones")
            GrandBoxCount = GrandBoxCount + row("n1")
            GrandKgCount = GrandKgCount + row("n2")
            If y > 1050 Then
                mCurrent = i + 1
                mUrut = urut
                mConesCount = ConesCount
                mBoxCount = BoxCount
                mKgCount = KgCount
                mLast = Last
                e.HasMorePages = True
                Exit Sub
            End If
        Next

        PrintFooter(e, urut, 0, y, ConesCount, BoxCount, KgCount)
        dTotal.Text = "GRAND TOTAL"
        PrintFooter(e, urut, 0, y + 50, GrandConesCount, GrandBoxCount, GrandKgCount)
        e.HasMorePages = False

    End Sub

    Public Sub DrawMe(ByVal c As Control, ByVal r As DataRow, ByVal e As System.Drawing.Printing.PrintPageEventArgs, Optional ByVal x As Single = 0, Optional ByVal y As Single = 0, Optional ByVal tStr As String = "")
        Dim t As String
        Dim f As New StringFormat
        Dim l As Single
        If tStr <> "" Then
            If c.Tag = "Header" Then
                t = tStr.ToString
            ElseIf c.Tag = "DateTime" Then
                t = convertToDate(r(tStr)).ToString
            ElseIf c.Tag = "Double" Then
                t = Format(r(tStr), "#,##0.00")
            ElseIf c.Tag = "Zerofill" Then
                If Len(r(tStr).ToString) < 4 Then
                    t = Format(r(tStr), "0000")
                Else
                    t = r(tStr)
                End If
            Else
                t = r(tStr).ToString
            End If
        Else
            If c.Tag = "Header" Then
                t = c.Text.ToString
            ElseIf c.Tag = "DateTime" Then
                t = convertToDate(r(c.Text)).ToString
            ElseIf c.Tag = "Double" Then
                t = Format(r(c.Text), "#,##0.00")
            ElseIf c.Tag = "Zerofill" Then
                If Len(r(c.Text).ToString) < 4 Then
                    t = Format(CInt(r(c.Text)), "0000")
                Else
                    t = r(c.Text)
                End If
            Else
                t = r(c.Text).ToString
            End If
        End If
        GetPosition(c, f, l)
        e.Graphics.DrawString(t, c.Font, New SolidBrush(c.ForeColor), l + x, c.Top + y, f)
    End Sub

    Public Sub GetPosition(ByRef c As Control, ByRef f As StringFormat, ByRef l As Single)
        If c.GetType Is GetType(Label) Then
            Dim o As Label = c
            If o.TextAlign = ContentAlignment.TopRight Then
                f.Alignment = StringAlignment.Far
                l = o.Right
            Else
                f.Alignment = StringAlignment.Near
                l = o.Left
            End If
        Else
            Dim o As TextBox = c
            If o.TextAlign = HorizontalAlignment.Right Then
                f.Alignment = StringAlignment.Far
                l = o.Right
            Else
                f.Alignment = StringAlignment.Near
                l = o.Left
            End If
        End If
    End Sub
    Private Sub PrintFooter(ByVal e As System.Drawing.Printing.PrintPageEventArgs, ByVal urut As Integer, ByVal xHeader As Single, ByVal y As Single, ByVal tConesCount As Integer, ByVal tBoxCount As Integer, ByVal tKgCount As Double)
        Dim i As Integer
        If urut < 5 Then i = urut - 1 Else i = 5
        'e.Graphics.DrawLine(Pens.Black, xHeader, yHeader, xHeader + i * 130, yHeader)
        'e.Graphics.DrawLine(Pens.Black, xHeader, yHeader + 20, xHeader + i * 130, yHeader + 20)
        'For j As Integer = 0 To i
        '    e.Graphics.DrawLine(Pens.Black, xHeader + j * 130, yHeader, xHeader + j * 130, yHeader + 20 * (2 + (urut \ 5)))
        'Next
        'For j As Integer = 1 To 2 + urut \ 5
        '    e.Graphics.DrawLine(Pens.Black, xHeader, yHeader + j * 20, xHeader + i * 130, yHeader + j * 20)
        'Next
        'y = yHeader + 20 * (2 + (urut \ 5))
        y = y + 20
        e.Graphics.DrawLine(Pens.Black, 50, y, 700, y)
        y = y + 5
        For Each c As Control In GroupFooter.Controls
            If c.Text = "ConesCount" Then
                DrawMe(c, dtPrint.Rows(0), e, 0, y - 15, tConesCount)
            ElseIf c.Text = "BoxCount" Then
                DrawMe(c, dtPrint.Rows(0), e, 0, y - 15, tBoxCount)
            ElseIf c.Text = "KgCount" Then
                DrawMe(c, dtPrint.Rows(0), e, 0, y - 15, tKgCount)
            Else
                DrawMe(c, dtPrint.Rows(0), e, 0, y - 15)
            End If

        Next

    End Sub

End Class
Public Class rptMutasi
    Private WithEvents prt As New Printing.PrintDocument
    Private WithEvents PrintPreviewDialog1 As New PrintPreviewDialog

    Private dtPrint As DataTable
    Dim mCurrent As Integer
    Dim inBox As Integer
    Dim outBox As Integer
    Dim akhirBox As Integer
    Dim inKg As Double
    Dim outKg As Double
    Dim akhirKg As Double
    Dim urut As Integer
    Dim LastKodeBarang As String

    Sub LoadMe(ByVal dt As DataTable, ByVal tDate As DateTime)
        dtPrint = dt
        mCurrent = 0
        dTgl.Text = Format(tDate, "dd-MM-yyyy")
        prt.DefaultPageSettings.PaperSize = New Printing.PaperSize("a", prt.DefaultPageSettings.PaperSize.Width, prt.DefaultPageSettings.PaperSize.Height)

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

    Private Sub prt_BeginPrint(sender As Object, e As Printing.PrintEventArgs) Handles prt.BeginPrint
        If Not prt.PrintController.IsPreview Then
            Dim pdi As PrintDialog = New PrintDialog
            pdi.ShowDialog()
        End If

    End Sub
    Private Sub prt_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prt.PrintPage
        Dim f As New StringFormat
        Dim ConesCount As Integer
        Dim BoxCount As Integer
        Dim KgCount As Double
        Dim x As Single
        Dim l As Single
        Dim y As Single
        Dim yPageHeader As Single
        Dim yHeader As Single
        Dim xHeader As Single
        Dim row As DataRow


        For Each c As Control In PageHeader.Controls
            DrawMe(c, dtPrint.Rows(0), e)


        Next





        xHeader = txtNo.Left - 10
        y = LineBottom.Y1 + 15

        yPageHeader = GroupHeader.Top
        For i As Integer = mCurrent To dtPrint.Rows.Count - 1

            row = dtPrint.Rows(i)
            If LastKodeBarang <> row("KodeBarang") Then
                If urut > 1 Then
                    PrintFooter(e, y - 20)
                    y = y + 25
                End If
                urut = 1
                For Each c As Control In GroupHeader.Controls
                    DrawMe(c, row, e, 0, y)
                Next
                DrawHorizontal(e, y + 35)

                y = y + 25
                LastKodeBarang = row("KodeBarang")
                yHeader = y - 10
                x = 0
                ConesCount = 0
                BoxCount = 0
                KgCount = 0
            End If
            inBox += row("inBox")
            inKg += row("inKg")
            outBox += row("outBox")
            outKg += row("outKg")
            akhirBox += row("akhirBox")
            akhirKg += row("akhirKg")
            For Each c As Control In Detail.Controls
                If c.Text = "No" Then
                    DrawMe(c, row, e, x, y, urut)
                Else
                    DrawMe(c, row, e, x, y)
                End If
            Next
            DrawHorizontal(e, y + 30)
            urut = urut + 1
            y = y + 20
            If y > 1000 Then
                mCurrent = i + 1
                If mCurrent < dtPrint.Rows.Count - 1 Then
                    row = dtPrint.Rows(mCurrent)
                    If LastKodeBarang <> row("KodeBarang") Then
                        PrintFooter(e, y - 20)
                        urut = 0
                        e.Graphics.DrawLine(Pens.Black, LineNo.X1, LineBottom.Y1, LineNo.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineJenis.X1, LineBottom.Y1, LineJenis.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineInBox.X1, LineBottom.Y1, LineInBox.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineInKg.X1, LineBottom.Y1, LineInKg.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineOutBox.X1, LineBottom.Y1, LineOutBox.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineOutKg.X1, LineBottom.Y1, LineOutKg.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineAkhirBox.X1, LineBottom.Y1, LineAkhirBox.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineAkhirKg.X1, LineBottom.Y1, LineAkhirKg.X1, y + 35)
                        e.Graphics.DrawLine(Pens.Black, LineLast.X1, LineBottom.Y1, LineLast.X1, y + 35)
                    Else
                        e.Graphics.DrawLine(Pens.Black, LineNo.X1, LineBottom.Y1, LineNo.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineJenis.X1, LineBottom.Y1, LineJenis.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineInBox.X1, LineBottom.Y1, LineInBox.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineInKg.X1, LineBottom.Y1, LineInKg.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineOutBox.X1, LineBottom.Y1, LineOutBox.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineOutKg.X1, LineBottom.Y1, LineOutKg.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineAkhirBox.X1, LineBottom.Y1, LineAkhirBox.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineAkhirKg.X1, LineBottom.Y1, LineAkhirKg.X1, y + 12)
                        e.Graphics.DrawLine(Pens.Black, LineLast.X1, LineBottom.Y1, LineLast.X1, y + 12)
                    End If
                    
                    e.HasMorePages = True
                    Exit Sub
                End If
            End If
        Next
        PrintFooter(e, y - 20)
        e.Graphics.DrawLine(Pens.Black, LineNo.X1, LineBottom.Y1, LineNo.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineJenis.X1, LineBottom.Y1, LineJenis.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineInBox.X1, LineBottom.Y1, LineInBox.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineInKg.X1, LineBottom.Y1, LineInKg.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineOutBox.X1, LineBottom.Y1, LineOutBox.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineOutKg.X1, LineBottom.Y1, LineOutKg.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineAkhirBox.X1, LineBottom.Y1, LineAkhirBox.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineAkhirKg.X1, LineBottom.Y1, LineAkhirKg.X1, y + 35)
        e.Graphics.DrawLine(Pens.Black, LineLast.X1, LineBottom.Y1, LineLast.X1, y + 35)
        e.HasMorePages = False

    End Sub
    Public Sub DrawHorizontal(ByVal e As System.Drawing.Printing.PrintPageEventArgs, ByVal y As Long)
        e.Graphics.DrawLine(Pens.Black, LineNo.X1, y, LineLast.X1, y)

    End Sub
    Public Sub DrawVertical(ByVal e As System.Drawing.Printing.PrintPageEventArgs, ByVal x As Long, ByVal y1 As Long, ByVal y2 As Long)
        e.Graphics.DrawLine(Pens.Black, x, y1, x, y2)

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
            ElseIf c.Tag = "Integer" Then
                t = Format(r(tStr), "#,##0")
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
            ElseIf c.Tag = "Integer" Then
                t = Format(r(c.Text), "#,##0")
            ElseIf c.Tag = "Zerofill" Then
                If Len(r(c.Text).ToString) < 4 Then
                    t = Format(CInt(r(c.Text)), "0000")
                Else
                    t = r(c.Text)
                End If
            ElseIf c.GetType.Name = "ShapeContainer" Then
                Dim sContainer As New PowerPacks.ShapeContainer
                sContainer = c
                For Each l1 As PowerPacks.LineShape In sContainer.Shapes
                    e.Graphics.DrawLine(New Pen(l1.BorderColor), l1.X1, l1.Y1 + 20, l1.X2, l1.Y2 + 20)
                Next
            Else
                t = r(c.Text).ToString
            End If
        End If
        GetPosition(c, f, l)
        e.Graphics.DrawString(t, c.Font, New SolidBrush(c.ForeColor), l + x, c.Top + y, f)
    End Sub

    Public Sub GetPosition(ByRef c As Control, ByRef f As StringFormat, ByRef l As Single)
        If c.GetType.Name = "ShapeContainer" Then
            Exit Sub
        End If
        If c.GetType Is GetType(Label) Then
            Dim o As Label = c
            If o.TextAlign = ContentAlignment.TopRight Then
                f.Alignment = StringAlignment.Far
                l = o.Right
            ElseIf o.TextAlign = ContentAlignment.TopCenter Then
                f.Alignment = StringAlignment.Near
                Dim a As Size = TextRenderer.MeasureText(c.Text, c.Font)
                l = o.Left + (c.Width - a.Width) / 2
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
    Private Sub PrintFooter(ByVal e As System.Drawing.Printing.PrintPageEventArgs, ByVal y As Single)
        toutBox.Text = Format(outBox, "#,##0")
        tinBox.Text = Format(inBox, "#,##0")
        takhirBox.Text = Format(akhirBox, "#,##0")
        toutKg.Text = Format(outKg, "#,##0.00")
        tinKg.Text = Format(inKg, "#,##0.00")
        takhirKg.Text = Format(akhirKg, "#,##0.00")
        urut = 1
        For Each c As Control In GroupFooter.Controls


            DrawMe(c, dtPrint.Rows(0), e, 0, y)
        Next
        DrawHorizontal(e, y + 55)
        y = y + 25
        inBox = 0
        inKg = 0
        outBox = 0
        outKg = 0
        akhirBox = 0
        akhirKg = 0
    End Sub



End Class
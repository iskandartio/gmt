Public Class frmRInputStock
    Dim a As New SQLtInputStock
    Private Sub frmRInputStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pTipe = "DTY"
        dtTgl.Value = Now
        dtTglAkhir.Value = Now
    End Sub
    Private Sub btnTipe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTipe.Click
        If lblTipe.Text = "DTY" Then
            lblTipe.Text = "PE"
            pTipe = "PE"
        Else
            lblTipe.Text = "DTY"
            pTipe = "DTY"
        End If

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim dtPrint As New DataTable
        dtPrint = a.qRekap(dtTgl.Value, dtTglAkhir.Value)
        If dtPrint.Rows.Count = 0 Then
            MsgBox("No Data")
            Exit Sub
        End If
        Dim b As New rptRekapInputStock
        b.LoadMe(dtPrint, Format(dtTgl.Value, "dd-MMM-yyyy") & " - " & Format(dtTglAkhir.Value, "dd-MMM-yyyy"))

    End Sub

    Private Sub btnSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSummary.Click
        Dim dtPrint As New DataTable
        dtPrint = a.qSummary(dtTgl.Value, dtTglAkhir.Value)
        If dtPrint.Rows.Count = 0 Then
            MsgBox("No Data")
            Exit Sub
        End If
        Dim b As New rptSummaryInputStock
        b.LoadMe(dtPrint, Format(dtTgl.Value, "dd-MMM-yyyy") & " - " & Format(dtTglAkhir.Value, "dd-MMM-yyyy"))
    End Sub

    Private Sub btnExportToExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportToExcel.Click
        On Error GoTo err
        Dim exc As New Object
        Dim wb As New Object
        Dim ws As New Object
        Dim dt As New DataTable
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        exc = CreateObject("Excel.Application")
        wb = exc.WorkBooks.Add
        ws = wb.Sheets
        ws(1).Name = "Rekap Input Stock"
        dt = a.qRekap(dtTgl.Value, dtTglAkhir.Value)

        ModFunction.PasteToExcel(ws(1), dt)
        exc.visible = True
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
err:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(exc)
    End Sub
End Class
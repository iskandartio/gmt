<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRInputStock
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblTipe = New System.Windows.Forms.Label
        Me.btnTipe = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.dtTgl = New ModDTPicker
        Me.dtTglAkhir = New ModDTPicker
        Me.btnSummary = New System.Windows.Forms.Button
        Me.btnExportToExcel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblTipe
        '
        Me.lblTipe.AutoSize = True
        Me.lblTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTipe.Location = New System.Drawing.Point(14, 14)
        Me.lblTipe.Name = "lblTipe"
        Me.lblTipe.Size = New System.Drawing.Size(50, 24)
        Me.lblTipe.TabIndex = 59
        Me.lblTipe.Text = "DTY"
        '
        'btnTipe
        '
        Me.btnTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTipe.Location = New System.Drawing.Point(75, 12)
        Me.btnTipe.Name = "btnTipe"
        Me.btnTipe.Size = New System.Drawing.Size(56, 28)
        Me.btnTipe.TabIndex = 58
        Me.btnTipe.Text = "Ubah"
        Me.btnTipe.UseVisualStyleBackColor = False
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(263, 36)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(82, 30)
        Me.btnPrint.TabIndex = 57
        Me.btnPrint.Text = "Print Rekap"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'dtTgl
        '
        Me.dtTgl.CustomFormat = "dd MMM yyyy"
        Me.dtTgl.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dtTgl.DisabledForeColor = System.Drawing.Color.Black
        Me.dtTgl.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTgl.Location = New System.Drawing.Point(11, 46)
        Me.dtTgl.Name = "dtTgl"
        Me.dtTgl.Size = New System.Drawing.Size(120, 20)
        Me.dtTgl.TabIndex = 56
        Me.dtTgl.Value = New Date(2011, 4, 22, 0, 0, 0, 0)
        '
        'dtTglAkhir
        '
        Me.dtTglAkhir.CustomFormat = "dd MMM yyyy"
        Me.dtTglAkhir.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dtTglAkhir.DisabledForeColor = System.Drawing.Color.Black
        Me.dtTglAkhir.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTglAkhir.Location = New System.Drawing.Point(137, 46)
        Me.dtTglAkhir.Name = "dtTglAkhir"
        Me.dtTglAkhir.Size = New System.Drawing.Size(120, 20)
        Me.dtTglAkhir.TabIndex = 60
        Me.dtTglAkhir.Value = New Date(2011, 4, 22, 0, 0, 0, 0)
        '
        'btnSummary
        '
        Me.btnSummary.Location = New System.Drawing.Point(351, 36)
        Me.btnSummary.Name = "btnSummary"
        Me.btnSummary.Size = New System.Drawing.Size(90, 30)
        Me.btnSummary.TabIndex = 61
        Me.btnSummary.Text = "Print Summary"
        Me.btnSummary.UseVisualStyleBackColor = True
        '
        'btnExportToExcel
        '
        Me.btnExportToExcel.Location = New System.Drawing.Point(447, 36)
        Me.btnExportToExcel.Name = "btnExportToExcel"
        Me.btnExportToExcel.Size = New System.Drawing.Size(105, 30)
        Me.btnExportToExcel.TabIndex = 125
        Me.btnExportToExcel.Text = "Export To Excel"
        Me.btnExportToExcel.UseVisualStyleBackColor = True
        '
        'frmRInputStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(564, 84)
        Me.Controls.Add(Me.btnExportToExcel)
        Me.Controls.Add(Me.btnSummary)
        Me.Controls.Add(Me.dtTglAkhir)
        Me.Controls.Add(Me.lblTipe)
        Me.Controls.Add(Me.btnTipe)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.dtTgl)
        Me.Name = "frmRInputStock"
        Me.Text = "frmRInputStock"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblTipe As System.Windows.Forms.Label
    Friend WithEvents btnTipe As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents dtTgl As ModDTPicker
    Friend WithEvents dtTglAkhir As ModDTPicker
    Friend WithEvents btnSummary As System.Windows.Forms.Button
    Friend WithEvents btnExportToExcel As System.Windows.Forms.Button
End Class

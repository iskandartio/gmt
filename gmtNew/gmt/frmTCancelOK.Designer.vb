<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTCancelOK
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
        Me.btnSearch = New System.Windows.Forms.Button
        Me.dtTgl = New ModDTPicker
        Me.dg = New modDataGridView
        Me.lblTipe = New System.Windows.Forms.Label
        Me.btnTipe = New System.Windows.Forms.Button
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(139, 47)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(77, 31)
        Me.btnSearch.TabIndex = 65
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'dtTgl
        '
        Me.dtTgl.CustomFormat = "dd MMM yyyy"
        Me.dtTgl.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dtTgl.DisabledForeColor = System.Drawing.Color.Black
        Me.dtTgl.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTgl.Location = New System.Drawing.Point(13, 58)
        Me.dtTgl.Name = "dtTgl"
        Me.dtTgl.Size = New System.Drawing.Size(120, 20)
        Me.dtTgl.TabIndex = 64
        Me.dtTgl.Value = New Date(2011, 4, 22, 0, 0, 0, 0)
        '
        'dg
        '
        Me.dg.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(13, 84)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(713, 371)
        Me.dg.TabIndex = 63
        '
        'lblTipe
        '
        Me.lblTipe.AutoSize = True
        Me.lblTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTipe.Location = New System.Drawing.Point(16, 15)
        Me.lblTipe.Name = "lblTipe"
        Me.lblTipe.Size = New System.Drawing.Size(50, 24)
        Me.lblTipe.TabIndex = 67
        Me.lblTipe.Text = "DTY"
        '
        'btnTipe
        '
        Me.btnTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTipe.Location = New System.Drawing.Point(77, 13)
        Me.btnTipe.Name = "btnTipe"
        Me.btnTipe.Size = New System.Drawing.Size(56, 28)
        Me.btnTipe.TabIndex = 66
        Me.btnTipe.Text = "Ubah"
        Me.btnTipe.UseVisualStyleBackColor = False
        '
        'frmTCancelOK
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 467)
        Me.Controls.Add(Me.lblTipe)
        Me.Controls.Add(Me.btnTipe)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.dtTgl)
        Me.Controls.Add(Me.dg)
        Me.Name = "frmTCancelOK"
        Me.Text = "frmTCancelOK"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents dtTgl As ModDTPicker
    Friend WithEvents dg As modDataGridView
    Friend WithEvents lblTipe As System.Windows.Forms.Label
    Friend WithEvents btnTipe As System.Windows.Forms.Button
End Class

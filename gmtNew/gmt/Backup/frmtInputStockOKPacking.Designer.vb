﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmtInputStockOKPacking
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
        Me.btnOK = New System.Windows.Forms.Button
        Me.lblTipe = New System.Windows.Forms.Label
        Me.btnTipe = New System.Windows.Forms.Button
        Me.btnSearch = New System.Windows.Forms.Button
        Me.dtTgl = New ModDTPicker
        Me.dg = New modDataGridView
        Me.dtTglGudang = New ModDTPicker
        Me.btnPilih = New System.Windows.Forms.Button
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(469, 48)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(64, 31)
        Me.btnOK.TabIndex = 65
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'lblTipe
        '
        Me.lblTipe.AutoSize = True
        Me.lblTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTipe.Location = New System.Drawing.Point(9, 14)
        Me.lblTipe.Name = "lblTipe"
        Me.lblTipe.Size = New System.Drawing.Size(50, 24)
        Me.lblTipe.TabIndex = 64
        Me.lblTipe.Text = "DTY"
        '
        'btnTipe
        '
        Me.btnTipe.BackColor = System.Drawing.SystemColors.Control
        Me.btnTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTipe.Location = New System.Drawing.Point(70, 12)
        Me.btnTipe.Name = "btnTipe"
        Me.btnTipe.Size = New System.Drawing.Size(70, 30)
        Me.btnTipe.TabIndex = 63
        Me.btnTipe.Text = "Ubah"
        Me.btnTipe.UseVisualStyleBackColor = True
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(132, 48)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(77, 31)
        Me.btnSearch.TabIndex = 62
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'dtTgl
        '
        Me.dtTgl.CustomFormat = "dd MMM yyyy"
        Me.dtTgl.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dtTgl.DisabledForeColor = System.Drawing.Color.Black
        Me.dtTgl.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTgl.Location = New System.Drawing.Point(6, 53)
        Me.dtTgl.Name = "dtTgl"
        Me.dtTgl.Size = New System.Drawing.Size(120, 20)
        Me.dtTgl.TabIndex = 61
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
        Me.dg.Location = New System.Drawing.Point(6, 85)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(698, 422)
        Me.dg.TabIndex = 60
        '
        'dtTglGudang
        '
        Me.dtTglGudang.CustomFormat = "dd MMM yyyy"
        Me.dtTglGudang.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dtTglGudang.DisabledForeColor = System.Drawing.Color.Black
        Me.dtTglGudang.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTglGudang.Location = New System.Drawing.Point(343, 53)
        Me.dtTglGudang.Name = "dtTglGudang"
        Me.dtTglGudang.Size = New System.Drawing.Size(120, 20)
        Me.dtTglGudang.TabIndex = 66
        Me.dtTglGudang.Value = New Date(2011, 4, 22, 0, 0, 0, 0)
        '
        'btnPilih
        '
        Me.btnPilih.Location = New System.Drawing.Point(215, 48)
        Me.btnPilih.Name = "btnPilih"
        Me.btnPilih.Size = New System.Drawing.Size(122, 31)
        Me.btnPilih.TabIndex = 67
        Me.btnPilih.Text = "Hapus Pilihan"
        Me.btnPilih.UseVisualStyleBackColor = True
        '
        'frmtInputStockOKPacking
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(716, 519)
        Me.Controls.Add(Me.btnPilih)
        Me.Controls.Add(Me.dtTglGudang)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lblTipe)
        Me.Controls.Add(Me.btnTipe)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.dtTgl)
        Me.Controls.Add(Me.dg)
        Me.Name = "frmtInputStockOKPacking"
        Me.Text = "frmtInputStockOKPacking"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents lblTipe As System.Windows.Forms.Label
    Friend WithEvents btnTipe As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents dtTgl As ModDTPicker
    Friend WithEvents dg As modDataGridView
    Friend WithEvents dtTglGudang As ModDTPicker
    Friend WithEvents btnPilih As System.Windows.Forms.Button
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMCustomer
    Inherits FormMain

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
        Me.btnUpdate = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFind = New StringTextBoxNoKeyPreview
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnUpdatePenerima = New System.Windows.Forms.Button
        Me.dgPenerima = New modDataGridView
        Me.dg = New modDataGridView
        CType(Me.dgPenerima, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(231, 17)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(79, 27)
        Me.btnUpdate.TabIndex = 2
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Find"
        '
        'txtFind
        '
        Me.txtFind.Location = New System.Drawing.Point(12, 24)
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(213, 20)
        Me.txtFind.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 391)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Penerima"
        '
        'btnUpdatePenerima
        '
        Me.btnUpdatePenerima.Location = New System.Drawing.Point(69, 377)
        Me.btnUpdatePenerima.Name = "btnUpdatePenerima"
        Me.btnUpdatePenerima.Size = New System.Drawing.Size(117, 27)
        Me.btnUpdatePenerima.TabIndex = 4
        Me.btnUpdatePenerima.Text = "Update Penerima"
        Me.btnUpdatePenerima.UseVisualStyleBackColor = True
        '
        'dgPenerima
        '
        Me.dgPenerima.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgPenerima.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgPenerima.IsOnPilih = False
        Me.dgPenerima.LastSearch = Nothing
        Me.dgPenerima.Location = New System.Drawing.Point(12, 407)
        Me.dgPenerima.Name = "dgPenerima"
        Me.dgPenerima.Size = New System.Drawing.Size(578, 120)
        Me.dgPenerima.TabIndex = 3
        '
        'dg
        '
        Me.dg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(12, 50)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(578, 321)
        Me.dg.TabIndex = 1
        '
        'frmMCustomer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(600, 539)
        Me.Controls.Add(Me.btnUpdatePenerima)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dgPenerima)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtFind)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.dg)
        Me.Name = "frmMCustomer"
        Me.Text = "frmMCustomer"
        CType(Me.dgPenerima, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dg As modDataGridView
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFind As StringTextBoxNoKeyPreview
    Friend WithEvents dgPenerima As modDataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnUpdatePenerima As System.Windows.Forms.Button
End Class

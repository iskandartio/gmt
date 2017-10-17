<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMLotMapping
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
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFind = New StringTextBoxNoKeyPreview
        Me.btnUpdate = New System.Windows.Forms.Button
        Me.dg = New modDataGridView
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Find"
        '
        'txtFind
        '
        Me.txtFind.Location = New System.Drawing.Point(12, 33)
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(213, 20)
        Me.txtFind.TabIndex = 5
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(231, 26)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(79, 27)
        Me.btnUpdate.TabIndex = 7
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'dg
        '
        Me.dg.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(12, 59)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(637, 394)
        Me.dg.TabIndex = 6
        '
        'frmMLotMapping
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(661, 465)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtFind)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.dg)
        Me.Name = "frmMLotMapping"
        Me.Text = "frmMLotMapping"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFind As StringTextBoxNoKeyPreview
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents dg As modDataGridView
End Class

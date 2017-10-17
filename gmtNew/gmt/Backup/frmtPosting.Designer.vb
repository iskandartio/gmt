<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmtPosting
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
        Me.btnPost = New System.Windows.Forms.Button
        Me.dtTgl = New ModDTPicker
        Me.SuspendLayout()
        '
        'btnPost
        '
        Me.btnPost.Location = New System.Drawing.Point(31, 43)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(82, 30)
        Me.btnPost.TabIndex = 4
        Me.btnPost.Text = "Post"
        Me.btnPost.UseVisualStyleBackColor = True
        '
        'dtTgl
        '
        Me.dtTgl.CustomFormat = "dd MMM yyyy"
        Me.dtTgl.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dtTgl.DisabledForeColor = System.Drawing.Color.Black
        Me.dtTgl.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTgl.Location = New System.Drawing.Point(31, 17)
        Me.dtTgl.Name = "dtTgl"
        Me.dtTgl.Size = New System.Drawing.Size(120, 20)
        Me.dtTgl.TabIndex = 3
        Me.dtTgl.Value = New Date(2011, 4, 22, 0, 0, 0, 0)
        '
        'frmtPosting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(191, 84)
        Me.Controls.Add(Me.btnPost)
        Me.Controls.Add(Me.dtTgl)
        Me.Name = "frmtPosting"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnPost As System.Windows.Forms.Button
    Friend WithEvents dtTgl As ModDTPicker
End Class

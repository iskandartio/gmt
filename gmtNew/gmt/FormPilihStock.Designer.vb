<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPilihStock
    Inherits FormPilih

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
        Me.grpFilter = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtNoWarna = New System.Windows.Forms.TextBox
        Me.btnClear = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFind = New System.Windows.Forms.Button
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.grpFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpFilter
        '
        Me.grpFilter.Controls.Add(Me.Label2)
        Me.grpFilter.Controls.Add(Me.txtNoWarna)
        Me.grpFilter.Controls.Add(Me.btnClear)
        Me.grpFilter.Controls.Add(Me.Label1)
        Me.grpFilter.Controls.Add(Me.btnFind)
        Me.grpFilter.Controls.Add(Me.txtCode)
        Me.grpFilter.Location = New System.Drawing.Point(5, 3)
        Me.grpFilter.Name = "grpFilter"
        Me.grpFilter.Size = New System.Drawing.Size(538, 58)
        Me.grpFilter.TabIndex = 26
        Me.grpFilter.TabStop = False
        Me.grpFilter.Text = "&Filter"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(96, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "No Warna"
        '
        'txtNoWarna
        '
        Me.txtNoWarna.Location = New System.Drawing.Point(95, 32)
        Me.txtNoWarna.Name = "txtNoWarna"
        Me.txtNoWarna.Size = New System.Drawing.Size(83, 20)
        Me.txtNoWarna.TabIndex = 7
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(422, 19)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(52, 29)
        Me.btnClear.TabIndex = 5
        Me.btnClear.Text = "&Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Kode Barang"
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(480, 19)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(52, 29)
        Me.btnFind.TabIndex = 6
        Me.btnFind.Text = "F&ind"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(6, 32)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(83, 20)
        Me.txtCode.TabIndex = 1
        '
        'FormPilihStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(560, 481)
        Me.Controls.Add(Me.grpFilter)
        Me.Name = "FormPilihStock"
        Me.Text = "FormPilihStock"
        Me.Controls.SetChildIndex(Me.btnChoose, 0)
        Me.Controls.SetChildIndex(Me.btnClose, 0)
        Me.Controls.SetChildIndex(Me.grpFilter, 0)
        Me.grpFilter.ResumeLayout(False)
        Me.grpFilter.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents grpFilter As System.Windows.Forms.GroupBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents txtNoWarna As System.Windows.Forms.TextBox
    Public WithEvents btnClear As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents btnFind As System.Windows.Forms.Button
    Public WithEvents txtCode As System.Windows.Forms.TextBox
End Class

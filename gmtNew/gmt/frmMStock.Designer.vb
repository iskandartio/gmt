<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMStock
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
        Me.dg = New modDataGridView
        Me.btnJenis = New System.Windows.Forms.Button
        Me.btnStock = New System.Windows.Forms.Button
        Me.FrameFilter = New System.Windows.Forms.Panel
        Me.fKode = New StringTextBox
        Me.fJenis = New StringTextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnUpdate = New System.Windows.Forms.Button
        Me.btnKodeBarang = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFind = New StringTextBoxNoKeyPreview
        Me.btnUbahNama = New System.Windows.Forms.Button
        Me.btnWarna = New System.Windows.Forms.Button
        Me.btnNoWarna = New System.Windows.Forms.Button
        Me.btnTube = New System.Windows.Forms.Button
        Me.btnGrade = New System.Windows.Forms.Button
        Me.fWarna = New StringTextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.fNoWarna = New StringTextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.fTube = New StringTextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.fGrade = New StringTextBox
        Me.Label7 = New System.Windows.Forms.Label
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FrameFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'dg
        '
        Me.dg.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(12, 127)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(676, 352)
        Me.dg.TabIndex = 2
        '
        'btnJenis
        '
        Me.btnJenis.Location = New System.Drawing.Point(108, 42)
        Me.btnJenis.Name = "btnJenis"
        Me.btnJenis.Size = New System.Drawing.Size(65, 33)
        Me.btnJenis.TabIndex = 19
        Me.btnJenis.Text = "Jenis"
        Me.btnJenis.UseVisualStyleBackColor = True
        '
        'btnStock
        '
        Me.btnStock.Location = New System.Drawing.Point(12, 42)
        Me.btnStock.Name = "btnStock"
        Me.btnStock.Size = New System.Drawing.Size(89, 33)
        Me.btnStock.TabIndex = 18
        Me.btnStock.Text = "Stock"
        Me.btnStock.UseVisualStyleBackColor = True
        '
        'FrameFilter
        '
        Me.FrameFilter.Controls.Add(Me.fGrade)
        Me.FrameFilter.Controls.Add(Me.Label7)
        Me.FrameFilter.Controls.Add(Me.fTube)
        Me.FrameFilter.Controls.Add(Me.Label6)
        Me.FrameFilter.Controls.Add(Me.fNoWarna)
        Me.FrameFilter.Controls.Add(Me.Label5)
        Me.FrameFilter.Controls.Add(Me.fWarna)
        Me.FrameFilter.Controls.Add(Me.Label4)
        Me.FrameFilter.Controls.Add(Me.fKode)
        Me.FrameFilter.Controls.Add(Me.fJenis)
        Me.FrameFilter.Controls.Add(Me.Label3)
        Me.FrameFilter.Controls.Add(Me.Label2)
        Me.FrameFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FrameFilter.Location = New System.Drawing.Point(120, 78)
        Me.FrameFilter.Name = "FrameFilter"
        Me.FrameFilter.Size = New System.Drawing.Size(579, 40)
        Me.FrameFilter.TabIndex = 20
        '
        'fKode
        '
        Me.fKode.AcceptsReturn = True
        Me.fKode.BackColor = System.Drawing.SystemColors.Window
        Me.fKode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fKode.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fKode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fKode.Location = New System.Drawing.Point(98, 15)
        Me.fKode.MaxLength = 0
        Me.fKode.Name = "fKode"
        Me.fKode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fKode.Size = New System.Drawing.Size(89, 22)
        Me.fKode.TabIndex = 3
        Me.fKode.Tag = "Kode"
        '
        'fJenis
        '
        Me.fJenis.AcceptsReturn = True
        Me.fJenis.BackColor = System.Drawing.SystemColors.Window
        Me.fJenis.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fJenis.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fJenis.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fJenis.Location = New System.Drawing.Point(3, 15)
        Me.fJenis.MaxLength = 0
        Me.fJenis.Name = "fJenis"
        Me.fJenis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fJenis.Size = New System.Drawing.Size(89, 22)
        Me.fJenis.TabIndex = 4
        Me.fJenis.Tag = "Jenis"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(95, -1)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(100, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Kode Barang"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(3, -1)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(65, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Jenis"
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(12, 3)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(89, 33)
        Me.btnUpdate.TabIndex = 21
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnKodeBarang
        '
        Me.btnKodeBarang.Location = New System.Drawing.Point(179, 42)
        Me.btnKodeBarang.Name = "btnKodeBarang"
        Me.btnKodeBarang.Size = New System.Drawing.Size(83, 33)
        Me.btnKodeBarang.TabIndex = 22
        Me.btnKodeBarang.Text = "Kode Barang"
        Me.btnKodeBarang.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 81)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Find"
        '
        'txtFind
        '
        Me.txtFind.Location = New System.Drawing.Point(12, 98)
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(102, 20)
        Me.txtFind.TabIndex = 23
        '
        'btnUbahNama
        '
        Me.btnUbahNama.Location = New System.Drawing.Point(108, 4)
        Me.btnUbahNama.Name = "btnUbahNama"
        Me.btnUbahNama.Size = New System.Drawing.Size(80, 32)
        Me.btnUbahNama.TabIndex = 25
        Me.btnUbahNama.Text = "Ubah Nama"
        Me.btnUbahNama.UseVisualStyleBackColor = True
        '
        'btnWarna
        '
        Me.btnWarna.Location = New System.Drawing.Point(268, 42)
        Me.btnWarna.Name = "btnWarna"
        Me.btnWarna.Size = New System.Drawing.Size(83, 33)
        Me.btnWarna.TabIndex = 26
        Me.btnWarna.Text = "Warna"
        Me.btnWarna.UseVisualStyleBackColor = True
        '
        'btnNoWarna
        '
        Me.btnNoWarna.Location = New System.Drawing.Point(357, 42)
        Me.btnNoWarna.Name = "btnNoWarna"
        Me.btnNoWarna.Size = New System.Drawing.Size(83, 33)
        Me.btnNoWarna.TabIndex = 27
        Me.btnNoWarna.Text = "No Warna"
        Me.btnNoWarna.UseVisualStyleBackColor = True
        '
        'btnTube
        '
        Me.btnTube.Location = New System.Drawing.Point(446, 42)
        Me.btnTube.Name = "btnTube"
        Me.btnTube.Size = New System.Drawing.Size(83, 33)
        Me.btnTube.TabIndex = 28
        Me.btnTube.Text = "Tube"
        Me.btnTube.UseVisualStyleBackColor = True
        '
        'btnGrade
        '
        Me.btnGrade.Location = New System.Drawing.Point(535, 42)
        Me.btnGrade.Name = "btnGrade"
        Me.btnGrade.Size = New System.Drawing.Size(83, 33)
        Me.btnGrade.TabIndex = 29
        Me.btnGrade.Text = "Grade"
        Me.btnGrade.UseVisualStyleBackColor = True
        '
        'fWarna
        '
        Me.fWarna.AcceptsReturn = True
        Me.fWarna.BackColor = System.Drawing.SystemColors.Window
        Me.fWarna.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fWarna.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fWarna.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fWarna.Location = New System.Drawing.Point(193, 15)
        Me.fWarna.MaxLength = 0
        Me.fWarna.Name = "fWarna"
        Me.fWarna.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fWarna.Size = New System.Drawing.Size(89, 22)
        Me.fWarna.TabIndex = 9
        Me.fWarna.Tag = "Kode"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(190, -1)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(100, 17)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Warna"
        '
        'fNoWarna
        '
        Me.fNoWarna.AcceptsReturn = True
        Me.fNoWarna.BackColor = System.Drawing.SystemColors.Window
        Me.fNoWarna.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fNoWarna.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fNoWarna.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fNoWarna.Location = New System.Drawing.Point(288, 15)
        Me.fNoWarna.MaxLength = 0
        Me.fNoWarna.Name = "fNoWarna"
        Me.fNoWarna.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fNoWarna.Size = New System.Drawing.Size(89, 22)
        Me.fNoWarna.TabIndex = 11
        Me.fNoWarna.Tag = "Kode"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(285, -1)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(100, 17)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "No Warna"
        '
        'fTube
        '
        Me.fTube.AcceptsReturn = True
        Me.fTube.BackColor = System.Drawing.SystemColors.Window
        Me.fTube.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fTube.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fTube.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fTube.Location = New System.Drawing.Point(383, 15)
        Me.fTube.MaxLength = 0
        Me.fTube.Name = "fTube"
        Me.fTube.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fTube.Size = New System.Drawing.Size(89, 22)
        Me.fTube.TabIndex = 13
        Me.fTube.Tag = "Kode"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(380, -1)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(100, 17)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Tube"
        '
        'fGrade
        '
        Me.fGrade.AcceptsReturn = True
        Me.fGrade.BackColor = System.Drawing.SystemColors.Window
        Me.fGrade.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fGrade.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fGrade.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fGrade.Location = New System.Drawing.Point(478, 15)
        Me.fGrade.MaxLength = 0
        Me.fGrade.Name = "fGrade"
        Me.fGrade.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fGrade.Size = New System.Drawing.Size(89, 22)
        Me.fGrade.TabIndex = 15
        Me.fGrade.Tag = "Kode"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(475, -1)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(100, 17)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Grade"
        '
        'frmMStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(700, 491)
        Me.Controls.Add(Me.btnGrade)
        Me.Controls.Add(Me.btnTube)
        Me.Controls.Add(Me.btnNoWarna)
        Me.Controls.Add(Me.btnWarna)
        Me.Controls.Add(Me.btnUbahNama)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFind)
        Me.Controls.Add(Me.btnKodeBarang)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.FrameFilter)
        Me.Controls.Add(Me.btnJenis)
        Me.Controls.Add(Me.btnStock)
        Me.Controls.Add(Me.dg)
        Me.Name = "frmMStock"
        Me.Text = "frmMStock"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FrameFilter.ResumeLayout(False)
        Me.FrameFilter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dg As modDataGridView
    Private WithEvents btnJenis As System.Windows.Forms.Button
    Private WithEvents btnStock As System.Windows.Forms.Button
    Private WithEvents FrameFilter As System.Windows.Forms.Panel
    Public WithEvents fKode As StringTextBox
    Public WithEvents fJenis As StringTextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents btnUpdate As System.Windows.Forms.Button
    Private WithEvents btnKodeBarang As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFind As StringTextBoxNoKeyPreview
    Friend WithEvents btnUbahNama As System.Windows.Forms.Button
    Private WithEvents btnWarna As System.Windows.Forms.Button
    Private WithEvents btnNoWarna As System.Windows.Forms.Button
    Private WithEvents btnTube As System.Windows.Forms.Button
    Private WithEvents btnGrade As System.Windows.Forms.Button
    Public WithEvents fGrade As StringTextBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents fTube As StringTextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents fNoWarna As StringTextBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents fWarna As StringTextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
End Class

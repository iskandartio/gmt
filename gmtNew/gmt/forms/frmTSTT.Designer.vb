<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTSTT
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
        Me.btnNext = New System.Windows.Forms.Button
        Me.btnPrev = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.fQuickFind = New StringTextBoxNoKeyPreview
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.dbCustomerID = New ModComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.dbTglSTT = New ModDTPicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dbNoSTT = New StringTextBox
        Me.dbWaktuUpdate = New ModDTPicker
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.dbPengupdate = New StringTextBoxNoKeyPreview
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.dgPelunasan = New modDataGridView
        Me.dgPembayaran = New modDataGridView
        Me.Label9 = New System.Windows.Forms.Label
        CType(Me.dgPelunasan, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgPembayaran, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(337, 12)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(69, 25)
        Me.btnNext.TabIndex = 83
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnPrev
        '
        Me.btnPrev.Location = New System.Drawing.Point(262, 12)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(69, 25)
        Me.btnPrev.TabIndex = 82
        Me.btnPrev.Text = "Prev"
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(17, 19)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 13)
        Me.Label7.TabIndex = 81
        Me.Label7.Text = "Quick Find"
        '
        'fQuickFind
        '
        Me.fQuickFind.AcceptsReturn = True
        Me.fQuickFind.BackColor = System.Drawing.SystemColors.Window
        Me.fQuickFind.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fQuickFind.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fQuickFind.Location = New System.Drawing.Point(125, 14)
        Me.fQuickFind.MaxLength = 0
        Me.fQuickFind.Name = "fQuickFind"
        Me.fQuickFind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fQuickFind.Size = New System.Drawing.Size(124, 20)
        Me.fQuickFind.TabIndex = 80
        Me.fQuickFind.Tag = "Jenis"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(262, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 79
        Me.Label3.Text = "Tgl KW"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(262, 49)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 13)
        Me.Label8.TabIndex = 93
        Me.Label8.Text = "Tanggal STT"
        '
        'dbCustomerID
        '
        Me.dbCustomerID.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbCustomerID.FormattingEnabled = True
        Me.dbCustomerID.Location = New System.Drawing.Point(128, 74)
        Me.dbCustomerID.Name = "dbCustomerID"
        Me.dbCustomerID.Size = New System.Drawing.Size(362, 21)
        Me.dbCustomerID.TabIndex = 92
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 77)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 91
        Me.Label4.Text = "Nama Customer"
        '
        'dbTglSTT
        '
        Me.dbTglSTT.Checked = False
        Me.dbTglSTT.CustomFormat = "dd MMM yyyy"
        Me.dbTglSTT.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglSTT.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglSTT.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglSTT.Location = New System.Drawing.Point(369, 45)
        Me.dbTglSTT.Name = "dbTglSTT"
        Me.dbTglSTT.Size = New System.Drawing.Size(118, 20)
        Me.dbTglSTT.TabIndex = 86
        Me.dbTglSTT.Value = New Date(2011, 1, 16, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 85
        Me.Label2.Text = "No STT"
        '
        'dbNoSTT
        '
        Me.dbNoSTT.AcceptsReturn = True
        Me.dbNoSTT.BackColor = System.Drawing.SystemColors.Window
        Me.dbNoSTT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNoSTT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNoSTT.Location = New System.Drawing.Point(126, 44)
        Me.dbNoSTT.MaxLength = 0
        Me.dbNoSTT.Name = "dbNoSTT"
        Me.dbNoSTT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNoSTT.Size = New System.Drawing.Size(124, 20)
        Me.dbNoSTT.TabIndex = 84
        Me.dbNoSTT.Tag = "Jenis"
        '
        'dbWaktuUpdate
        '
        Me.dbWaktuUpdate.Checked = False
        Me.dbWaktuUpdate.CustomFormat = "dd MMM yyyy HH:mm:ss"
        Me.dbWaktuUpdate.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbWaktuUpdate.DisabledForeColor = System.Drawing.Color.Black
        Me.dbWaktuUpdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbWaktuUpdate.Location = New System.Drawing.Point(606, 70)
        Me.dbWaktuUpdate.Name = "dbWaktuUpdate"
        Me.dbWaktuUpdate.Size = New System.Drawing.Size(160, 20)
        Me.dbWaktuUpdate.TabIndex = 97
        Me.dbWaktuUpdate.Value = New Date(2011, 1, 16, 0, 0, 0, 0)
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(500, 74)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 96
        Me.Label10.Text = "Waktu Update"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(500, 44)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(65, 13)
        Me.Label11.TabIndex = 95
        Me.Label11.Text = "Pengupdate"
        '
        'dbPengupdate
        '
        Me.dbPengupdate.AcceptsReturn = True
        Me.dbPengupdate.BackColor = System.Drawing.SystemColors.Window
        Me.dbPengupdate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbPengupdate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbPengupdate.Location = New System.Drawing.Point(607, 44)
        Me.dbPengupdate.MaxLength = 0
        Me.dbPengupdate.Name = "dbPengupdate"
        Me.dbPengupdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbPengupdate.Size = New System.Drawing.Size(160, 20)
        Me.dbPengupdate.TabIndex = 94
        Me.dbPengupdate.Tag = "Jenis"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(193, 101)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(83, 25)
        Me.btnNew.TabIndex = 108
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(104, 101)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(83, 25)
        Me.btnCancel.TabIndex = 107
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(15, 101)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(83, 25)
        Me.btnSave.TabIndex = 106
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(18, 347)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 13)
        Me.Label6.TabIndex = 105
        Me.Label6.Text = "Pelunasan"
        '
        'dgPelunasan
        '
        Me.dgPelunasan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgPelunasan.IsOnPilih = False
        Me.dgPelunasan.LastSearch = Nothing
        Me.dgPelunasan.Location = New System.Drawing.Point(19, 363)
        Me.dgPelunasan.Name = "dgPelunasan"
        Me.dgPelunasan.Size = New System.Drawing.Size(886, 186)
        Me.dgPelunasan.TabIndex = 104
        '
        'dgPembayaran
        '
        Me.dgPembayaran.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgPembayaran.IsOnPilih = False
        Me.dgPembayaran.LastSearch = Nothing
        Me.dgPembayaran.Location = New System.Drawing.Point(19, 181)
        Me.dgPembayaran.Name = "dgPembayaran"
        Me.dgPembayaran.Size = New System.Drawing.Size(886, 152)
        Me.dgPembayaran.TabIndex = 103
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(18, 165)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(66, 13)
        Me.Label9.TabIndex = 109
        Me.Label9.Text = "Pembayaran"
        '
        'frmTSTT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(917, 559)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dgPelunasan)
        Me.Controls.Add(Me.dgPembayaran)
        Me.Controls.Add(Me.dbWaktuUpdate)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.dbPengupdate)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dbCustomerID)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dbTglSTT)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dbNoSTT)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrev)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.fQuickFind)
        Me.Controls.Add(Me.Label3)
        Me.Name = "frmTSTT"
        Me.Text = "frmTSTT"
        CType(Me.dgPelunasan, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgPembayaran, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnPrev As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents fQuickFind As StringTextBoxNoKeyPreview
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dbCustomerID As ModComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dbTglSTT As ModDTPicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents dbNoSTT As StringTextBox
    Friend WithEvents dbWaktuUpdate As ModDTPicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents dbPengupdate As StringTextBoxNoKeyPreview
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dgPelunasan As modDataGridView
    Friend WithEvents dgPembayaran As modDataGridView
    Friend WithEvents Label9 As System.Windows.Forms.Label
End Class

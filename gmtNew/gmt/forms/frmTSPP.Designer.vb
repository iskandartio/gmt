<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTSPP
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.fQuickFind = New StringTextBoxNoKeyPreview
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnApprove = New System.Windows.Forms.Button
        Me.cmbNamaPenerima = New System.Windows.Forms.ComboBox
        Me.txtStatus = New StringTextBox
        Me.lblStatus = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.dbAlamatPenerima = New StringTextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.dg = New modDataGridView
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.dbPengupdate = New StringTextBoxNoKeyPreview
        Me.Label9 = New System.Windows.Forms.Label
        Me.dbKeteranganSPP = New StringTextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.dbWaktuPembayaran = New StringTextBox
        Me.dbMataUang = New ModComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.dbCustomerID = New ModComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dbNoSPP = New StringTextBox
        Me.dbWaktuUpdate = New ModDTPicker
        Me.dbTglKirim = New ModDTPicker
        Me.dbTglSPP = New ModDTPicker
        Me.Panel1.SuspendLayout()
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(331, 10)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(69, 25)
        Me.btnNext.TabIndex = 74
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnPrev
        '
        Me.btnPrev.Location = New System.Drawing.Point(256, 10)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(69, 25)
        Me.btnPrev.TabIndex = 73
        Me.btnPrev.Text = "Prev"
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 42
        Me.Label1.Text = "Quick Find"
        '
        'fQuickFind
        '
        Me.fQuickFind.AcceptsReturn = True
        Me.fQuickFind.BackColor = System.Drawing.SystemColors.Window
        Me.fQuickFind.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fQuickFind.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fQuickFind.Location = New System.Drawing.Point(119, 12)
        Me.fQuickFind.MaxLength = 0
        Me.fQuickFind.Name = "fQuickFind"
        Me.fQuickFind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fQuickFind.Size = New System.Drawing.Size(124, 20)
        Me.fQuickFind.TabIndex = 41
        Me.fQuickFind.Tag = "Jenis"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.btnApprove)
        Me.Panel1.Controls.Add(Me.dbWaktuUpdate)
        Me.Panel1.Controls.Add(Me.cmbNamaPenerima)
        Me.Panel1.Controls.Add(Me.txtStatus)
        Me.Panel1.Controls.Add(Me.lblStatus)
        Me.Panel1.Controls.Add(Me.dbTglKirim)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.dbAlamatPenerima)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.btnNew)
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Controls.Add(Me.dg)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.dbPengupdate)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.dbKeteranganSPP)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.dbWaktuPembayaran)
        Me.Panel1.Controls.Add(Me.dbMataUang)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.dbCustomerID)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.dbTglSPP)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.dbNoSPP)
        Me.Panel1.Location = New System.Drawing.Point(14, 38)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(902, 523)
        Me.Panel1.TabIndex = 86
        '
        'btnApprove
        '
        Me.btnApprove.Location = New System.Drawing.Point(185, 215)
        Me.btnApprove.Name = "btnApprove"
        Me.btnApprove.Size = New System.Drawing.Size(83, 25)
        Me.btnApprove.TabIndex = 113
        Me.btnApprove.Text = "Approve"
        Me.btnApprove.UseVisualStyleBackColor = True
        '
        'cmbNamaPenerima
        '
        Me.cmbNamaPenerima.FormattingEnabled = True
        Me.cmbNamaPenerima.Location = New System.Drawing.Point(586, 14)
        Me.cmbNamaPenerima.Name = "cmbNamaPenerima"
        Me.cmbNamaPenerima.Size = New System.Drawing.Size(161, 21)
        Me.cmbNamaPenerima.TabIndex = 111
        '
        'txtStatus
        '
        Me.txtStatus.AcceptsReturn = True
        Me.txtStatus.BackColor = System.Drawing.SystemColors.Window
        Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStatus.Location = New System.Drawing.Point(586, 163)
        Me.txtStatus.MaxLength = 0
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatus.Size = New System.Drawing.Size(123, 20)
        Me.txtStatus.TabIndex = 110
        Me.txtStatus.Tag = "Jenis"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(484, 168)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(37, 13)
        Me.lblStatus.TabIndex = 109
        Me.lblStatus.Text = "Status"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(243, 71)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(47, 13)
        Me.Label8.TabIndex = 107
        Me.Label8.Text = "Tgl Kirim"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(480, 46)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(86, 13)
        Me.Label13.TabIndex = 106
        Me.Label13.Text = "Alamat Penerima"
        '
        'dbAlamatPenerima
        '
        Me.dbAlamatPenerima.AcceptsReturn = True
        Me.dbAlamatPenerima.BackColor = System.Drawing.SystemColors.Window
        Me.dbAlamatPenerima.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbAlamatPenerima.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbAlamatPenerima.Location = New System.Drawing.Point(586, 41)
        Me.dbAlamatPenerima.MaxLength = 0
        Me.dbAlamatPenerima.Multiline = True
        Me.dbAlamatPenerima.Name = "dbAlamatPenerima"
        Me.dbAlamatPenerima.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbAlamatPenerima.Size = New System.Drawing.Size(245, 56)
        Me.dbAlamatPenerima.TabIndex = 105
        Me.dbAlamatPenerima.Tag = "Jenis"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(480, 17)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(82, 13)
        Me.Label12.TabIndex = 104
        Me.Label12.Text = "Nama Penerima"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(7, 215)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(83, 25)
        Me.btnNew.TabIndex = 103
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(96, 215)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(83, 25)
        Me.btnSave.TabIndex = 102
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dg
        '
        Me.dg.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(9, 246)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(884, 266)
        Me.dg.TabIndex = 101
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 193)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 100
        Me.Label10.Text = "Waktu Update"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 163)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(65, 13)
        Me.Label11.TabIndex = 99
        Me.Label11.Text = "Pengupdate"
        '
        'dbPengupdate
        '
        Me.dbPengupdate.AcceptsReturn = True
        Me.dbPengupdate.BackColor = System.Drawing.SystemColors.Window
        Me.dbPengupdate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbPengupdate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbPengupdate.Location = New System.Drawing.Point(113, 163)
        Me.dbPengupdate.MaxLength = 0
        Me.dbPengupdate.Name = "dbPengupdate"
        Me.dbPengupdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbPengupdate.Size = New System.Drawing.Size(160, 20)
        Me.dbPengupdate.TabIndex = 98
        Me.dbPengupdate.Tag = "Jenis"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 103)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 13)
        Me.Label9.TabIndex = 97
        Me.Label9.Text = "Keterangan SPP"
        '
        'dbKeteranganSPP
        '
        Me.dbKeteranganSPP.AcceptsReturn = True
        Me.dbKeteranganSPP.BackColor = System.Drawing.SystemColors.Window
        Me.dbKeteranganSPP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbKeteranganSPP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbKeteranganSPP.Location = New System.Drawing.Point(113, 98)
        Me.dbKeteranganSPP.MaxLength = 0
        Me.dbKeteranganSPP.Multiline = True
        Me.dbKeteranganSPP.Name = "dbKeteranganSPP"
        Me.dbKeteranganSPP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbKeteranganSPP.Size = New System.Drawing.Size(281, 56)
        Me.dbKeteranganSPP.TabIndex = 96
        Me.dbKeteranganSPP.Tag = "Jenis"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(480, 106)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 13)
        Me.Label7.TabIndex = 95
        Me.Label7.Text = "Waktu Pembayaran"
        '
        'dbWaktuPembayaran
        '
        Me.dbWaktuPembayaran.AcceptsReturn = True
        Me.dbWaktuPembayaran.BackColor = System.Drawing.SystemColors.Window
        Me.dbWaktuPembayaran.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbWaktuPembayaran.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbWaktuPembayaran.Location = New System.Drawing.Point(586, 103)
        Me.dbWaktuPembayaran.MaxLength = 0
        Me.dbWaktuPembayaran.Name = "dbWaktuPembayaran"
        Me.dbWaktuPembayaran.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbWaktuPembayaran.Size = New System.Drawing.Size(123, 20)
        Me.dbWaktuPembayaran.TabIndex = 94
        Me.dbWaktuPembayaran.Tag = "Jenis"
        '
        'dbMataUang
        '
        Me.dbMataUang.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbMataUang.FormattingEnabled = True
        Me.dbMataUang.Location = New System.Drawing.Point(353, 42)
        Me.dbMataUang.Name = "dbMataUang"
        Me.dbMataUang.Size = New System.Drawing.Size(121, 21)
        Me.dbMataUang.TabIndex = 93
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(247, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 92
        Me.Label5.Text = "Mata Uang"
        '
        'dbCustomerID
        '
        Me.dbCustomerID.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbCustomerID.FormattingEnabled = True
        Me.dbCustomerID.Location = New System.Drawing.Point(112, 14)
        Me.dbCustomerID.Name = "dbCustomerID"
        Me.dbCustomerID.Size = New System.Drawing.Size(362, 21)
        Me.dbCustomerID.TabIndex = 91
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 90
        Me.Label4.Text = "Nama Customer"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 88
        Me.Label3.Text = "Tgl SPP"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 87
        Me.Label2.Text = "No SPP"
        '
        'dbNoSPP
        '
        Me.dbNoSPP.AcceptsReturn = True
        Me.dbNoSPP.BackColor = System.Drawing.SystemColors.Window
        Me.dbNoSPP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNoSPP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNoSPP.Location = New System.Drawing.Point(112, 42)
        Me.dbNoSPP.MaxLength = 0
        Me.dbNoSPP.Name = "dbNoSPP"
        Me.dbNoSPP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNoSPP.Size = New System.Drawing.Size(124, 20)
        Me.dbNoSPP.TabIndex = 86
        Me.dbNoSPP.Tag = "Jenis"
        '
        'dbWaktuUpdate
        '
        Me.dbWaktuUpdate.Checked = False
        Me.dbWaktuUpdate.CustomFormat = "dd MMM yyyy HH:mm:ss"
        Me.dbWaktuUpdate.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbWaktuUpdate.DisabledForeColor = System.Drawing.Color.Black
        Me.dbWaktuUpdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbWaktuUpdate.Location = New System.Drawing.Point(112, 189)
        Me.dbWaktuUpdate.Name = "dbWaktuUpdate"
        Me.dbWaktuUpdate.Size = New System.Drawing.Size(160, 20)
        Me.dbWaktuUpdate.TabIndex = 112
        '
        'dbTglKirim
        '
        Me.dbTglKirim.Checked = False
        Me.dbTglKirim.CustomFormat = "dd MMM yyyy"
        Me.dbTglKirim.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglKirim.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglKirim.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglKirim.Location = New System.Drawing.Point(353, 69)
        Me.dbTglKirim.Name = "dbTglKirim"
        Me.dbTglKirim.Size = New System.Drawing.Size(118, 20)
        Me.dbTglKirim.TabIndex = 108
        '
        'dbTglSPP
        '
        Me.dbTglSPP.Checked = False
        Me.dbTglSPP.CustomFormat = "dd MMM yyyy"
        Me.dbTglSPP.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglSPP.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglSPP.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglSPP.Location = New System.Drawing.Point(113, 68)
        Me.dbTglSPP.Name = "dbTglSPP"
        Me.dbTglSPP.Size = New System.Drawing.Size(118, 20)
        Me.dbTglSPP.TabIndex = 89
        '
        'frmTSPP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(919, 564)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrev)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.fQuickFind)
        Me.Name = "frmTSPP"
        Me.Text = "frmTSPP"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnPrev As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents fQuickFind As StringTextBoxNoKeyPreview
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnApprove As System.Windows.Forms.Button
    Friend WithEvents dbWaktuUpdate As ModDTPicker
    Friend WithEvents cmbNamaPenerima As System.Windows.Forms.ComboBox
    Public WithEvents txtStatus As StringTextBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents dbTglKirim As ModDTPicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents dbAlamatPenerima As StringTextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dg As modDataGridView
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents dbPengupdate As StringTextBoxNoKeyPreview
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents dbKeteranganSPP As StringTextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents dbWaktuPembayaran As StringTextBox
    Friend WithEvents dbMataUang As ModComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dbCustomerID As ModComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dbTglSPP As ModDTPicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents dbNoSPP As StringTextBox

End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTSC
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
        Me.fQuickFind = New StringTextBoxNoKeyPreview
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dbNoSC = New StringTextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dbTglSC = New ModDTPicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.dbCustomerID = New ModComboBox
        Me.dbMataUang = New ModComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.dbLamaKontrak = New StringTextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.dbWaktuPembayaran = New StringTextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.dbNilaiKontrak = New NumericTextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.dbKeterangan = New StringTextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.dbPengupdate = New StringTextBoxNoKeyPreview
        Me.chkDP = New System.Windows.Forms.CheckBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.dbNamaMarketing = New StringTextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.dbNamaCustomerSC = New StringTextBox
        Me.dg = New modDataGridView
        Me.btnSave = New System.Windows.Forms.Button
        Me.dbWaktuUpdate = New StringTextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrev = New System.Windows.Forms.Button
        Me.btnNext = New System.Windows.Forms.Button
        Me.txtStatus = New StringTextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.btnApprove = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.fQuickFind.TabIndex = 5
        Me.fQuickFind.Tag = "Jenis"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Quick Find"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "No SC"
        '
        'dbNoSC
        '
        Me.dbNoSC.AcceptsReturn = True
        Me.dbNoSC.BackColor = System.Drawing.SystemColors.Window
        Me.dbNoSC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNoSC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNoSC.Location = New System.Drawing.Point(118, 69)
        Me.dbNoSC.MaxLength = 0
        Me.dbNoSC.Name = "dbNoSC"
        Me.dbNoSC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNoSC.Size = New System.Drawing.Size(124, 20)
        Me.dbNoSC.TabIndex = 7
        Me.dbNoSC.Tag = "Jenis"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(251, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Tgl SC"
        '
        'dbTglSC
        '
        Me.dbTglSC.Checked = False
        Me.dbTglSC.CustomFormat = "dd MMM yyyy"
        Me.dbTglSC.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglSC.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglSC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglSC.Location = New System.Drawing.Point(361, 70)
        Me.dbTglSC.Name = "dbTglSC"
        Me.dbTglSC.Size = New System.Drawing.Size(118, 20)
        Me.dbTglSC.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Nama Customer"
        '
        'dbCustomerID
        '
        Me.dbCustomerID.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbCustomerID.FormattingEnabled = True
        Me.dbCustomerID.Location = New System.Drawing.Point(118, 41)
        Me.dbCustomerID.Name = "dbCustomerID"
        Me.dbCustomerID.Size = New System.Drawing.Size(597, 21)
        Me.dbCustomerID.TabIndex = 13
        '
        'dbMataUang
        '
        Me.dbMataUang.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbMataUang.FormattingEnabled = True
        Me.dbMataUang.Location = New System.Drawing.Point(359, 97)
        Me.dbMataUang.Name = "dbMataUang"
        Me.dbMataUang.Size = New System.Drawing.Size(121, 21)
        Me.dbMataUang.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(250, 100)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Mata Uang"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(485, 70)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 13)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Lama Kontrak"
        '
        'dbLamaKontrak
        '
        Me.dbLamaKontrak.AcceptsReturn = True
        Me.dbLamaKontrak.BackColor = System.Drawing.SystemColors.Window
        Me.dbLamaKontrak.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbLamaKontrak.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbLamaKontrak.Location = New System.Drawing.Point(592, 69)
        Me.dbLamaKontrak.MaxLength = 0
        Me.dbLamaKontrak.Name = "dbLamaKontrak"
        Me.dbLamaKontrak.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbLamaKontrak.Size = New System.Drawing.Size(123, 20)
        Me.dbLamaKontrak.TabIndex = 16
        Me.dbLamaKontrak.Tag = "Jenis"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(484, 98)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 13)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Waktu Pembayaran"
        '
        'dbWaktuPembayaran
        '
        Me.dbWaktuPembayaran.AcceptsReturn = True
        Me.dbWaktuPembayaran.BackColor = System.Drawing.SystemColors.Window
        Me.dbWaktuPembayaran.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbWaktuPembayaran.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbWaktuPembayaran.Location = New System.Drawing.Point(591, 97)
        Me.dbWaktuPembayaran.MaxLength = 0
        Me.dbWaktuPembayaran.Name = "dbWaktuPembayaran"
        Me.dbWaktuPembayaran.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbWaktuPembayaran.Size = New System.Drawing.Size(123, 20)
        Me.dbWaktuPembayaran.TabIndex = 18
        Me.dbWaktuPembayaran.Tag = "Jenis"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(11, 102)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(67, 13)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Nilai Kontrak"
        '
        'dbNilaiKontrak
        '
        Me.dbNilaiKontrak.AcceptsReturn = True
        Me.dbNilaiKontrak.BackColor = System.Drawing.SystemColors.Window
        Me.dbNilaiKontrak.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNilaiKontrak.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNilaiKontrak.Location = New System.Drawing.Point(119, 97)
        Me.dbNilaiKontrak.MaxLength = 0
        Me.dbNilaiKontrak.Name = "dbNilaiKontrak"
        Me.dbNilaiKontrak.NumberFormat = "#,##0.00"
        Me.dbNilaiKontrak.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNilaiKontrak.Size = New System.Drawing.Size(123, 20)
        Me.dbNilaiKontrak.TabIndex = 20
        Me.dbNilaiKontrak.Tag = "Jenis"
        Me.dbNilaiKontrak.Text = "0.00"
        Me.dbNilaiKontrak.Value = New Decimal(New Integer() {0, 0, 0, 131072})
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(11, 130)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(62, 13)
        Me.Label9.TabIndex = 23
        Me.Label9.Text = "Keterangan"
        '
        'dbKeterangan
        '
        Me.dbKeterangan.AcceptsReturn = True
        Me.dbKeterangan.BackColor = System.Drawing.SystemColors.Window
        Me.dbKeterangan.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbKeterangan.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbKeterangan.Location = New System.Drawing.Point(119, 125)
        Me.dbKeterangan.MaxLength = 0
        Me.dbKeterangan.Multiline = True
        Me.dbKeterangan.Name = "dbKeterangan"
        Me.dbKeterangan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbKeterangan.Size = New System.Drawing.Size(596, 56)
        Me.dbKeterangan.TabIndex = 22
        Me.dbKeterangan.Tag = "Jenis"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(249, 191)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "Waktu Update"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(11, 192)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(65, 13)
        Me.Label11.TabIndex = 26
        Me.Label11.Text = "Pengupdate"
        '
        'dbPengupdate
        '
        Me.dbPengupdate.AcceptsReturn = True
        Me.dbPengupdate.BackColor = System.Drawing.SystemColors.Window
        Me.dbPengupdate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbPengupdate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbPengupdate.Location = New System.Drawing.Point(119, 187)
        Me.dbPengupdate.MaxLength = 0
        Me.dbPengupdate.Name = "dbPengupdate"
        Me.dbPengupdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbPengupdate.Size = New System.Drawing.Size(123, 20)
        Me.dbPengupdate.TabIndex = 25
        Me.dbPengupdate.Tag = "Jenis"
        '
        'chkDP
        '
        Me.chkDP.AutoSize = True
        Me.chkDP.Location = New System.Drawing.Point(663, 191)
        Me.chkDP.Name = "chkDP"
        Me.chkDP.Size = New System.Drawing.Size(41, 17)
        Me.chkDP.TabIndex = 30
        Me.chkDP.Text = "DP"
        Me.chkDP.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(12, 220)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(85, 13)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "Nama Marketing"
        '
        'dbNamaMarketing
        '
        Me.dbNamaMarketing.AcceptsReturn = True
        Me.dbNamaMarketing.BackColor = System.Drawing.SystemColors.Window
        Me.dbNamaMarketing.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNamaMarketing.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNamaMarketing.Location = New System.Drawing.Point(118, 215)
        Me.dbNamaMarketing.MaxLength = 0
        Me.dbNamaMarketing.Name = "dbNamaMarketing"
        Me.dbNamaMarketing.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNamaMarketing.Size = New System.Drawing.Size(123, 20)
        Me.dbNamaMarketing.TabIndex = 31
        Me.dbNamaMarketing.Tag = "Jenis"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(251, 220)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(99, 13)
        Me.Label13.TabIndex = 34
        Me.Label13.Text = "Nama Customer SC"
        '
        'dbNamaCustomerSC
        '
        Me.dbNamaCustomerSC.AcceptsReturn = True
        Me.dbNamaCustomerSC.BackColor = System.Drawing.SystemColors.Window
        Me.dbNamaCustomerSC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNamaCustomerSC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNamaCustomerSC.Location = New System.Drawing.Point(359, 215)
        Me.dbNamaCustomerSC.MaxLength = 0
        Me.dbNamaCustomerSC.Name = "dbNamaCustomerSC"
        Me.dbNamaCustomerSC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNamaCustomerSC.Size = New System.Drawing.Size(123, 20)
        Me.dbNamaCustomerSC.TabIndex = 33
        Me.dbNamaCustomerSC.Tag = "Jenis"
        '
        'dg
        '
        Me.dg.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(12, 284)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(898, 303)
        Me.dg.TabIndex = 35
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(103, 253)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(83, 25)
        Me.btnSave.TabIndex = 36
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dbWaktuUpdate
        '
        Me.dbWaktuUpdate.AcceptsReturn = True
        Me.dbWaktuUpdate.BackColor = System.Drawing.SystemColors.Window
        Me.dbWaktuUpdate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbWaktuUpdate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbWaktuUpdate.Location = New System.Drawing.Point(359, 188)
        Me.dbWaktuUpdate.MaxLength = 0
        Me.dbWaktuUpdate.Name = "dbWaktuUpdate"
        Me.dbWaktuUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbWaktuUpdate.Size = New System.Drawing.Size(123, 20)
        Me.dbWaktuUpdate.TabIndex = 37
        Me.dbWaktuUpdate.Tag = "Jenis"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(14, 253)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(83, 25)
        Me.btnNew.TabIndex = 38
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrev
        '
        Me.btnPrev.Location = New System.Drawing.Point(256, 10)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(69, 25)
        Me.btnPrev.TabIndex = 39
        Me.btnPrev.Text = "Prev"
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(331, 10)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(69, 25)
        Me.btnNext.TabIndex = 40
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'txtStatus
        '
        Me.txtStatus.AcceptsReturn = True
        Me.txtStatus.BackColor = System.Drawing.SystemColors.Window
        Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStatus.Location = New System.Drawing.Point(564, 213)
        Me.txtStatus.MaxLength = 0
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatus.Size = New System.Drawing.Size(151, 20)
        Me.txtStatus.TabIndex = 42
        Me.txtStatus.Tag = "Jenis"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(521, 218)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(37, 13)
        Me.Label14.TabIndex = 41
        Me.Label14.Text = "Status"
        '
        'btnApprove
        '
        Me.btnApprove.Location = New System.Drawing.Point(192, 253)
        Me.btnApprove.Name = "btnApprove"
        Me.btnApprove.Size = New System.Drawing.Size(83, 25)
        Me.btnApprove.TabIndex = 43
        Me.btnApprove.Text = "Approve"
        Me.btnApprove.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(281, 253)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(83, 25)
        Me.btnClose.TabIndex = 44
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'frmTSC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(920, 600)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnApprove)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrev)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.dbWaktuUpdate)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dg)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.dbNamaCustomerSC)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.dbNamaMarketing)
        Me.Controls.Add(Me.chkDP)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.dbPengupdate)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.dbKeterangan)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dbNilaiKontrak)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.dbWaktuPembayaran)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dbLamaKontrak)
        Me.Controls.Add(Me.dbMataUang)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dbCustomerID)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dbTglSC)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dbNoSC)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.fQuickFind)
        Me.Name = "frmTSC"
        Me.Text = "frmTSC"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents fQuickFind As StringTextBoxNoKeyPreview
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents dbNoSC As StringTextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dbTglSC As ModDTPicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dbCustomerID As ModComboBox
    Friend WithEvents dbMataUang As ModComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents dbLamaKontrak As StringTextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents dbWaktuPembayaran As StringTextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents dbNilaiKontrak As NumericTextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents dbKeterangan As StringTextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents dbPengupdate As StringTextBoxNoKeyPreview
    Friend WithEvents chkDP As System.Windows.Forms.CheckBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents dbNamaMarketing As StringTextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents dbNamaCustomerSC As StringTextBox
    Friend WithEvents dg As modDataGridView
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Public WithEvents dbWaktuUpdate As StringTextBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnPrev As System.Windows.Forms.Button
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Public WithEvents txtStatus As StringTextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnApprove As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class

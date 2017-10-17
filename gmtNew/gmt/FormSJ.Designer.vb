<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTSJ
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
        Me.btnSave = New System.Windows.Forms.Button
        Me.dg = New modDataGridView
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.dbPengupdate = New StringTextBoxNoKeyPreview
        Me.Label9 = New System.Windows.Forms.Label
        Me.dbKeteranganSPP = New StringTextBox
        Me.dbTglSPP = New ModDTPicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dbNoSPP = New StringTextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.fQuickFind = New StringTextBoxNoKeyPreview
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.dbAlamatPenerima = New StringTextBox
        Me.dbTglKirim = New ModDTPicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.dbWaktuUpdate = New ModDTPicker
        Me.dbNamaPenerima = New StringTextBox
        Me.dbTglSJ = New ModDTPicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.dbNamaSopir = New StringTextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.dbNamaAngkutan = New StringTextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.dbNoKendaraan = New StringTextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.dbKeteranganSJ = New StringTextBox
        Me.txtStatus = New StringTextBox
        Me.lblStatus = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
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
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(14, 338)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(83, 25)
        Me.btnSave.TabIndex = 70
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dg
        '
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(15, 369)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(884, 266)
        Me.dg.TabIndex = 69
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(16, 295)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 62
        Me.Label10.Text = "Waktu Update"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(16, 265)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(65, 13)
        Me.Label11.TabIndex = 61
        Me.Label11.Text = "Pengupdate"
        '
        'dbPengupdate
        '
        Me.dbPengupdate.AcceptsReturn = True
        Me.dbPengupdate.BackColor = System.Drawing.SystemColors.Window
        Me.dbPengupdate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbPengupdate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbPengupdate.Location = New System.Drawing.Point(123, 265)
        Me.dbPengupdate.MaxLength = 0
        Me.dbPengupdate.Name = "dbPengupdate"
        Me.dbPengupdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbPengupdate.Size = New System.Drawing.Size(160, 20)
        Me.dbPengupdate.TabIndex = 60
        Me.dbPengupdate.Tag = "Jenis"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(15, 205)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 13)
        Me.Label9.TabIndex = 58
        Me.Label9.Text = "Keterangan SPP"
        '
        'dbKeteranganSPP
        '
        Me.dbKeteranganSPP.AcceptsReturn = True
        Me.dbKeteranganSPP.BackColor = System.Drawing.SystemColors.Window
        Me.dbKeteranganSPP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbKeteranganSPP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbKeteranganSPP.Location = New System.Drawing.Point(123, 200)
        Me.dbKeteranganSPP.MaxLength = 0
        Me.dbKeteranganSPP.Multiline = True
        Me.dbKeteranganSPP.Name = "dbKeteranganSPP"
        Me.dbKeteranganSPP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbKeteranganSPP.Size = New System.Drawing.Size(358, 56)
        Me.dbKeteranganSPP.TabIndex = 57
        Me.dbKeteranganSPP.Tag = "Jenis"
        '
        'dbTglSPP
        '
        Me.dbTglSPP.Checked = False
        Me.dbTglSPP.CustomFormat = "dd MMM yyyy"
        Me.dbTglSPP.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglSPP.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglSPP.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglSPP.Location = New System.Drawing.Point(124, 78)
        Me.dbTglSPP.Name = "dbTglSPP"
        Me.dbTglSPP.Size = New System.Drawing.Size(123, 20)
        Me.dbTglSPP.TabIndex = 46
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(18, 81)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 45
        Me.Label3.Text = "Tgl SPP"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 44
        Me.Label2.Text = "No SPP"
        '
        'dbNoSPP
        '
        Me.dbNoSPP.AcceptsReturn = True
        Me.dbNoSPP.BackColor = System.Drawing.SystemColors.Window
        Me.dbNoSPP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNoSPP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNoSPP.Location = New System.Drawing.Point(123, 52)
        Me.dbNoSPP.MaxLength = 0
        Me.dbNoSPP.Name = "dbNoSPP"
        Me.dbNoSPP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNoSPP.Size = New System.Drawing.Size(124, 20)
        Me.dbNoSPP.TabIndex = 43
        Me.dbNoSPP.Tag = "Jenis"
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
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(18, 109)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(82, 13)
        Me.Label12.TabIndex = 75
        Me.Label12.Text = "Nama Penerima"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(18, 138)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(86, 13)
        Me.Label13.TabIndex = 78
        Me.Label13.Text = "Alamat Penerima"
        '
        'dbAlamatPenerima
        '
        Me.dbAlamatPenerima.AcceptsReturn = True
        Me.dbAlamatPenerima.BackColor = System.Drawing.SystemColors.Window
        Me.dbAlamatPenerima.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbAlamatPenerima.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbAlamatPenerima.Location = New System.Drawing.Point(124, 133)
        Me.dbAlamatPenerima.MaxLength = 0
        Me.dbAlamatPenerima.Multiline = True
        Me.dbAlamatPenerima.Name = "dbAlamatPenerima"
        Me.dbAlamatPenerima.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbAlamatPenerima.Size = New System.Drawing.Size(245, 56)
        Me.dbAlamatPenerima.TabIndex = 77
        Me.dbAlamatPenerima.Tag = "Jenis"
        '
        'dbTglKirim
        '
        Me.dbTglKirim.Checked = False
        Me.dbTglKirim.CustomFormat = "dd MMM yyyy"
        Me.dbTglKirim.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglKirim.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglKirim.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglKirim.Location = New System.Drawing.Point(364, 79)
        Me.dbTglKirim.Name = "dbTglKirim"
        Me.dbTglKirim.Size = New System.Drawing.Size(118, 20)
        Me.dbTglKirim.TabIndex = 80
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(291, 83)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(47, 13)
        Me.Label8.TabIndex = 79
        Me.Label8.Text = "Tgl Kirim"
        '
        'dbWaktuUpdate
        '
        Me.dbWaktuUpdate.CustomFormat = "dd MMM yyyy HH:mm:ss"
        Me.dbWaktuUpdate.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbWaktuUpdate.DisabledForeColor = System.Drawing.Color.Black
        Me.dbWaktuUpdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbWaktuUpdate.Location = New System.Drawing.Point(122, 291)
        Me.dbWaktuUpdate.Name = "dbWaktuUpdate"
        Me.dbWaktuUpdate.Size = New System.Drawing.Size(160, 20)
        Me.dbWaktuUpdate.TabIndex = 84
        Me.dbWaktuUpdate.Value = New Date(2010, 11, 27, 0, 0, 0, 0)
        '
        'dbNamaPenerima
        '
        Me.dbNamaPenerima.AcceptsReturn = True
        Me.dbNamaPenerima.BackColor = System.Drawing.SystemColors.Window
        Me.dbNamaPenerima.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNamaPenerima.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNamaPenerima.Location = New System.Drawing.Point(123, 106)
        Me.dbNamaPenerima.MaxLength = 0
        Me.dbNamaPenerima.Name = "dbNamaPenerima"
        Me.dbNamaPenerima.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNamaPenerima.Size = New System.Drawing.Size(124, 20)
        Me.dbNamaPenerima.TabIndex = 85
        Me.dbNamaPenerima.Tag = "Jenis"
        '
        'dbTglSJ
        '
        Me.dbTglSJ.Checked = False
        Me.dbTglSJ.CustomFormat = "dd MMM yyyy"
        Me.dbTglSJ.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglSJ.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglSJ.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglSJ.Location = New System.Drawing.Point(651, 51)
        Me.dbTglSJ.Name = "dbTglSJ"
        Me.dbTglSJ.Size = New System.Drawing.Size(123, 20)
        Me.dbTglSJ.TabIndex = 87
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(505, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 13)
        Me.Label4.TabIndex = 86
        Me.Label4.Text = "Tgl SJ"
        '
        'dbNamaSopir
        '
        Me.dbNamaSopir.AcceptsReturn = True
        Me.dbNamaSopir.BackColor = System.Drawing.SystemColors.Window
        Me.dbNamaSopir.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNamaSopir.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNamaSopir.Location = New System.Drawing.Point(650, 78)
        Me.dbNamaSopir.MaxLength = 0
        Me.dbNamaSopir.Name = "dbNamaSopir"
        Me.dbNamaSopir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNamaSopir.Size = New System.Drawing.Size(124, 20)
        Me.dbNamaSopir.TabIndex = 89
        Me.dbNamaSopir.Tag = "Jenis"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(505, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 13)
        Me.Label5.TabIndex = 88
        Me.Label5.Text = "Nama Sopir"
        '
        'dbNamaAngkutan
        '
        Me.dbNamaAngkutan.AcceptsReturn = True
        Me.dbNamaAngkutan.BackColor = System.Drawing.SystemColors.Window
        Me.dbNamaAngkutan.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNamaAngkutan.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNamaAngkutan.Location = New System.Drawing.Point(651, 106)
        Me.dbNamaAngkutan.MaxLength = 0
        Me.dbNamaAngkutan.Name = "dbNamaAngkutan"
        Me.dbNamaAngkutan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNamaAngkutan.Size = New System.Drawing.Size(124, 20)
        Me.dbNamaAngkutan.TabIndex = 91
        Me.dbNamaAngkutan.Tag = "Jenis"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(506, 113)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(84, 13)
        Me.Label6.TabIndex = 90
        Me.Label6.Text = "Nama Angkutan"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(505, 184)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(77, 13)
        Me.Label7.TabIndex = 92
        Me.Label7.Text = "Keterangan SJ"
        '
        'dbNoKendaraan
        '
        Me.dbNoKendaraan.AcceptsReturn = True
        Me.dbNoKendaraan.BackColor = System.Drawing.SystemColors.Window
        Me.dbNoKendaraan.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNoKendaraan.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNoKendaraan.Location = New System.Drawing.Point(651, 133)
        Me.dbNoKendaraan.MaxLength = 0
        Me.dbNoKendaraan.Name = "dbNoKendaraan"
        Me.dbNoKendaraan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNoKendaraan.Size = New System.Drawing.Size(128, 20)
        Me.dbNoKendaraan.TabIndex = 95
        Me.dbNoKendaraan.Tag = "Jenis"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(506, 138)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(76, 13)
        Me.Label14.TabIndex = 94
        Me.Label14.Text = "No Kendaraan"
        '
        'dbKeteranganSJ
        '
        Me.dbKeteranganSJ.AcceptsReturn = True
        Me.dbKeteranganSJ.BackColor = System.Drawing.SystemColors.Window
        Me.dbKeteranganSJ.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbKeteranganSJ.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbKeteranganSJ.Location = New System.Drawing.Point(504, 200)
        Me.dbKeteranganSJ.MaxLength = 0
        Me.dbKeteranganSJ.Multiline = True
        Me.dbKeteranganSJ.Name = "dbKeteranganSJ"
        Me.dbKeteranganSJ.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbKeteranganSJ.Size = New System.Drawing.Size(358, 56)
        Me.dbKeteranganSJ.TabIndex = 96
        Me.dbKeteranganSJ.Tag = "Jenis"
        '
        'txtStatus
        '
        Me.txtStatus.AcceptsReturn = True
        Me.txtStatus.BackColor = System.Drawing.SystemColors.Window
        Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStatus.Location = New System.Drawing.Point(608, 265)
        Me.txtStatus.MaxLength = 0
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatus.Size = New System.Drawing.Size(123, 20)
        Me.txtStatus.TabIndex = 98
        Me.txtStatus.Tag = "Jenis"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(506, 270)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(37, 13)
        Me.lblStatus.TabIndex = 97
        Me.lblStatus.Text = "Status"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(103, 338)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(83, 25)
        Me.btnCancel.TabIndex = 99
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmTSJ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(907, 647)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.dbKeteranganSJ)
        Me.Controls.Add(Me.dbNoKendaraan)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.dbNamaAngkutan)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dbNamaSopir)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dbTglSJ)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dbNamaPenerima)
        Me.Controls.Add(Me.dbWaktuUpdate)
        Me.Controls.Add(Me.dbTglKirim)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.dbAlamatPenerima)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrev)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dg)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.dbPengupdate)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.dbKeteranganSPP)
        Me.Controls.Add(Me.dbTglSPP)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dbNoSPP)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.fQuickFind)
        Me.Name = "frmTSJ"
        Me.Text = "frmTSJ"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnPrev As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dg As modDataGridView
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents dbPengupdate As StringTextBoxNoKeyPreview
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents dbKeteranganSPP As StringTextBox
    Friend WithEvents dbTglSPP As ModDTPicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents dbNoSPP As StringTextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents fQuickFind As StringTextBoxNoKeyPreview
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents dbAlamatPenerima As StringTextBox
    Friend WithEvents dbTglKirim As ModDTPicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dbWaktuUpdate As ModDTPicker
    Public WithEvents dbNamaPenerima As StringTextBox
    Friend WithEvents dbTglSJ As ModDTPicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents dbNamaSopir As StringTextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents dbNamaAngkutan As StringTextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents dbNoKendaraan As StringTextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents dbKeteranganSJ As StringTextBox
    Public WithEvents txtStatus As StringTextBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button

End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTKW
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
        Me.txtNilai = New NumericTextBox
        Me.dbMataUang = New ModComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dbNoKW = New StringTextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dbCustomerID = New ModComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.dg = New modDataGridView
        Me.dgDetail = New modDataGridView
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnNext = New System.Windows.Forms.Button
        Me.btnPrev = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.fQuickFind = New StringTextBoxNoKeyPreview
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.dbPengupdate = New StringTextBoxNoKeyPreview
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.dbWaktuUpdate = New ModDTPicker
        Me.dbTglKW = New ModDTPicker
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtNilai
        '
        Me.txtNilai.AcceptsReturn = True
        Me.txtNilai.BackColor = System.Drawing.SystemColors.Window
        Me.txtNilai.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNilai.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNilai.Location = New System.Drawing.Point(129, 84)
        Me.txtNilai.MaxLength = 0
        Me.txtNilai.Name = "txtNilai"
        Me.txtNilai.NumberFormat = "#,##0.00"
        Me.txtNilai.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNilai.Size = New System.Drawing.Size(123, 20)
        Me.txtNilai.TabIndex = 27
        Me.txtNilai.Tag = "Jenis"
        Me.txtNilai.Text = "0.00"
        Me.txtNilai.Value = New Decimal(New Integer() {0, 0, 0, 131072})
        '
        'dbMataUang
        '
        Me.dbMataUang.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbMataUang.FormattingEnabled = True
        Me.dbMataUang.Location = New System.Drawing.Point(369, 84)
        Me.dbMataUang.Name = "dbMataUang"
        Me.dbMataUang.Size = New System.Drawing.Size(121, 21)
        Me.dbMataUang.TabIndex = 26
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(264, 87)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Mata Uang"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(263, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Tgl KW"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "No KW"
        '
        'dbNoKW
        '
        Me.dbNoKW.AcceptsReturn = True
        Me.dbNoKW.BackColor = System.Drawing.SystemColors.Window
        Me.dbNoKW.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbNoKW.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbNoKW.Location = New System.Drawing.Point(128, 56)
        Me.dbNoKW.MaxLength = 0
        Me.dbNoKW.Name = "dbNoKW"
        Me.dbNoKW.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbNoKW.Size = New System.Drawing.Size(124, 20)
        Me.dbNoKW.TabIndex = 21
        Me.dbNoKW.Tag = "Jenis"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 87)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Nilai"
        '
        'dbCustomerID
        '
        Me.dbCustomerID.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbCustomerID.FormattingEnabled = True
        Me.dbCustomerID.Location = New System.Drawing.Point(127, 111)
        Me.dbCustomerID.Name = "dbCustomerID"
        Me.dbCustomerID.Size = New System.Drawing.Size(362, 21)
        Me.dbCustomerID.TabIndex = 50
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(18, 114)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 49
        Me.Label4.Text = "Nama Customer"
        '
        'dg
        '
        Me.dg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(21, 184)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(933, 174)
        Me.dg.TabIndex = 51
        '
        'dgDetail
        '
        Me.dgDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgDetail.IsOnPilih = False
        Me.dgDetail.LastSearch = Nothing
        Me.dgDetail.Location = New System.Drawing.Point(21, 377)
        Me.dgDetail.Name = "dgDetail"
        Me.dgDetail.Size = New System.Drawing.Size(933, 197)
        Me.dgDetail.TabIndex = 52
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(20, 361)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(34, 13)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Detail"
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(338, 11)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(69, 25)
        Me.btnNext.TabIndex = 78
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnPrev
        '
        Me.btnPrev.Location = New System.Drawing.Point(263, 11)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(69, 25)
        Me.btnPrev.TabIndex = 77
        Me.btnPrev.Text = "Prev"
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(18, 18)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 13)
        Me.Label7.TabIndex = 76
        Me.Label7.Text = "Quick Find"
        '
        'fQuickFind
        '
        Me.fQuickFind.AcceptsReturn = True
        Me.fQuickFind.BackColor = System.Drawing.SystemColors.Window
        Me.fQuickFind.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.fQuickFind.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fQuickFind.Location = New System.Drawing.Point(126, 13)
        Me.fQuickFind.MaxLength = 0
        Me.fQuickFind.Name = "fQuickFind"
        Me.fQuickFind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fQuickFind.Size = New System.Drawing.Size(124, 20)
        Me.fQuickFind.TabIndex = 75
        Me.fQuickFind.Tag = "Jenis"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(264, 61)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(67, 13)
        Me.Label8.TabIndex = 79
        Me.Label8.Text = "Tanggal KW"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(498, 110)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 87
        Me.Label10.Text = "Waktu Update"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(498, 80)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(65, 13)
        Me.Label11.TabIndex = 86
        Me.Label11.Text = "Pengupdate"
        '
        'dbPengupdate
        '
        Me.dbPengupdate.AcceptsReturn = True
        Me.dbPengupdate.BackColor = System.Drawing.SystemColors.Window
        Me.dbPengupdate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbPengupdate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbPengupdate.Location = New System.Drawing.Point(605, 80)
        Me.dbPengupdate.MaxLength = 0
        Me.dbPengupdate.Name = "dbPengupdate"
        Me.dbPengupdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbPengupdate.Size = New System.Drawing.Size(160, 20)
        Me.dbPengupdate.TabIndex = 85
        Me.dbPengupdate.Tag = "Jenis"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(109, 153)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(83, 25)
        Me.btnCancel.TabIndex = 101
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(20, 153)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(83, 25)
        Me.btnSave.TabIndex = 100
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(198, 153)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(83, 25)
        Me.btnNew.TabIndex = 102
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'dbWaktuUpdate
        '
        Me.dbWaktuUpdate.Checked = False
        Me.dbWaktuUpdate.CustomFormat = "dd MMM yyyy HH:mm:ss"
        Me.dbWaktuUpdate.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbWaktuUpdate.DisabledForeColor = System.Drawing.Color.Black
        Me.dbWaktuUpdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbWaktuUpdate.Location = New System.Drawing.Point(604, 106)
        Me.dbWaktuUpdate.Name = "dbWaktuUpdate"
        Me.dbWaktuUpdate.Size = New System.Drawing.Size(160, 20)
        Me.dbWaktuUpdate.TabIndex = 88
        '
        'dbTglKW
        '
        Me.dbTglKW.Checked = False
        Me.dbTglKW.CustomFormat = "dd MMM yyyy"
        Me.dbTglKW.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dbTglKW.DisabledForeColor = System.Drawing.Color.Black
        Me.dbTglKW.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dbTglKW.Location = New System.Drawing.Point(371, 57)
        Me.dbTglKW.Name = "dbTglKW"
        Me.dbTglKW.Size = New System.Drawing.Size(118, 20)
        Me.dbTglKW.TabIndex = 24
        '
        'frmTKW
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(970, 586)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dbWaktuUpdate)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.dbPengupdate)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrev)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.fQuickFind)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dgDetail)
        Me.Controls.Add(Me.dg)
        Me.Controls.Add(Me.dbCustomerID)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNilai)
        Me.Controls.Add(Me.dbMataUang)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dbTglKW)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dbNoKW)
        Me.Name = "frmTKW"
        Me.Text = "frmTKW"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents txtNilai As NumericTextBox
    Friend WithEvents dbMataUang As ModComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dbTglKW As ModDTPicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents dbNoKW As StringTextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dbCustomerID As ModComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dg As modDataGridView
    Friend WithEvents dgDetail As modDataGridView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnPrev As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents fQuickFind As StringTextBoxNoKeyPreview
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dbWaktuUpdate As ModDTPicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents dbPengupdate As StringTextBoxNoKeyPreview
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
End Class

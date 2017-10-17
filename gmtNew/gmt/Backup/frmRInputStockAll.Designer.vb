<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRInputStockAll
    'Inherits FormMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRInputStockAll))
        Me.dg = New modDataGridView
        Me.dtTgl = New ModDTPicker
        Me.btnSearch = New System.Windows.Forms.Button
        Me.dgDetail = New System.Windows.Forms.DataGridView
        Me.txtKg = New NumericTextBox
        Me.txtCns = New NumericTextBox
        Me.txtNoUrut = New NumericTextBox
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.dbNoBukti = New System.Windows.Forms.TextBox
        Me.dbTube = New System.Windows.Forms.TextBox
        Me.dbJenis = New System.Windows.Forms.TextBox
        Me.dbNoWarna = New System.Windows.Forms.TextBox
        Me.dbWarna = New System.Windows.Forms.TextBox
        Me.dbKodeBarang = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.dbGrade = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.dbSatKecil = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.dbSatBesar = New System.Windows.Forms.TextBox
        Me.dbn1 = New NumericTextBox
        Me.dbn2 = New NumericTextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.dbKeterangan = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label12 = New System.Windows.Forms.Label
        Me.dbShift = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.dbLot = New System.Windows.Forms.TextBox
        Me.btnClear = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnTipe = New System.Windows.Forms.Button
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog
        Me.lblTipe = New System.Windows.Forms.Label
        CType(Me.dg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'dg
        '
        Me.dg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg.IsOnPilih = False
        Me.dg.LastSearch = Nothing
        Me.dg.Location = New System.Drawing.Point(12, 65)
        Me.dg.Name = "dg"
        Me.dg.Size = New System.Drawing.Size(710, 171)
        Me.dg.TabIndex = 0
        '
        'dtTgl
        '
        Me.dtTgl.CustomFormat = "dd MMM yyyy"
        Me.dtTgl.DisabledBackColor = System.Drawing.Color.Gainsboro
        Me.dtTgl.DisabledForeColor = System.Drawing.Color.Black
        Me.dtTgl.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTgl.Location = New System.Drawing.Point(12, 39)
        Me.dtTgl.Name = "dtTgl"
        Me.dtTgl.Size = New System.Drawing.Size(120, 20)
        Me.dtTgl.TabIndex = 1
        Me.dtTgl.Value = New Date(2011, 4, 22, 0, 0, 0, 0)
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(156, 28)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(82, 30)
        Me.btnSearch.TabIndex = 2
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'dgDetail
        '
        Me.dgDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgDetail.Location = New System.Drawing.Point(89, 8)
        Me.dgDetail.Name = "dgDetail"
        Me.dgDetail.Size = New System.Drawing.Size(219, 181)
        Me.dgDetail.TabIndex = 3
        '
        'txtKg
        '
        Me.txtKg.AcceptsReturn = True
        Me.txtKg.BackColor = System.Drawing.SystemColors.Window
        Me.txtKg.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtKg.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtKg.Location = New System.Drawing.Point(11, 8)
        Me.txtKg.MaxLength = 0
        Me.txtKg.Name = "txtKg"
        Me.txtKg.NumberFormat = "#,##0.00"
        Me.txtKg.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtKg.Size = New System.Drawing.Size(72, 20)
        Me.txtKg.TabIndex = 21
        Me.txtKg.Tag = ""
        Me.txtKg.Text = "0,00"
        Me.txtKg.Value = New Decimal(New Integer() {0, 0, 0, 131072})
        '
        'txtCns
        '
        Me.txtCns.AcceptsReturn = True
        Me.txtCns.BackColor = System.Drawing.SystemColors.Window
        Me.txtCns.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCns.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCns.Location = New System.Drawing.Point(11, 34)
        Me.txtCns.MaxLength = 0
        Me.txtCns.Name = "txtCns"
        Me.txtCns.NumberFormat = "#,##0"
        Me.txtCns.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCns.Size = New System.Drawing.Size(44, 20)
        Me.txtCns.TabIndex = 22
        Me.txtCns.Tag = ""
        Me.txtCns.Text = "24"
        Me.txtCns.Value = New Decimal(New Integer() {24, 0, 0, 0})
        '
        'txtNoUrut
        '
        Me.txtNoUrut.AcceptsReturn = True
        Me.txtNoUrut.BackColor = System.Drawing.SystemColors.Window
        Me.txtNoUrut.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNoUrut.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNoUrut.Location = New System.Drawing.Point(11, 60)
        Me.txtNoUrut.MaxLength = 0
        Me.txtNoUrut.Name = "txtNoUrut"
        Me.txtNoUrut.NumberFormat = "#,##0"
        Me.txtNoUrut.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNoUrut.Size = New System.Drawing.Size(72, 20)
        Me.txtNoUrut.TabIndex = 23
        Me.txtNoUrut.Tag = ""
        Me.txtNoUrut.Text = "1"
        Me.txtNoUrut.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(11, 86)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(63, 23)
        Me.btnAdd.TabIndex = 24
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(11, 115)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(63, 23)
        Me.btnSave.TabIndex = 25
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dbNoBukti
        '
        Me.dbNoBukti.Location = New System.Drawing.Point(9, 24)
        Me.dbNoBukti.Name = "dbNoBukti"
        Me.dbNoBukti.Size = New System.Drawing.Size(100, 20)
        Me.dbNoBukti.TabIndex = 26
        '
        'dbTube
        '
        Me.dbTube.Location = New System.Drawing.Point(222, 63)
        Me.dbTube.Name = "dbTube"
        Me.dbTube.Size = New System.Drawing.Size(64, 20)
        Me.dbTube.TabIndex = 27
        '
        'dbJenis
        '
        Me.dbJenis.Location = New System.Drawing.Point(9, 63)
        Me.dbJenis.Name = "dbJenis"
        Me.dbJenis.Size = New System.Drawing.Size(100, 20)
        Me.dbJenis.TabIndex = 28
        '
        'dbNoWarna
        '
        Me.dbNoWarna.Location = New System.Drawing.Point(113, 102)
        Me.dbNoWarna.Name = "dbNoWarna"
        Me.dbNoWarna.Size = New System.Drawing.Size(100, 20)
        Me.dbNoWarna.TabIndex = 29
        '
        'dbWarna
        '
        Me.dbWarna.Location = New System.Drawing.Point(9, 102)
        Me.dbWarna.Name = "dbWarna"
        Me.dbWarna.Size = New System.Drawing.Size(100, 20)
        Me.dbWarna.TabIndex = 30
        '
        'dbKodeBarang
        '
        Me.dbKodeBarang.Location = New System.Drawing.Point(113, 63)
        Me.dbKodeBarang.Name = "dbKodeBarang"
        Me.dbKodeBarang.Size = New System.Drawing.Size(100, 20)
        Me.dbKodeBarang.TabIndex = 31
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "No Bukti"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(31, 13)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Jenis"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(111, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 13)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Kode Barang"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 35
        Me.Label4.Text = "Warna"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(110, 86)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 13)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "No Warna"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(222, 47)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "Tube"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(288, 47)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 13)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Grade"
        '
        'dbGrade
        '
        Me.dbGrade.Location = New System.Drawing.Point(291, 63)
        Me.dbGrade.Name = "dbGrade"
        Me.dbGrade.Size = New System.Drawing.Size(44, 20)
        Me.dbGrade.TabIndex = 38
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(286, 86)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(49, 13)
        Me.Label8.TabIndex = 44
        Me.Label8.Text = "Sat Kecil"
        '
        'dbSatKecil
        '
        Me.dbSatKecil.Location = New System.Drawing.Point(289, 102)
        Me.dbSatKecil.Name = "dbSatKecil"
        Me.dbSatKecil.Size = New System.Drawing.Size(44, 20)
        Me.dbSatKecil.TabIndex = 43
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(222, 86)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(53, 13)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "Sat Besar"
        '
        'dbSatBesar
        '
        Me.dbSatBesar.Location = New System.Drawing.Point(220, 102)
        Me.dbSatBesar.Name = "dbSatBesar"
        Me.dbSatBesar.Size = New System.Drawing.Size(64, 20)
        Me.dbSatBesar.TabIndex = 41
        '
        'dbn1
        '
        Me.dbn1.AcceptsReturn = True
        Me.dbn1.BackColor = System.Drawing.SystemColors.Window
        Me.dbn1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbn1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbn1.Location = New System.Drawing.Point(341, 84)
        Me.dbn1.MaxLength = 0
        Me.dbn1.Name = "dbn1"
        Me.dbn1.NumberFormat = "#,##0"
        Me.dbn1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbn1.Size = New System.Drawing.Size(44, 20)
        Me.dbn1.TabIndex = 46
        Me.dbn1.Tag = ""
        Me.dbn1.Text = "0"
        Me.dbn1.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'dbn2
        '
        Me.dbn2.AcceptsReturn = True
        Me.dbn2.BackColor = System.Drawing.SystemColors.Window
        Me.dbn2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.dbn2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.dbn2.Location = New System.Drawing.Point(341, 110)
        Me.dbn2.MaxLength = 0
        Me.dbn2.Name = "dbn2"
        Me.dbn2.NumberFormat = "#,##0.00"
        Me.dbn2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dbn2.Size = New System.Drawing.Size(44, 20)
        Me.dbn2.TabIndex = 45
        Me.dbn2.Tag = ""
        Me.dbn2.Text = "0,00"
        Me.dbn2.Value = New Decimal(New Integer() {0, 0, 0, 131072})
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(9, 130)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 48
        Me.Label10.Text = "Keterangan"
        '
        'dbKeterangan
        '
        Me.dbKeterangan.Location = New System.Drawing.Point(7, 146)
        Me.dbKeterangan.Name = "dbKeterangan"
        Me.dbKeterangan.Size = New System.Drawing.Size(179, 20)
        Me.dbKeterangan.TabIndex = 47
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.dbShift)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.dbLot)
        Me.Panel1.Controls.Add(Me.btnClear)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.dbKeterangan)
        Me.Panel1.Controls.Add(Me.dbn1)
        Me.Panel1.Controls.Add(Me.dbn2)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.dbSatKecil)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.dbSatBesar)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.dbGrade)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.dbKodeBarang)
        Me.Panel1.Controls.Add(Me.dbWarna)
        Me.Panel1.Controls.Add(Me.dbNoWarna)
        Me.Panel1.Controls.Add(Me.dbJenis)
        Me.Panel1.Controls.Add(Me.dbTube)
        Me.Panel1.Controls.Add(Me.dbNoBukti)
        Me.Panel1.Location = New System.Drawing.Point(12, 242)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(396, 189)
        Me.Panel1.TabIndex = 49
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(217, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(28, 13)
        Me.Label12.TabIndex = 53
        Me.Label12.Text = "Shift"
        '
        'dbShift
        '
        Me.dbShift.Location = New System.Drawing.Point(220, 24)
        Me.dbShift.Name = "dbShift"
        Me.dbShift.Size = New System.Drawing.Size(100, 20)
        Me.dbShift.TabIndex = 52
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(112, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(22, 13)
        Me.Label11.TabIndex = 51
        Me.Label11.Text = "Lot"
        '
        'dbLot
        '
        Me.dbLot.Location = New System.Drawing.Point(115, 24)
        Me.dbLot.Name = "dbLot"
        Me.dbLot.Size = New System.Drawing.Size(100, 20)
        Me.dbLot.TabIndex = 50
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(341, 29)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(44, 36)
        Me.btnClear.TabIndex = 49
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.btnSave)
        Me.Panel2.Controls.Add(Me.btnAdd)
        Me.Panel2.Controls.Add(Me.txtCns)
        Me.Panel2.Controls.Add(Me.txtNoUrut)
        Me.Panel2.Controls.Add(Me.txtKg)
        Me.Panel2.Controls.Add(Me.dgDetail)
        Me.Panel2.Location = New System.Drawing.Point(414, 237)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(314, 194)
        Me.Panel2.TabIndex = 50
        '
        'btnTipe
        '
        Me.btnTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTipe.Location = New System.Drawing.Point(76, 5)
        Me.btnTipe.Name = "btnTipe"
        Me.btnTipe.Size = New System.Drawing.Size(56, 28)
        Me.btnTipe.TabIndex = 52
        Me.btnTipe.Text = "Ubah"
        Me.btnTipe.UseVisualStyleBackColor = False
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Visible = False
        '
        'lblTipe
        '
        Me.lblTipe.AutoSize = True
        Me.lblTipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTipe.Location = New System.Drawing.Point(15, 7)
        Me.lblTipe.Name = "lblTipe"
        Me.lblTipe.Size = New System.Drawing.Size(50, 24)
        Me.lblTipe.TabIndex = 55
        Me.lblTipe.Text = "DTY"
        '
        'frmRInputStockAll
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(734, 437)
        Me.Controls.Add(Me.lblTipe)
        Me.Controls.Add(Me.btnTipe)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.dtTgl)
        Me.Controls.Add(Me.dg)
        Me.Name = "frmRInputStockAll"
        Me.Text = "frmtInputStock"
        CType(Me.dg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dg As modDataGridView
    Friend WithEvents dtTgl As ModDTPicker
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents dgDetail As System.Windows.Forms.DataGridView
    Public WithEvents txtKg As NumericTextBox
    Public WithEvents txtCns As NumericTextBox
    Public WithEvents txtNoUrut As NumericTextBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dbNoBukti As System.Windows.Forms.TextBox
    Friend WithEvents dbTube As System.Windows.Forms.TextBox
    Friend WithEvents dbJenis As System.Windows.Forms.TextBox
    Friend WithEvents dbNoWarna As System.Windows.Forms.TextBox
    Friend WithEvents dbWarna As System.Windows.Forms.TextBox
    Friend WithEvents dbKodeBarang As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents dbGrade As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dbSatKecil As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents dbSatBesar As System.Windows.Forms.TextBox
    Public WithEvents dbn1 As NumericTextBox
    Public WithEvents dbn2 As NumericTextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents dbKeterangan As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnTipe As System.Windows.Forms.Button
    Friend WithEvents PrintPreviewDialog1 As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents lblTipe As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents dbLot As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents dbShift As System.Windows.Forms.TextBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class rptMutasi
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
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblHeader = New System.Windows.Forms.Label()
        Me.dTgl = New System.Windows.Forms.Label()
        Me.PageHeader = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblBoxAkhir = New System.Windows.Forms.Label()
        Me.lblKgAkhir = New System.Windows.Forms.Label()
        Me.lblBoxOut = New System.Windows.Forms.Label()
        Me.lblKgOut = New System.Windows.Forms.Label()
        Me.lblBox = New System.Windows.Forms.Label()
        Me.lblKg = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblNo = New System.Windows.Forms.Label()
        Me.lblJenisBarang = New System.Windows.Forms.Label()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineAkhirKg = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineOutKg = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineInKg = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineLast = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineAkhirBox = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineShape7 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineOutBox = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineBottom = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineInBox = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineJenis = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineNo = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.GroupHeader = New System.Windows.Forms.GroupBox()
        Me.Detail = New System.Windows.Forms.GroupBox()
        Me.TextBox9 = New System.Windows.Forms.TextBox()
        Me.TextBox10 = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.txtNo = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupFooter = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.takhirBox = New System.Windows.Forms.TextBox()
        Me.toutBox = New System.Windows.Forms.TextBox()
        Me.takhirKg = New System.Windows.Forms.TextBox()
        Me.tinKg = New System.Windows.Forms.TextBox()
        Me.tinBox = New System.Windows.Forms.TextBox()
        Me.toutKg = New System.Windows.Forms.TextBox()
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog()
        Me.PageHeader.SuspendLayout()
        Me.GroupHeader.SuspendLayout()
        Me.Detail.SuspendLayout()
        Me.GroupFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(56, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(75, 15)
        Me.Label7.TabIndex = 13
        Me.Label7.Tag = "Group"
        Me.Label7.Text = "KodeBarang"
        '
        'lblHeader
        '
        Me.lblHeader.AutoSize = True
        Me.lblHeader.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeader.Location = New System.Drawing.Point(20, 26)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(54, 15)
        Me.lblHeader.TabIndex = 0
        Me.lblHeader.Tag = "Header"
        Me.lblHeader.Text = "MUTASI"
        '
        'dTgl
        '
        Me.dTgl.AutoSize = True
        Me.dTgl.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dTgl.Location = New System.Drawing.Point(80, 26)
        Me.dTgl.Name = "dTgl"
        Me.dTgl.Size = New System.Drawing.Size(50, 15)
        Me.dTgl.TabIndex = 1
        Me.dTgl.Tag = "Header"
        Me.dTgl.Text = "Tanggal"
        '
        'PageHeader
        '
        Me.PageHeader.Controls.Add(Me.Label12)
        Me.PageHeader.Controls.Add(Me.Label6)
        Me.PageHeader.Controls.Add(Me.lblBoxAkhir)
        Me.PageHeader.Controls.Add(Me.lblKgAkhir)
        Me.PageHeader.Controls.Add(Me.lblBoxOut)
        Me.PageHeader.Controls.Add(Me.lblKgOut)
        Me.PageHeader.Controls.Add(Me.lblBox)
        Me.PageHeader.Controls.Add(Me.lblKg)
        Me.PageHeader.Controls.Add(Me.Label10)
        Me.PageHeader.Controls.Add(Me.lblNo)
        Me.PageHeader.Controls.Add(Me.lblJenisBarang)
        Me.PageHeader.Controls.Add(Me.lblHeader)
        Me.PageHeader.Controls.Add(Me.dTgl)
        Me.PageHeader.Controls.Add(Me.ShapeContainer1)
        Me.PageHeader.Location = New System.Drawing.Point(4, 7)
        Me.PageHeader.Name = "PageHeader"
        Me.PageHeader.Size = New System.Drawing.Size(772, 97)
        Me.PageHeader.TabIndex = 28
        Me.PageHeader.TabStop = False
        Me.PageHeader.Text = "PageHeader"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(604, 52)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(57, 15)
        Me.Label12.TabIndex = 23
        Me.Label12.Tag = "Header"
        Me.Label12.Text = "Akhir"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(480, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(58, 15)
        Me.Label6.TabIndex = 22
        Me.Label6.Tag = "Header"
        Me.Label6.Text = "Out"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblBoxAkhir
        '
        Me.lblBoxAkhir.AutoSize = True
        Me.lblBoxAkhir.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoxAkhir.Location = New System.Drawing.Point(580, 72)
        Me.lblBoxAkhir.Name = "lblBoxAkhir"
        Me.lblBoxAkhir.Size = New System.Drawing.Size(28, 15)
        Me.lblBoxAkhir.TabIndex = 21
        Me.lblBoxAkhir.Tag = "Header"
        Me.lblBoxAkhir.Text = "Box"
        Me.lblBoxAkhir.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblKgAkhir
        '
        Me.lblKgAkhir.AutoSize = True
        Me.lblKgAkhir.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKgAkhir.Location = New System.Drawing.Point(640, 72)
        Me.lblKgAkhir.Name = "lblKgAkhir"
        Me.lblKgAkhir.Size = New System.Drawing.Size(23, 15)
        Me.lblKgAkhir.TabIndex = 20
        Me.lblKgAkhir.Tag = "Header"
        Me.lblKgAkhir.Text = "Kg"
        Me.lblKgAkhir.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblBoxOut
        '
        Me.lblBoxOut.AutoSize = True
        Me.lblBoxOut.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoxOut.Location = New System.Drawing.Point(461, 72)
        Me.lblBoxOut.Name = "lblBoxOut"
        Me.lblBoxOut.Size = New System.Drawing.Size(28, 15)
        Me.lblBoxOut.TabIndex = 18
        Me.lblBoxOut.Tag = "Header"
        Me.lblBoxOut.Text = "Box"
        Me.lblBoxOut.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblKgOut
        '
        Me.lblKgOut.AutoSize = True
        Me.lblKgOut.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKgOut.Location = New System.Drawing.Point(524, 72)
        Me.lblKgOut.Name = "lblKgOut"
        Me.lblKgOut.Size = New System.Drawing.Size(23, 15)
        Me.lblKgOut.TabIndex = 17
        Me.lblKgOut.Tag = "Header"
        Me.lblKgOut.Text = "Kg"
        Me.lblKgOut.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblBox
        '
        Me.lblBox.AutoSize = True
        Me.lblBox.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBox.Location = New System.Drawing.Point(340, 72)
        Me.lblBox.Name = "lblBox"
        Me.lblBox.Size = New System.Drawing.Size(28, 15)
        Me.lblBox.TabIndex = 15
        Me.lblBox.Tag = "Header"
        Me.lblBox.Text = "Box"
        Me.lblBox.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblKg
        '
        Me.lblKg.AutoSize = True
        Me.lblKg.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKg.Location = New System.Drawing.Point(408, 72)
        Me.lblKg.Name = "lblKg"
        Me.lblKg.Size = New System.Drawing.Size(23, 15)
        Me.lblKg.TabIndex = 14
        Me.lblKg.Tag = "Header"
        Me.lblKg.Text = "Kg"
        Me.lblKg.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(359, 52)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(60, 15)
        Me.Label10.TabIndex = 13
        Me.Label10.Tag = "Header"
        Me.Label10.Text = "In"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblNo
        '
        Me.lblNo.AutoSize = True
        Me.lblNo.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNo.Location = New System.Drawing.Point(27, 60)
        Me.lblNo.Name = "lblNo"
        Me.lblNo.Size = New System.Drawing.Size(22, 15)
        Me.lblNo.TabIndex = 11
        Me.lblNo.Tag = "Header"
        Me.lblNo.Text = "No"
        '
        'lblJenisBarang
        '
        Me.lblJenisBarang.AutoSize = True
        Me.lblJenisBarang.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJenisBarang.Location = New System.Drawing.Point(115, 60)
        Me.lblJenisBarang.Name = "lblJenisBarang"
        Me.lblJenisBarang.Size = New System.Drawing.Size(81, 15)
        Me.lblJenisBarang.TabIndex = 10
        Me.lblJenisBarang.Tag = "Header"
        Me.lblJenisBarang.Text = "Jenis Barang"
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(3, 16)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineAkhirKg, Me.LineOutKg, Me.LineInKg, Me.LineLast, Me.LineAkhirBox, Me.LineShape7, Me.LineOutBox, Me.LineBottom, Me.LineInBox, Me.LineJenis, Me.LineNo, Me.LineShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(766, 78)
        Me.ShapeContainer1.TabIndex = 24
        Me.ShapeContainer1.TabStop = False
        '
        'LineAkhirKg
        '
        Me.LineAkhirKg.Name = "LineAkhirKg"
        Me.LineAkhirKg.X1 = 617
        Me.LineAkhirKg.X2 = 617
        Me.LineAkhirKg.Y1 = 52
        Me.LineAkhirKg.Y2 = 72
        '
        'LineOutKg
        '
        Me.LineOutKg.Name = "LineOutKg"
        Me.LineOutKg.X1 = 497
        Me.LineOutKg.X2 = 497
        Me.LineOutKg.Y1 = 52
        Me.LineOutKg.Y2 = 72
        '
        'LineInKg
        '
        Me.LineInKg.Name = "LineInKg"
        Me.LineInKg.X1 = 377
        Me.LineInKg.X2 = 377
        Me.LineInKg.Y1 = 52
        Me.LineInKg.Y2 = 72
        '
        'LineLast
        '
        Me.LineLast.Name = "LineLast"
        Me.LineLast.X1 = 687
        Me.LineLast.X2 = 687
        Me.LineLast.Y1 = 32
        Me.LineLast.Y2 = 71
        '
        'LineAkhirBox
        '
        Me.LineAkhirBox.Name = "LineAkhirBox"
        Me.LineAkhirBox.X1 = 567
        Me.LineAkhirBox.X2 = 567
        Me.LineAkhirBox.Y1 = 32
        Me.LineAkhirBox.Y2 = 73
        '
        'LineShape7
        '
        Me.LineShape7.Name = "LineShape7"
        Me.LineShape7.X1 = 327
        Me.LineShape7.X2 = 687
        Me.LineShape7.Y1 = 52
        Me.LineShape7.Y2 = 52
        '
        'LineOutBox
        '
        Me.LineOutBox.Name = "LineOutBox"
        Me.LineOutBox.X1 = 447
        Me.LineOutBox.X2 = 447
        Me.LineOutBox.Y1 = 32
        Me.LineOutBox.Y2 = 72
        '
        'LineBottom
        '
        Me.LineBottom.Name = "LineBottom"
        Me.LineBottom.X1 = 18
        Me.LineBottom.X2 = 687
        Me.LineBottom.Y1 = 72
        Me.LineBottom.Y2 = 72
        '
        'LineInBox
        '
        Me.LineInBox.Name = "LineInBox"
        Me.LineInBox.X1 = 327
        Me.LineInBox.X2 = 327
        Me.LineInBox.Y1 = 32
        Me.LineInBox.Y2 = 71
        '
        'LineJenis
        '
        Me.LineJenis.Name = "LineJenis"
        Me.LineJenis.X1 = 49
        Me.LineJenis.X2 = 49
        Me.LineJenis.Y1 = 32
        Me.LineJenis.Y2 = 72
        '
        'LineNo
        '
        Me.LineNo.Name = "LineNo"
        Me.LineNo.X1 = 18
        Me.LineNo.X2 = 18
        Me.LineNo.Y1 = 32
        Me.LineNo.Y2 = 72
        '
        'LineShape1
        '
        Me.LineShape1.Name = "LineShape1"
        Me.LineShape1.X1 = 18
        Me.LineShape1.X2 = 687
        Me.LineShape1.Y1 = 32
        Me.LineShape1.Y2 = 32
        '
        'GroupHeader
        '
        Me.GroupHeader.Controls.Add(Me.Label7)
        Me.GroupHeader.Location = New System.Drawing.Point(4, 107)
        Me.GroupHeader.Name = "GroupHeader"
        Me.GroupHeader.Size = New System.Drawing.Size(498, 34)
        Me.GroupHeader.TabIndex = 29
        Me.GroupHeader.TabStop = False
        Me.GroupHeader.Text = "GroupHeader"
        '
        'Detail
        '
        Me.Detail.Controls.Add(Me.TextBox9)
        Me.Detail.Controls.Add(Me.TextBox10)
        Me.Detail.Controls.Add(Me.TextBox7)
        Me.Detail.Controls.Add(Me.TextBox8)
        Me.Detail.Controls.Add(Me.TextBox6)
        Me.Detail.Controls.Add(Me.TextBox5)
        Me.Detail.Controls.Add(Me.TextBox4)
        Me.Detail.Controls.Add(Me.TextBox3)
        Me.Detail.Controls.Add(Me.txtNo)
        Me.Detail.Controls.Add(Me.TextBox2)
        Me.Detail.Controls.Add(Me.TextBox1)
        Me.Detail.Location = New System.Drawing.Point(4, 147)
        Me.Detail.Name = "Detail"
        Me.Detail.Size = New System.Drawing.Size(720, 68)
        Me.Detail.TabIndex = 30
        Me.Detail.TabStop = False
        Me.Detail.Text = "Detail"
        '
        'TextBox9
        '
        Me.TextBox9.Location = New System.Drawing.Point(571, 16)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(42, 20)
        Me.TextBox9.TabIndex = 21
        Me.TextBox9.Tag = "Integer"
        Me.TextBox9.Text = "AkhirBox"
        Me.TextBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox9.Visible = False
        '
        'TextBox10
        '
        Me.TextBox10.Location = New System.Drawing.Point(621, 16)
        Me.TextBox10.Name = "TextBox10"
        Me.TextBox10.Size = New System.Drawing.Size(62, 20)
        Me.TextBox10.TabIndex = 20
        Me.TextBox10.Tag = "Double"
        Me.TextBox10.Text = "AkhirKg"
        Me.TextBox10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox10.Visible = False
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(451, 16)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(42, 20)
        Me.TextBox7.TabIndex = 19
        Me.TextBox7.Tag = "Integer"
        Me.TextBox7.Text = "OutBox"
        Me.TextBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox7.Visible = False
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(501, 16)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(62, 20)
        Me.TextBox8.TabIndex = 18
        Me.TextBox8.Tag = "Double"
        Me.TextBox8.Text = "OutKg"
        Me.TextBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox8.Visible = False
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(331, 16)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(42, 20)
        Me.TextBox6.TabIndex = 17
        Me.TextBox6.Tag = "Integer"
        Me.TextBox6.Text = "InBox"
        Me.TextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox6.Visible = False
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(300, 16)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(30, 20)
        Me.TextBox5.TabIndex = 15
        Me.TextBox5.Tag = "String"
        Me.TextBox5.Text = "Grade"
        Me.TextBox5.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(236, 16)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(54, 20)
        Me.TextBox4.TabIndex = 14
        Me.TextBox4.Tag = "String"
        Me.TextBox4.Text = "NoWarna"
        Me.TextBox4.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(192, 16)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(41, 20)
        Me.TextBox3.TabIndex = 13
        Me.TextBox3.Tag = "String"
        Me.TextBox3.Text = "Warna"
        Me.TextBox3.Visible = False
        '
        'txtNo
        '
        Me.txtNo.Location = New System.Drawing.Point(20, 16)
        Me.txtNo.Name = "txtNo"
        Me.txtNo.Size = New System.Drawing.Size(27, 20)
        Me.txtNo.TabIndex = 12
        Me.txtNo.Tag = "Header"
        Me.txtNo.Text = "No"
        Me.txtNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(56, 16)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(129, 20)
        Me.TextBox2.TabIndex = 11
        Me.TextBox2.Tag = "String"
        Me.TextBox2.Text = "KodeBarang"
        Me.TextBox2.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(381, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(62, 20)
        Me.TextBox1.TabIndex = 10
        Me.TextBox1.Tag = "Double"
        Me.TextBox1.Text = "InKg"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox1.Visible = False
        '
        'GroupFooter
        '
        Me.GroupFooter.Controls.Add(Me.Label1)
        Me.GroupFooter.Controls.Add(Me.takhirBox)
        Me.GroupFooter.Controls.Add(Me.toutBox)
        Me.GroupFooter.Controls.Add(Me.takhirKg)
        Me.GroupFooter.Controls.Add(Me.tinKg)
        Me.GroupFooter.Controls.Add(Me.tinBox)
        Me.GroupFooter.Controls.Add(Me.toutKg)
        Me.GroupFooter.Location = New System.Drawing.Point(4, 251)
        Me.GroupFooter.Name = "GroupFooter"
        Me.GroupFooter.Size = New System.Drawing.Size(724, 81)
        Me.GroupFooter.TabIndex = 31
        Me.GroupFooter.TabStop = False
        Me.GroupFooter.Text = "GroupFooter"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(264, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 15)
        Me.Label1.TabIndex = 25
        Me.Label1.Tag = "Header"
        Me.Label1.Text = "TOTAL"
        '
        'takhirBox
        '
        Me.takhirBox.Location = New System.Drawing.Point(571, 36)
        Me.takhirBox.Name = "takhirBox"
        Me.takhirBox.Size = New System.Drawing.Size(42, 20)
        Me.takhirBox.TabIndex = 27
        Me.takhirBox.Tag = "Header"
        Me.takhirBox.Text = "AkhirBox"
        Me.takhirBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.takhirBox.Visible = False
        '
        'toutBox
        '
        Me.toutBox.Location = New System.Drawing.Point(451, 36)
        Me.toutBox.Name = "toutBox"
        Me.toutBox.Size = New System.Drawing.Size(42, 20)
        Me.toutBox.TabIndex = 25
        Me.toutBox.Tag = "Header"
        Me.toutBox.Text = "OutBox"
        Me.toutBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.toutBox.Visible = False
        '
        'takhirKg
        '
        Me.takhirKg.Location = New System.Drawing.Point(621, 36)
        Me.takhirKg.Name = "takhirKg"
        Me.takhirKg.Size = New System.Drawing.Size(62, 20)
        Me.takhirKg.TabIndex = 26
        Me.takhirKg.Tag = "Header"
        Me.takhirKg.Text = "AkhirKg"
        Me.takhirKg.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.takhirKg.Visible = False
        '
        'tinKg
        '
        Me.tinKg.Location = New System.Drawing.Point(381, 36)
        Me.tinKg.Name = "tinKg"
        Me.tinKg.Size = New System.Drawing.Size(62, 20)
        Me.tinKg.TabIndex = 22
        Me.tinKg.Tag = "Header"
        Me.tinKg.Text = "InKg"
        Me.tinKg.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.tinKg.Visible = False
        '
        'tinBox
        '
        Me.tinBox.Location = New System.Drawing.Point(331, 36)
        Me.tinBox.Name = "tinBox"
        Me.tinBox.Size = New System.Drawing.Size(42, 20)
        Me.tinBox.TabIndex = 23
        Me.tinBox.Tag = "Header"
        Me.tinBox.Text = "InBox"
        Me.tinBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.tinBox.Visible = False
        '
        'toutKg
        '
        Me.toutKg.Location = New System.Drawing.Point(501, 36)
        Me.toutKg.Name = "toutKg"
        Me.toutKg.Size = New System.Drawing.Size(62, 20)
        Me.toutKg.TabIndex = 24
        Me.toutKg.Tag = "Header"
        Me.toutKg.Text = "OutKg"
        Me.toutKg.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.toutKg.Visible = False
        '
        'rptMutasi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 351)
        Me.Controls.Add(Me.GroupFooter)
        Me.Controls.Add(Me.Detail)
        Me.Controls.Add(Me.GroupHeader)
        Me.Controls.Add(Me.PageHeader)
        Me.Name = "rptMutasi"
        Me.Text = "rptMutasi"
        Me.PageHeader.ResumeLayout(False)
        Me.PageHeader.PerformLayout()
        Me.GroupHeader.ResumeLayout(False)
        Me.GroupHeader.PerformLayout()
        Me.Detail.ResumeLayout(False)
        Me.Detail.PerformLayout()
        Me.GroupFooter.ResumeLayout(False)
        Me.GroupFooter.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblHeader As System.Windows.Forms.Label
    Friend WithEvents dTgl As System.Windows.Forms.Label
    Friend WithEvents PageHeader As System.Windows.Forms.GroupBox
    Friend WithEvents GroupHeader As System.Windows.Forms.GroupBox
    Friend WithEvents Detail As System.Windows.Forms.GroupBox
    Friend WithEvents txtNo As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupFooter As System.Windows.Forms.GroupBox
    Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents lblKg As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblNo As System.Windows.Forms.Label
    Friend WithEvents lblJenisBarang As System.Windows.Forms.Label
    Friend WithEvents lblBoxAkhir As System.Windows.Forms.Label
    Friend WithEvents lblKgAkhir As System.Windows.Forms.Label
    Friend WithEvents lblBoxOut As System.Windows.Forms.Label
    Friend WithEvents lblKgOut As System.Windows.Forms.Label
    Friend WithEvents lblBox As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents LineOutBox As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineBottom As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineInBox As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineJenis As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineNo As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineShape1 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents LineOutKg As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineInKg As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineLast As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineAkhirBox As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineShape7 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LineAkhirKg As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents takhirBox As System.Windows.Forms.TextBox
    Friend WithEvents toutBox As System.Windows.Forms.TextBox
    Friend WithEvents takhirKg As System.Windows.Forms.TextBox
    Friend WithEvents tinKg As System.Windows.Forms.TextBox
    Friend WithEvents tinBox As System.Windows.Forms.TextBox
    Friend WithEvents toutKg As System.Windows.Forms.TextBox
End Class

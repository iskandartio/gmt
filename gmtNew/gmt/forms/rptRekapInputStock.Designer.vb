<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class rptRekapInputStock
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
        Me.lblHeader = New System.Windows.Forms.Label
        Me.dTgl = New System.Windows.Forms.Label
        Me.PageHeader = New System.Windows.Forms.GroupBox
        Me.GroupHeader = New System.Windows.Forms.GroupBox
        Me.Detail = New System.Windows.Forms.GroupBox
        Me.GroupFooter = New System.Windows.Forms.GroupBox
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.dTotal = New System.Windows.Forms.Label
        Me.PageHeader.SuspendLayout()
        Me.GroupHeader.SuspendLayout()
        Me.Detail.SuspendLayout()
        Me.GroupFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblHeader
        '
        Me.lblHeader.AutoSize = True
        Me.lblHeader.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeader.Location = New System.Drawing.Point(85, 28)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(249, 15)
        Me.lblHeader.TabIndex = 0
        Me.lblHeader.Tag = "Header"
        Me.lblHeader.Text = "LAPORAN HARIAN PRODUKSI PACKING"
        '
        'dTgl
        '
        Me.dTgl.AutoSize = True
        Me.dTgl.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dTgl.Location = New System.Drawing.Point(85, 43)
        Me.dTgl.Name = "dTgl"
        Me.dTgl.Size = New System.Drawing.Size(51, 15)
        Me.dTgl.TabIndex = 1
        Me.dTgl.Tag = "Header"
        Me.dTgl.Text = "Tanggal"
        '
        'PageHeader
        '
        Me.PageHeader.Controls.Add(Me.Label11)
        Me.PageHeader.Controls.Add(Me.Label10)
        Me.PageHeader.Controls.Add(Me.Label9)
        Me.PageHeader.Controls.Add(Me.Label8)
        Me.PageHeader.Controls.Add(Me.Label6)
        Me.PageHeader.Controls.Add(Me.Label5)
        Me.PageHeader.Controls.Add(Me.Label4)
        Me.PageHeader.Controls.Add(Me.Label3)
        Me.PageHeader.Controls.Add(Me.lblHeader)
        Me.PageHeader.Controls.Add(Me.dTgl)
        Me.PageHeader.Location = New System.Drawing.Point(4, 7)
        Me.PageHeader.Name = "PageHeader"
        Me.PageHeader.Size = New System.Drawing.Size(684, 100)
        Me.PageHeader.TabIndex = 28
        Me.PageHeader.TabStop = False
        Me.PageHeader.Text = "PageHeader"
        '
        'GroupHeader
        '
        Me.GroupHeader.Controls.Add(Me.Label1)
        Me.GroupHeader.Controls.Add(Me.Label7)
        Me.GroupHeader.Location = New System.Drawing.Point(4, 113)
        Me.GroupHeader.Name = "GroupHeader"
        Me.GroupHeader.Size = New System.Drawing.Size(498, 37)
        Me.GroupHeader.TabIndex = 29
        Me.GroupHeader.TabStop = False
        Me.GroupHeader.Text = "GroupHeader"
        '
        'Detail
        '
        Me.Detail.Controls.Add(Me.TextBox1)
        Me.Detail.Controls.Add(Me.TextBox6)
        Me.Detail.Controls.Add(Me.TextBox5)
        Me.Detail.Controls.Add(Me.TextBox4)
        Me.Detail.Controls.Add(Me.TextBox3)
        Me.Detail.Controls.Add(Me.TextBox2)
        Me.Detail.Location = New System.Drawing.Point(4, 156)
        Me.Detail.Name = "Detail"
        Me.Detail.Size = New System.Drawing.Size(684, 44)
        Me.Detail.TabIndex = 30
        Me.Detail.TabStop = False
        Me.Detail.Text = "Detail"
        '
        'GroupFooter
        '
        Me.GroupFooter.Controls.Add(Me.dTotal)
        Me.GroupFooter.Controls.Add(Me.Label2)
        Me.GroupFooter.Controls.Add(Me.Label17)
        Me.GroupFooter.Controls.Add(Me.Label15)
        Me.GroupFooter.Location = New System.Drawing.Point(4, 216)
        Me.GroupFooter.Name = "GroupFooter"
        Me.GroupFooter.Size = New System.Drawing.Size(684, 46)
        Me.GroupFooter.TabIndex = 31
        Me.GroupFooter.TabStop = False
        Me.GroupFooter.Text = "GroupFooter"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(50, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(74, 15)
        Me.Label7.TabIndex = 14
        Me.Label7.Tag = "Group"
        Me.Label7.Text = "KodeBarang"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(520, 16)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(76, 20)
        Me.TextBox2.TabIndex = 14
        Me.TextBox2.Tag = "Integer"
        Me.TextBox2.Text = "n1"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox2.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(602, 16)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(76, 20)
        Me.TextBox3.TabIndex = 15
        Me.TextBox3.Tag = "Double"
        Me.TextBox3.Text = "n2"
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox3.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(287, 16)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(65, 20)
        Me.TextBox4.TabIndex = 16
        Me.TextBox4.Tag = "String"
        Me.TextBox4.Text = "Tube"
        Me.TextBox4.Visible = False
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(357, 16)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(76, 20)
        Me.TextBox5.TabIndex = 17
        Me.TextBox5.Tag = "String"
        Me.TextBox5.Text = "Grade"
        Me.TextBox5.Visible = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(622, 16)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 15)
        Me.Label17.TabIndex = 34
        Me.Label17.Tag = "Header"
        Me.Label17.Text = "KgCount"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(536, 16)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(60, 15)
        Me.Label15.TabIndex = 32
        Me.Label15.Tag = "Header"
        Me.Label15.Text = "BoxCount"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(160, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(25, 15)
        Me.Label1.TabIndex = 15
        Me.Label1.Tag = "Group"
        Me.Label1.Text = "Lot"
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(439, 16)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(76, 20)
        Me.TextBox6.TabIndex = 18
        Me.TextBox6.Tag = "Integer"
        Me.TextBox6.Text = "nCones"
        Me.TextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox6.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(441, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 15)
        Me.Label2.TabIndex = 35
        Me.Label2.Tag = "Header"
        Me.Label2.Text = "ConesCount"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(216, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(65, 20)
        Me.TextBox1.TabIndex = 19
        Me.TextBox1.Tag = "String"
        Me.TextBox1.Text = "NoWarna"
        Me.TextBox1.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(50, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 15)
        Me.Label3.TabIndex = 2
        Me.Label3.Tag = "Header"
        Me.Label3.Text = "JENIS BARANG"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(160, 82)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(33, 15)
        Me.Label4.TabIndex = 3
        Me.Label4.Tag = "Header"
        Me.Label4.Text = "LOT"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(213, 82)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 15)
        Me.Label5.TabIndex = 4
        Me.Label5.Tag = "Header"
        Me.Label5.Text = "WARNA"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(284, 82)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 15)
        Me.Label6.TabIndex = 5
        Me.Label6.Tag = "Header"
        Me.Label6.Text = "TUBE"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(354, 82)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(50, 15)
        Me.Label8.TabIndex = 6
        Me.Label8.Tag = "Header"
        Me.Label8.Text = "GRADE"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(482, 82)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(33, 15)
        Me.Label9.TabIndex = 7
        Me.Label9.Tag = "Header"
        Me.Label9.Text = "CHS"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(561, 82)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(35, 15)
        Me.Label10.TabIndex = 8
        Me.Label10.Tag = "Header"
        Me.Label10.Text = "BOX"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(653, 82)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(25, 15)
        Me.Label11.TabIndex = 9
        Me.Label11.Tag = "Header"
        Me.Label11.Text = "KG"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dTotal
        '
        Me.dTotal.AutoSize = True
        Me.dTotal.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dTotal.Location = New System.Drawing.Point(324, 16)
        Me.dTotal.Name = "dTotal"
        Me.dTotal.Size = New System.Drawing.Size(50, 15)
        Me.dTotal.TabIndex = 36
        Me.dTotal.Tag = "Header"
        Me.dTotal.Text = "TOTAL"
        '
        'rptRekapInputStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(700, 317)
        Me.Controls.Add(Me.GroupFooter)
        Me.Controls.Add(Me.Detail)
        Me.Controls.Add(Me.GroupHeader)
        Me.Controls.Add(Me.PageHeader)
        Me.Name = "rptRekapInputStock"
        Me.Text = "rptRekapInputStock"
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
    Friend WithEvents lblHeader As System.Windows.Forms.Label
    Friend WithEvents dTgl As System.Windows.Forms.Label
    Friend WithEvents PageHeader As System.Windows.Forms.GroupBox
    Friend WithEvents GroupHeader As System.Windows.Forms.GroupBox
    Friend WithEvents Detail As System.Windows.Forms.GroupBox
    Friend WithEvents GroupFooter As System.Windows.Forms.GroupBox
    Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dTotal As System.Windows.Forms.Label

End Class

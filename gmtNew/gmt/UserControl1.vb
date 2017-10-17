Public Class UserControl1
   Inherits System.Windows.Forms.UserControl
   Implements Popup.IPopupUserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
   Friend WithEvents Button1 As System.Windows.Forms.Button
   Friend WithEvents Button2 As System.Windows.Forms.Button
   Friend WithEvents rbNorth As System.Windows.Forms.RadioButton
   Friend WithEvents rbWest As System.Windows.Forms.RadioButton
   Friend WithEvents rbSouth As System.Windows.Forms.RadioButton
   Friend WithEvents rbEast As System.Windows.Forms.RadioButton
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents RbNone As System.Windows.Forms.RadioButton
   Friend WithEvents Button3 As System.Windows.Forms.Button
   Friend WithEvents TrackBar1 As System.Windows.Forms.TrackBar
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.Button1 = New System.Windows.Forms.Button
      Me.Button2 = New System.Windows.Forms.Button
      Me.rbWest = New System.Windows.Forms.RadioButton
      Me.rbSouth = New System.Windows.Forms.RadioButton
      Me.rbNorth = New System.Windows.Forms.RadioButton
      Me.rbEast = New System.Windows.Forms.RadioButton
      Me.Label1 = New System.Windows.Forms.Label
      Me.Label2 = New System.Windows.Forms.Label
      Me.Label3 = New System.Windows.Forms.Label
      Me.Label4 = New System.Windows.Forms.Label
      Me.RbNone = New System.Windows.Forms.RadioButton
      Me.Button3 = New System.Windows.Forms.Button
      Me.TrackBar1 = New System.Windows.Forms.TrackBar
      CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'Button1
      '
      Me.Button1.Location = New System.Drawing.Point(8, 8)
      Me.Button1.Name = "Button1"
      Me.Button1.TabIndex = 6
      Me.Button1.Text = "This is"
      '
      'Button2
      '
      Me.Button2.Location = New System.Drawing.Point(88, 8)
      Me.Button2.Name = "Button2"
      Me.Button2.TabIndex = 7
      Me.Button2.Text = "a sample"
      '
      'rbWest
      '
      Me.rbWest.Location = New System.Drawing.Point(40, 120)
      Me.rbWest.Name = "rbWest"
      Me.rbWest.Size = New System.Drawing.Size(12, 12)
      Me.rbWest.TabIndex = 10
      '
      'rbSouth
      '
      Me.rbSouth.Location = New System.Drawing.Point(56, 136)
      Me.rbSouth.Name = "rbSouth"
      Me.rbSouth.Size = New System.Drawing.Size(12, 12)
      Me.rbSouth.TabIndex = 9
      '
      'rbNorth
      '
      Me.rbNorth.Checked = True
      Me.rbNorth.Location = New System.Drawing.Point(56, 104)
      Me.rbNorth.Name = "rbNorth"
      Me.rbNorth.Size = New System.Drawing.Size(12, 12)
      Me.rbNorth.TabIndex = 8
      Me.rbNorth.TabStop = True
      '
      'rbEast
      '
      Me.rbEast.Location = New System.Drawing.Point(72, 120)
      Me.rbEast.Name = "rbEast"
      Me.rbEast.Size = New System.Drawing.Size(12, 12)
      Me.rbEast.TabIndex = 11
      '
      'Label1
      '
      Me.Label1.Location = New System.Drawing.Point(24, 72)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(63, 24)
      Me.Label1.TabIndex = 12
      Me.Label1.Text = "North"
      Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
      '
      'Label2
      '
      Me.Label2.Location = New System.Drawing.Point(32, 152)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(56, 16)
      Me.Label2.TabIndex = 13
      Me.Label2.Text = "South"
      Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
      '
      'Label3
      '
      Me.Label3.Location = New System.Drawing.Point(88, 104)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(48, 40)
      Me.Label3.TabIndex = 14
      Me.Label3.Text = "East"
      Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
      '
      'Label4
      '
      Me.Label4.Location = New System.Drawing.Point(0, 104)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(32, 40)
      Me.Label4.TabIndex = 15
      Me.Label4.Text = "West"
      Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'RbNone
      '
      Me.RbNone.Location = New System.Drawing.Point(160, 144)
      Me.RbNone.Name = "RbNone"
      Me.RbNone.Size = New System.Drawing.Size(88, 16)
      Me.RbNone.TabIndex = 11
      Me.RbNone.Text = "Not decided"
      '
      'Button3
      '
      Me.Button3.Location = New System.Drawing.Point(168, 8)
      Me.Button3.Name = "Button3"
      Me.Button3.TabIndex = 16
      Me.Button3.Text = "usercontrol"
      '
      'TrackBar1
      '
      Me.TrackBar1.Location = New System.Drawing.Point(0, 40)
      Me.TrackBar1.Name = "TrackBar1"
      Me.TrackBar1.Size = New System.Drawing.Size(248, 45)
      Me.TrackBar1.TabIndex = 17
      '
      'UserControl1
      '
      Me.BackColor = System.Drawing.SystemColors.Control
      Me.Controls.Add(Me.TrackBar1)
      Me.Controls.Add(Me.Button3)
      Me.Controls.Add(Me.Label4)
      Me.Controls.Add(Me.Label3)
      Me.Controls.Add(Me.Label2)
      Me.Controls.Add(Me.Label1)
      Me.Controls.Add(Me.rbWest)
      Me.Controls.Add(Me.rbSouth)
      Me.Controls.Add(Me.rbNorth)
      Me.Controls.Add(Me.rbEast)
      Me.Controls.Add(Me.Button2)
      Me.Controls.Add(Me.Button1)
      Me.Controls.Add(Me.RbNone)
      Me.Name = "UserControl1"
      Me.Size = New System.Drawing.Size(248, 176)
      CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region

   Private Sub UserControl1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

   End Sub

   Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
      Me.FindForm.Invalidate()
   End Sub

   Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

   End Sub

   Public Function AcceptPopupClosing() As Boolean Implements Popup.IPopupUserControl.AcceptPopupClosing
      If RbNone.Checked Then
         CustomTooltip.ShowTooltip(Label3, popup.ePlacement.Right, "You must select a direction" & vbCrLf & "The popup won't disappear until you make a choice")
         Return False
      Else
         Return True
      End If
   End Function

   Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll

   End Sub
End Class

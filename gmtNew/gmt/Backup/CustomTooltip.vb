Public Class CustomTooltip
   Inherits System.Windows.Forms.UserControl

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
   Private WithEvents Label1 As System.Windows.Forms.Label
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.Label1 = New System.Windows.Forms.Label
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.Location = New System.Drawing.Point(8, 8)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(56, 32)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "..."
      '
      'CustomTooltip
      '
      Me.BackColor = System.Drawing.SystemColors.Info
      Me.Controls.Add(Me.Label1)
      Me.Name = "CustomTooltip"
      Me.Size = New System.Drawing.Size(72, 48)
      Me.ResumeLayout(False)

   End Sub

#End Region

   Shared instance As New CustomTooltip

   Shared Sub ShowTooltip(ByVal parent As Control, ByVal placement As Popup.ePlacement, ByVal text As String)
      With instance
         Dim g As Graphics = .CreateGraphics()
         Dim w As Integer = 100
         Dim ls As SizeF
         Do
            ls = g.MeasureString(text, .Font, New SizeF(w, Integer.MaxValue))
            If ls.Height < w Then
               Exit Do
            Else
               w *= 1.414
            End If
         Loop
         g.Dispose()
         .Label1.Width = CInt(ls.Width) + 1
         .Label1.Height = CInt(ls.Height) + 1
         .Label1.Text = text
         .Size = New Size(.Label1.Width + 16, .Label1.Height + 16)
         Dim popup As New Popup(instance, parent)
         popup.HorizontalPlacement = placement
         popup.AnimationSpeed = 0
         popup.ShowShadow = False
         popup.Show()
      End With
   End Sub
End Class

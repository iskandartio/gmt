Imports System.Windows.Forms

Public Class ModComboBox
    Inherits System.Windows.Forms.ComboBox

    

    Private _Enabled As Boolean = True
    Private _enabledBackcolor As Color = MyBase.BackColor
    Private _disabledBackColor As Color = Color.Gainsboro
    Private m_DisplayMember As String()
    Private m_Width As String()

    Public Shadows Property Enabled() As Boolean
        Get
            Return _Enabled
        End Get
        Set(ByVal Value As Boolean)
            If _Enabled <> Value Then
                _Enabled = Value
                OnEnabledChanged(New EventArgs)
            End If
        End Set
    End Property

    Public Property DisabledBackColor() As Color
        Get
            Return _disabledBackColor
        End Get
        Set(ByVal value As Color)
            _disabledBackColor = value
            If Not _Enabled Then
                MyBase.BackColor = _disabledBackColor
            End If
        End Set
    End Property

    Public Shadows Property BackColor() As Color
        Get
            Return _enabledBackcolor
        End Get
        Set(ByVal value As Color)
            If _enabledBackcolor <> value Then
                _enabledBackcolor = value
                MyBase.BackColor = _enabledBackcolor
            End If
        End Set
    End Property

    Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
        ToggleEnabled()
        MyBase.OnEnabledChanged(e)
    End Sub

    Public Overrides Function PreProcessMessage(ByRef msg As Message) As Boolean
        If Not _Enabled Then
            If msg.Msg = &H100 Then
                Dim key As Int32 = msg.WParam.ToInt32
                If key <> Keys.Tab OrElse _
                   key <> Keys.Left OrElse _
                   key <> Keys.Right Then
                    Return True
                End If
            End If
        End If
        Return MyBase.PreProcessMessage(msg)
    End Function

    Protected Overrides Sub WndProc(ByRef m As Message)
        If Not _Enabled Then
            If m.Msg = &H201 OrElse m.Msg = &H203 Then
                Return
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Protected Overrides Sub OnParentEnabledChanged(ByVal e As System.EventArgs)
        _Enabled = MyBase.Parent.Enabled
        If _Enabled Then
            MyBase.OnParentEnabledChanged(e)
        Else
            ToggleEnabled()
        End If
    End Sub
    Private Sub ToggleEnabled()
        MyBase.TabStop = _Enabled
        If Not _Enabled Then
            MyBase.ContextMenuStrip = New ContextMenuStrip
            MyBase.BackColor = _disabledBackColor
        Else
            MyBase.ContextMenuStrip = Nothing
            MyBase.BackColor = _enabledBackcolor
        End If
        Me.Focus()
    End Sub

    Public Sub LoadMe(ByVal dt As DataTable, ByVal tValueMember As String, ByVal tDisplayMember As String, Optional ByVal tWidth As String = "")
        m_DisplayMember = Split(tDisplayMember, "@")
        If m_DisplayMember.GetUpperBound(0) = 0 Then
            Me.DataSource = dt
            Me.ValueMember = tValueMember
            Me.DisplayMember = tDisplayMember
            Me.DrawMode = Windows.Forms.DrawMode.
        Else

            m_Width = Split(tWidth, "@")
            Me.DataSource = dt
            Me.ValueMember = tValueMember
            Me.DrawMode = Windows.Forms.DrawMode.OwnerDrawFixed
            Me.FormatString = "@"
        End If
    End Sub

    Private Sub ModComboBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles Me.DrawItem
        If e.Index = -1 Then Exit Sub
        Dim brush As Brush
        Dim brush2 As Brush
        e.DrawBackground()
        If e.BackColor.Name = "Highlight" Then
            brush = Brushes.White
        Else
            brush = Brushes.Black
        End If

        Dim w As Integer
        Dim w2 As Integer
        w = e.Bounds.X
        For i As Integer = 0 To m_DisplayMember.GetUpperBound(0)
            If e.BackColor.Name <> "Highlight" Then
                If i > m_Width.GetUpperBound(0) Then
                    w2 = 50
                Else
                    w2 = Convert.ToInt32(m_Width(i))
                End If
                If i = 0 Then brush2 = Brushes.LightBlue
                If i = 1 Then brush2 = Brushes.LightCoral
                If i = 2 Then brush2 = Brushes.LightCyan
                If i = 3 Then brush2 = Brushes.LightGoldenrodYellow
                If i = 4 Then brush2 = Brushes.LightGray
                If i < m_DisplayMember.GetUpperBound(0) Then e.Graphics.FillRectangle(brush2, w, e.Bounds.Y, w2, e.Bounds.Height)
            End If
            e.Graphics.DrawString(sender.Items(e.Index)(m_DisplayMember(i)), sender.Font, brush, w, e.Bounds.Y)
            If i > m_Width.GetUpperBound(0) Then
                w = w + 50
            Else
                w = w + m_Width(i)
            End If

        Next
    End Sub
End Class
Public Class ModDTPicker
    Inherits DateTimePicker

    <System.ComponentModel.Description("Returns a window handler."), _
    Runtime.InteropServices.DllImport("user32.dll", _
    SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Auto, ExactSpelling:=True, _
    CallingConvention:=Runtime.InteropServices.CallingConvention.StdCall)> _
    Public Shared Function GetWindowDC(ByVal hWnd As IntPtr) As IntPtr
    End Function

    <System.ComponentModel.Description("frees a window handler."), _
    Runtime.InteropServices.DllImport("user32.dll", _
    SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Auto, ExactSpelling:=True, _
    CallingConvention:=Runtime.InteropServices.CallingConvention.StdCall)> _
    Public Shared Function ReleaseDC(ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Integer
    End Function

    Const WM_PAINT = &HF

#Region "WM_PAINT"
    Private mDisabledBackColor As Color

    Public Property DisabledBackColor() As Color
        Get
            Return mDisabledBackColor
        End Get
        Set(ByVal value As Color)
            mDisabledBackColor = value
        End Set
    End Property
    Private mDisabledForeColor As Color
    Public Property DisabledForeColor() As Color
        Get
            Return mDisabledForeColor
        End Get
        Set(ByVal value As Color)
            mDisabledForeColor = value
        End Set
    End Property
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_PAINT
                MyBase.WndProc(m)

                If Me.Enabled = False Then
                    Dim hDC As IntPtr = GetWindowDC([Handle])
                    Dim gdc As Drawing.Graphics = Drawing.Graphics.FromHdc(hDC)
                    Dim drawBrush As SolidBrush = New SolidBrush(Me.DisabledForeColor)
                    Dim drawBrush2 As SolidBrush = New SolidBrush(Me.DisabledBackColor)
                    gdc.FillRectangle(drawBrush2, 1.6F, 2.0F, Me.Width - 3.2F, Me.Height - 4.0F)
                    gdc.DrawString(Text, Font, drawBrush, 1.6F, 2.5F)

                    gdc.Dispose()
                    gdc = Nothing
                    ReleaseDC(m.HWnd, hDC)
                End If
            Case Else
                MyBase.WndProc(m)
        End Select
    End Sub

#End Region

    Public Sub New()
        MyBase.New()
        Me.DisabledBackColor = Color.Gainsboro
        Me.DisabledForeColor = Color.Black
        SetStyle(Windows.Forms.ControlStyles.DoubleBuffer, True)
    End Sub
End Class
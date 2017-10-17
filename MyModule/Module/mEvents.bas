Attribute VB_Name = "mEvents"
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private mProc As Long
Dim mf As Form

Sub AttachMessage(ByVal hwnd As Long, ByVal tForm As Form)
    Set mf = tForm
    mProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    WindowProc = CallWindowProc(mProc, hwnd, iMsg, wParam, ByVal lParam)
    mf.EventModule iMsg, wParam, lParam, hwnd
End Function

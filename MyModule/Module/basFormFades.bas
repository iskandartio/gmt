Attribute VB_Name = "basFormFades"
Option Explicit

Sub XFormBlueFade(vForm As Form)
    Dim intLoop As Integer
    Dim frm     As Form
    
    Set frm = vForm
    
    On Error Resume Next
    
    With vForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = 256
        .AutoRedraw = True
        .ClipControls = False
    End With



    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B 'Draw boxes With specified color of loop
    Next intLoop
    
    Set vForm = frm
End Sub


Sub XFormFireFade(vForm As Form)
    'This code works best when called in the
    '     paint event
    Dim intLoop As Integer
    Dim frm     As Form
    
    Set frm = vForm
    
    On Error Resume Next

    With vForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = 256
        .AutoRedraw = True
        .ClipControls = False
    End With


    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B 'Draw boxes With specified color of loop
    Next intLoop
    
    Set vForm = frm
End Sub


Sub XFormGreenFade(vForm As Form)
    
    Dim intLoop As Integer
    Dim frm     As Form
    
    Set frm = vForm
    
    On Error Resume Next
    
    With vForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = 256
        .AutoRedraw = True
        .ClipControls = False
    End With



    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B 'Draw boxes With specified color of loop
    Next intLoop
    
    Set vForm = frm
End Sub


Sub XFormIceFade(vForm As Form)
    Dim intLoop As Integer
    Dim frm     As Form
    
    Set frm = vForm
    
    On Error Resume Next
    
    With vForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = 256
        .AutoRedraw = True
        .ClipControls = False
    End With



    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B 'Draw boxes With specified color of loop
    Next intLoop
    
    Set vForm = frm
End Sub


Sub XFormPurpleFade(vForm As Form)
    Dim intLoop As Integer
    Dim frm     As Form
    
    Set frm = vForm
    
    On Error Resume Next
    
    With vForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = 256
        .AutoRedraw = True
        .ClipControls = False
    End With



    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(25, 0, 100 - intLoop), B 'Draw boxes With specified color of loop
    Next intLoop
    
    Set vForm = frm
End Sub


Sub XFormRedFade(vForm As Form)
    
    Dim intLoop As Integer
    Dim frm     As Form
    
    Set frm = vForm
    
    On Error Resume Next
    
    With vForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = 256
        .AutoRedraw = True
        .ClipControls = False
    End With



    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B 'Draw boxes With specified color of loop
    Next intLoop
    
    Set vForm = frm
End Sub


Sub XFormSilverFade(vForm As Form)
    Dim intLoop As Integer
    Dim frm     As Form
    
    Set frm = vForm
    
    On Error Resume Next
    
    With vForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = 256
        .AutoRedraw = True
        .ClipControls = False
    End With


    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B 'Draw boxes With specified color of loop
    Next intLoop
    
    Set vForm = frm
End Sub


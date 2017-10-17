VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Vals() As Variant
Dim n As Integer
Private Sub Class_Initialize()
    n = 0
    ReDim Vals(10)
End Sub
Public Sub Add(ByVal val As Variant)
    If IsObject(val) Then
        Set Vals(n) = val
    Else
        Vals(n) = val
    End If
    n = n + 1
    If n > UBound(Vals) Then
        ReDim Preserve Vals(UBound(Vals) + 10)
    End If
End Sub

Public Property Get Prop(ByVal idx As Long) As Variant
Attribute Prop.VB_UserMemId = 0
    If IsObject(Vals(idx)) Then
        Set Prop = Vals(idx)
    Else
        Prop = Vals(idx)
    End If
End Property

Public Property Let Prop(ByVal idx As Long, ByVal tVal As Variant)
    Vals(idx) = tVal
End Property

Public Function count() As Long
    count = n
End Function

Public Sub PushArray(ByVal tArr As ArrayList)
Dim i As Long
    ReDim Preserve Vals(UBound(Vals) + tArr.count)
    For i = 0 To tArr.count - 1
        Vals(n) = tArr(i)
        n = n + 1
    Next
End Sub

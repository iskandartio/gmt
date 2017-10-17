VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Vals() As Variant
Dim Props() As String
Dim mIndex() As String
Dim mIndex2() As Long
Dim num As Long
Dim IndexIns As Long

Private Sub Class_Initialize()
    ReDim Vals(10)
    ReDim Props(10)
    ReDim mIndex(10)
    ReDim mIndex2(10)
    num = 0
End Sub

Function GetVal(ByVal tIndex As Long)
    GetVal = Vals(tIndex)
End Function

Function GetNum() As String
    GetNum = num
End Function

Sub Clear()
    Class_Initialize
End Sub

Public Property Let Prop(ByVal tProp As String, ByVal tVal As Variant)
Dim i As Long
    i = FindIndex(tProp)
    If i = -1 Then
        InsertRow IndexIns, tProp
        i = num - 1
    End If
    Props(i) = tProp
    If IsObject(tVal) Then
    
        Set Vals(i) = tVal
    Else
        Vals(i) = tVal
    End If
End Property

Public Property Get Prop(ByVal tProp As String) As Variant
Attribute Prop.VB_UserMemId = 0
'On Error Resume Next
    Dim idx As Long
    idx = FindIndex(tProp)
    If idx = -1 Then Exit Property
    If IsObject(Vals(idx)) Then
        Set Prop = Vals(idx)
    Else
        Prop = Vals(idx)
    End If
End Property

Public Function GetPropName(ByVal tIndex As Long)
    GetPropName = Props(tIndex)
End Function

Private Function FindIndex(ByVal tVal As String) As Long
Dim l As Long
Dim h As Long
Dim m As Long
Dim v As Integer
    FindIndex = -1
    l = 0
    If num > 0 Then
        h = num - 1
        While l <= h
            m = (l + h) \ 2
            v = StrComp(mIndex(m), tVal)
            If v = 0 Then
                FindIndex = mIndex2(m)
                Exit Function
            ElseIf v = 1 Then
                h = m - 1
            Else
                l = m + 1
            End If
        Wend
    End If
    IndexIns = l
End Function

Private Sub InsertRow(ByVal tIndex As Long, ByVal tVal As String)
    num = num + 1
    If num > UBound(mIndex) Then
        ReDim Preserve mIndex(UBound(mIndex) + 10)
        ReDim Preserve mIndex2(UBound(mIndex2) + 10)
        ReDim Preserve Vals(UBound(Vals) + 10)
        ReDim Preserve Props(UBound(Props) + 10)
    End If
    If tIndex = num - 1 Then
        mIndex(tIndex) = tVal
        mIndex2(tIndex) = num - 1
        Exit Sub
    End If
    Dim a() As String
    Dim b() As String
    ReDim a(UBound(mIndex))
    ReDim b(UBound(mIndex))
    Dim j As Long
    Dim i As Long
    j = 0
    For i = tIndex To num - 2
        a(j) = mIndex(i)
        b(j) = mIndex2(i)
        j = j + 1
    Next
    mIndex(tIndex) = tVal
    mIndex2(tIndex) = num - 1
    Dim k As Long
    k = tIndex + 1
    For i = 0 To j - 1
        mIndex(k) = a(i)
        mIndex2(k) = b(i)
        k = k + 1
    Next
End Sub

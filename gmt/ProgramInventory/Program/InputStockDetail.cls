VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputStockDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public kg As Double
Public IDTrans As Long
Public NoUrut As Long
Public Sub Init(tNoUrut As Long, tKg As Double, tIDTrans As Long)
    NoUrut = tNoUrut
    kg = tKg
    IDTrans = tIDTrans
End Sub

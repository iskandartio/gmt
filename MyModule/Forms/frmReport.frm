VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   5295
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrinterSetup 
      Caption         =   "Printer Setup..."
      Height          =   240
      Left            =   1125
      TabIndex        =   1
      Top             =   450
      Width           =   1590
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   5640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      lastProp        =   600
      _cx             =   11695
      _cy             =   9948
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mApp As New CRAXDRT.Application
Dim mRpt As CRAXDRT.Report
Dim mAfterPrint As String

Private Sub cmdPrinterSetup_Click()
    mRpt.PrinterSetup Me.hwnd
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    If mAfterPrint = "" Then Exit Sub
    GDBConn.Execute mAfterPrint
End Sub

Private Sub Form_Resize()
    CRViewer1.Width = ScaleWidth
    CRViewer1.Height = ScaleHeight
End Sub

Function MakeDateTimeCR(ByVal tDate As Date) As String
Dim d As Byte
Dim m As Byte
Dim Y As Integer
    d = Day(tDate)
    m = Month(tDate)
    Y = Year(tDate)
    MakeDateTimeCR = "DateTime(" & Y & "," & m & "," & d & ",0,0,0)"
End Function

Sub LoadMe(ByVal tRptName As String, Optional ByVal tParams As String = "", Optional ByVal tRecordFormula As String = "", Optional ByVal tSQL As String = "", Optional ByVal tSQLTable As String = "", Optional ByVal tVoucher As Boolean = False, Optional ByVal tAfterPrint As String = "", Optional ByVal tWidth As Single = 0, Optional ByVal tHeight As Single = 0)
On Error Resume Next
    mAfterPrint = tAfterPrint
    Msg = MsgBox("Langsung Print?", vbYesNoCancel + vbQuestion)
    If Msg = vbCancel Then Exit Sub
    Set mRpt = Nothing
    Set mRpt = mApp.OpenReport(App.Path & "\" & tRptName)
    For i = 1 To mRpt.Database.Tables.Count
        mRpt.Database.Tables(i).ConnectBufferString = pConnectionString
        If tSQLTable <> "" Then
            If mRpt.Database.Tables(i).Name = tSQLTable Then
                Dim rs As New ADODB.Recordset
                rs.Open tSQL, GDBConn, adOpenKeyset, adLockOptimistic
                If rs.EOF Then
                    MsgBox "No Data"
                    Unload Me
                    Exit Sub
                End If
                mRpt.Database.Tables(i).SetDataSource rs
            End If
        End If
    Next
    mRpt.DiscardSavedData
Dim b() As String
    b = Split(tParams, "@")
    For i = 1 To mRpt.ParameterFields.Count
        If mRpt.ParameterFields(i).ValueType = crDateTimeField Then
            mRpt.ParameterFields(i).AddCurrentValue CDate(b(i - 1))
        ElseIf mRpt.ParameterFields(i).ValueType = crNumberField Then
            mRpt.ParameterFields(i).AddCurrentValue CDbl(b(i - 1))
        Else
            mRpt.ParameterFields(i).AddCurrentValue b(i - 1)
        End If
    Next
    mRpt.RecordSelectionFormula = tRecordFormula
    
    If tVoucher Then
        If tWidth = 0 Then
            mRpt.SetUserPaperSize 1400, 2200
        Else
            mRpt.SetUserPaperSize tHeight, tWidth
        End If
    End If
    Icon = MDiFrm.Icon
    If Msg = vbNo Then
        CRViewer1.ReportSource = mRpt
        CRViewer1.ViewReport
        While CRViewer1.IsBusy
            DoEvents
        Wend
        Show vbModal
    Else
        mRpt.PrintOut
        Unload Me
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mRpt = Nothing
End Sub

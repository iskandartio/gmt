VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form FormReport 
   Caption         =   "REPORT"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   Icon            =   "FormReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5925
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      lastProp        =   500
      _cx             =   5741
      _cy             =   4260
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton fHideDetail 
      Caption         =   "Command1"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "FormReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSQL As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub CRViewer91_OnContextMenu(ByVal ObjectDescription As Variant, ByVal x As Long, ByVal y As Long, UseDefault As Boolean)
    UseDefault = False
End Sub

Private Sub CRViewer91_OnReportSourceError(ByVal errorMsg As String, ByVal errorCode As Long, UseDefault As Boolean)
    UseDefault = False
    Rpt.BottomMargin = 100
End Sub

Private Sub CRViewer91_PrintButtonClicked(UseDefault As Boolean)
    Rpt.PrinterSetup hwnd
    If mSQL <> "" Then ExecMe mSQL
End Sub

Private Sub fHideDetail_Click()
On Error Resume Next
    If Rpt.Sections("DetailSection1").Suppress Then a = False Else a = True
    Rpt.Sections("DetailSection1").Suppress = a
    For i = 1 To Rpt.Sections.Count - 1
        For j = 1 To Rpt.Sections(i).ReportObjects.Count - 1
            If Rpt.Sections(i).ReportObjects(j).Kind = crSubreportObject Then
                Dim rpt2 As New CRAXDRT.Report
                Set rpt2 = Rpt.OpenSubreport(Rpt.Sections(i).ReportObjects(j).SubreportName)
                rpt2.Sections("DetailSection1").Suppress = a
            End If
        Next
    Next
    If CRViewer91.ViewCount = 1 Then
        CRViewer91.ViewReport
        Exit Sub
    End If
    Dim ViewName() As String
    ReDim ViewName(CRViewer91.ViewCount - 2)
    For i = 0 To UBound(ViewName)
        ViewName(i) = CRViewer91.GetViewPath(i + 2)(0)
    Next
    activeindex = CRViewer91.ActiveViewIndex
    CRViewer91.ViewReport
    While CRViewer91.IsBusy
        DoEvents
    Wend
    For i = 0 To UBound(ViewName)
        CRViewer91.AddView ViewName(i)
        While CRViewer91.IsBusy
        DoEvents
        Wend
    Next
    CRViewer91.ActivateView activeindex
End Sub

Private Sub Form_Resize()
On Error Resume Next
    CRViewer91.Width = ScaleWidth
    CRViewer91.Height = ScaleHeight
End Sub

Sub InitReport(ByVal tReportFile As String)
On Error Resume Next
    MousePointer = vbHourglass
    Set Rpt = Nothing
    Set Rpt = Crystal.OpenReport(pReportDir & tReportFile)
    For i = 1 To Rpt.Database.Tables.Count
        Rpt.Database.Tables(i).ConnectBufferString = pConnectionString
    Next
    Rpt.DiscardSavedData
    For i = 1 To Rpt.Sections.Count - 1
        For j = 1 To Rpt.Sections(i).ReportObjects.Count - 1
            If Rpt.Sections(i).ReportObjects(j).Kind = crSubreportObject Then
                Dim rpt2 As New CRAXDRT.Report
                Set rpt2 = Rpt.OpenSubreport(Rpt.Sections(i).ReportObjects(j).SubreportName)
                For k = 1 To rpt2.Database.Tables.Count
                    rpt2.Database.Tables(k).ConnectBufferString = pConnectionString
                Next
            End If
        Next
    Next
    MousePointer = vbDefault
End Sub

Sub LoadMe(ByVal ReportFile As String, Optional ByVal tParams As String = "", Optional ByVal tSQLAfterPrint As String = "", Optional ByVal tSQLData As String = "", Optional ByVal tSelectionFormula As String = "")
On Error Resume Next
    c = MsgBox("Langsung Print dengan " & Printer.DeviceName & "?", vbYesNoCancel)
    If c = vbCancel Then Exit Sub
    InitReport ReportFile
    If tSQLData <> "" Then
        query tSQLData
        Rpt.Database.SetDataSource RS
    End If
    If tParams <> "" Then
        Dim a() As String
        a = Split(tParams, "@")
        For i = 0 To UBound(a)
            Dim b As Variant
            If IsNumeric(a(i)) Then b = CLng(a(i)) Else b = a(i)
            Rpt.ParameterFields(i + 1).AddCurrentValue b
        Next
    End If
    Rpt.RecordSelectionFormula = Rpt.RecordSelectionFormula & tSelectionFormula
    Rpt.LeftMargin = pLeftMargin
    Rpt.TopMargin = pTopMargin
    If Rpt.BottomMargin > 6500 Then
        Rpt.SetUserPaperSize 1400, 2200
    End If
    mSQL = tSQLAfterPrint
    If c = vbYes Then
        Rpt.PrintOutEx False
        If mSQL <> "" Then ExecMe mSQL
        Unload Me
    Else
        Show
        CRViewer91.ReportSource = Rpt
        CRViewer91.ViewReport
    End If
End Sub

Sub LoadData(ByVal ReportFile As String, tTableIndex As Integer, tData As CrystalDataObject.CrystalComObject, Optional ByVal tParams As String, Optional ByVal tFormula As String)
    c = MsgBox("Langsung Print dengan " & Printer.DeviceName & "?", vbYesNoCancel)
    If c = vbCancel Then Exit Sub
    InitReport ReportFile
    Rpt.Database.Tables(tTableIndex).SetDataSource tData
    If tParams <> "" Then
        Dim a() As String
        a = Split(tParams, "@")
        For i = 0 To UBound(a)
            Dim b As Variant
            If IsNumeric(a(i)) Then b = CLng(a(i)) Else b = a(i)
            Rpt.ParameterFields(i + 1).AddCurrentValue b
        Next
    End If
    Rpt.RecordSelectionFormula = Rpt.RecordSelectionFormula & tFormula
    If Rpt.BottomMargin > 6500 Then
        Rpt.SetUserPaperSize 1400, 2200
    End If
    If c = vbYes Then
        Rpt.PrintOutEx False
        If mSQL <> "" Then ExecMe mSQL
        Unload Me
    Else
        Show
        CRViewer91.ReportSource = Rpt
        CRViewer91.ViewReport
        CRViewer91.Zoom 90
        While CRViewer91.IsBusy
            DoEvents
        Wend
    End If
End Sub

VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormPenomoran 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PENOMORAN"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Tag             =   "39"
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama Form"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Penomoran"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3757"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3678"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Named:id=29:Normal"
      _StyleDefs(39)  =   ":id=29,.parent=0"
      _StyleDefs(40)  =   "Named:id=30:Heading"
      _StyleDefs(41)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=30,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=31:Footing"
      _StyleDefs(44)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=32:Selected"
      _StyleDefs(46)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=33:Caption"
      _StyleDefs(48)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(49)  =   "Named:id=34:HighlightRow"
      _StyleDefs(50)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(51)  =   "Named:id=35:EvenRow"
      _StyleDefs(52)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=36:OddRow"
      _StyleDefs(54)  =   ":id=36,.parent=29"
      _StyleDefs(55)  =   "Named:id=39:RecordSelector"
      _StyleDefs(56)  =   ":id=39,.parent=30"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=29"
   End
End
Attribute VB_Name = "FormPenomoran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
    Caption = Caption & "---" & pTipe

    Set TDBGrid1.Array = x
    DoQuery
err:
End Sub

Private Sub DoQuery()
    a = "select * from m_penomoran~"
    query a
    x.LoadRows RS.GetRows
    TDBGrid1.Rebind
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left * 2
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TDBGrid1_BeforeUpdate False
End Sub

Private Sub TDBGrid1_BeforeUpdate(Cancel As Integer)
On Error GoTo err
    a = "update m_penomoran" & pTipe & " set Penomoran='" & TDBGrid1.Columns("Penomoran").Value & "' where NamaForm='" & TDBGrid1.Columns("Nama Form").Value & "'"
    b = ExecMe(a)
    a = "select NamaForm, Penomoran from m_Penomoran~ order by NamaForm"
    query a
    For i = 0 To RS.RecordCount - 1
        If RS.Fields("NamaForm").Value = "SC" Then
            pNomorSC = RS.Fields("Penomoran").Value
        ElseIf RS.Fields("NamaForm").Value = "SPP" Then
            pNomorSPP = RS.Fields("Penomoran").Value
        ElseIf RS.Fields("NamaForm").Value = "SJ" Then
            pNomorSJ = RS.Fields("Penomoran").Value
        ElseIf RS.Fields("NamaForm").Value = "KW" Then
            pNomorKW = RS.Fields("Penomoran").Value
        ElseIf RS.Fields("NamaForm").Value = "NI" Then
            pNomorNI = RS.Fields("Penomoran").Value
        ElseIf RS.Fields("NamaForm").Value = "NR" Then
            pNomorNR = RS.Fields("Penomoran").Value
        End If
        RS.MoveNext
    Next
    Exit Sub
err:
    Cancel = True
End Sub






VERSION 5.00
Object = "{5B6E0E90-AB64-4D5D-AC5E-5DC35FA1D835}#1.0#0"; "USRTEXT.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form FormInputMaster 
   Caption         =   "MASTER"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin UsrText.IText Texts 
      Height          =   270
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4048
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
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
   Begin UsrText.IText fFind 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labelxxx 
      Caption         =   "&FIND"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Labels 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FormInputMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSQLString As String
Dim mFields() As String
Dim mTypes() As String
Dim mHeaders() As String
Dim mWidths() As String
Dim mOrder As String
Dim mRowCount As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Sub LoadMe(ByVal tSQLString As String, ByVal tFilterHeaders As String, ByVal tFilterFields As String, tFilterWidths As String, ByVal tFilterTypes As String, ByVal tHeaders As String, ByVal tFields As String, ByVal tWidths As String, ByVal tTypes As String, ByVal tOrder As String)
    mSQLString = tSQLString
    a = Split(tFilterHeaders, "@")
    b = Split(tFilterFields, "@")
    c = Split(tFilterWidths, "@")
    d = Split(tFilterTypes, "@")
    Labels(0) = a(0)
    Labels(0).Width = c(0)
    Texts(0).Width = c(0)
    Labels(0).Tag = b(0)
    Texts(0).Tag = d(0)
    For i = 1 To UBound(a)
        Load Labels(i)
        Load Texts(i)
        Labels(i).Visible = True
        Labels(i).Left = Labels(i - 1).Left + Labels(i - 1).Width + 100
        Texts(i).Visible = True
        Texts(i).Left = Labels(i).Left
        Labels(i) = a(i)
        Labels(i).Width = c(i)
        Texts(i).Width = c(i)
        Labels(i).Tag = b(i)
        Texts(i).Tag = d(i)
    Next
    mHeaders = Split(tHeaders, "@")
    mWidths = Split(tWidths, "@")
    mFields = Split(tFields, "@")
    mTypes = Split(tTypes, "@")
    mOrder = tOrder
    RefreshData tSQLString & " order by " & mOrder
    Show vbModal
End Sub

Private Sub RefreshData(ByVal tSQL As String)
    FormMenu.Adodc1.RecordSource = tSQL
    TDBGrid1.DataSource = FormMenu.Adodc1
    FormMenu.Adodc1.Refresh
    FormMenu.Adodc1.Recordset.MoveLast
    mRowCount = FormMenu.Adodc1.Recordset.Bookmark + 1
    FormMenu.Adodc1.Recordset.MoveFirst
    For i = 0 To TDBGrid1.Columns.Count - 1
        TDBGrid1.Columns(i).Caption = mHeaders(i)
        TDBGrid1.Columns(i).Width = mWidths(i)
        TDBGrid1.Columns(i).ButtonHeader = True
    Next
End Sub

Private Sub fFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        LastRow = TDBGrid1.Bookmark
        LastCol = TDBGrid1.Col
        Row = LastRow
        Col = LastCol
        Do
            Col = Col + 1
            If Col = TDBGrid1.Columns.Count Then
                Col = 0
                Row = Row + 1
                If Row = mRowCount Then
                    Row = 1
                End If
            End If
            If InStr(1, TDBGrid1.Columns(Col).CellValue(Row), fFind, vbTextCompare) <> 0 Then
                TDBGrid1.Bookmark = Row
                TDBGrid1.Col = Col
                TDBGrid1.SetFocus
                fFind.Cancel = True
                Exit Sub
            End If
            If Col = LastCol And Row = LastRow Then
                MsgBox "Not Found"
                Exit Sub
            End If
        Loop
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - 2 * TDBGrid1.Left
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub


Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If mTypes(ColIndex) = "Date" Then
        Value = cTanggal(Value)
    ElseIf mTypes(ColIndex) = "Decimal" Then
        Value = cDecimal(Value)
    End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    mOrder = mFields(ColIndex)
    RefreshData mSQLString & MyFilter & " order by " & mOrder
End Sub

Private Function MyFilter()
    MyFilter = ""
    For i = 0 To Labels.Count - 1
        If Texts(i).Tag = "String" Then
            If Texts(i) <> "" Then MyFilter = MyFilter & " and " & Labels(i).Tag & " like '" & Texts(i) & "%'"
        ElseIf Texts(i).Tag = "Date" Then
            If cD(Texts(i)) <> 0 Then MyFilter = MyFilter & " and " & Labels(i).Tag & "=" & cD(Texts(i))
        Else
            If Texts(i) <> "" Then MyFilter = MyFilter & " and " & Labels(i).Tag & "=" & Texts(i)
        End If
    Next
    If MyFilter <> "" Then MyFilter = " where " & Mid(MyFilter, 6)
End Function

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        fFind_KeyDown 13, 0
    End If
End Sub

Private Sub Texts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        a = mSQLString & MyFilter & " order by " & mOrder
        RefreshData a
    End If
End Sub


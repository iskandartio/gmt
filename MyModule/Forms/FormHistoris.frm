VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{E2D3646A-2684-4DDE-BE47-3323E01328EE}#1.0#0"; "usrtext.ocx"
Begin VB.Form FormHistoris 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton fCopy 
      Caption         =   "&COPY"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
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
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Named:id=33:Normal"
      _StyleDefs(31)  =   ":id=33,.parent=0"
      _StyleDefs(32)  =   "Named:id=34:Heading"
      _StyleDefs(33)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(34)  =   ":id=34,.wraptext=-1"
      _StyleDefs(35)  =   "Named:id=35:Footing"
      _StyleDefs(36)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(37)  =   "Named:id=36:Selected"
      _StyleDefs(38)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(39)  =   "Named:id=37:Caption"
      _StyleDefs(40)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(41)  =   "Named:id=38:HighlightRow"
      _StyleDefs(42)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(43)  =   "Named:id=39:EvenRow"
      _StyleDefs(44)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(45)  =   "Named:id=40:OddRow"
      _StyleDefs(46)  =   ":id=40,.parent=33"
      _StyleDefs(47)  =   "Named:id=41:RecordSelector"
      _StyleDefs(48)  =   ":id=41,.parent=34"
      _StyleDefs(49)  =   "Named:id=42:FilterBar"
      _StyleDefs(50)  =   ":id=42,.parent=33"
   End
   Begin UsrText.IText IText1 
      Height          =   330
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FormHistoris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col1 As TrueOleDBGrid80.Columns
Dim x As New XArrayDB
Dim mOptions() As String
Dim mHeaders() As String
Dim mHeader() As String
Dim mWidths() As String
Dim mWidth() As String
Dim mTypes() As String
Dim mType() As String
Dim mSQLs() As String
Dim mLastOption As Integer
Dim mLastQuery() As String
Dim mfHeaders() As String
Dim mfHeader() As String
Dim mfWidths() As String
Dim mfWidth() As String
Dim mfTypes() As String
Dim mfType() As String
Dim mfSQLFields() As String
Dim mfSQLField() As String
Dim mProp As New ClassProperties
Dim mPrepared As Boolean
Dim mf As Form
Dim mReportForm() As String
Dim mParams() As String
Dim mParam() As String
Dim mUseParam As Boolean

Sub SetForm(ByVal f As Form)
    Set mf = f
End Sub

Sub SetReportForm(ByVal tReportForm As String, Optional ByVal tIndex As Integer)
    For i = tIndex To UBound(mOptions)
        mReportForm(i) = tReportForm
    Next
End Sub

Private Sub SetCaption(ByVal tCaption As String)
    Caption = tCaption
End Sub

Private Sub fCopy_Click()
    Clipboard.Clear
    a = ""
    For i = 0 To x.UpperBound(1)
        b = ""
        For j = 0 To x.UpperBound(2)
            If col1(j).Visible Then
                b = b & vbTab & x(i, j)
            End If
        Next
        a = a & vbCrLf & Mid(b, 2)
    Next
    Clipboard.SetText Mid(a, 3)
End Sub

Private Sub Form_Load()
    mUseParam = False
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    mLastOption = 0
    mPrepared = False
    TDBGrid1.Font.Size = 10
End Sub

Sub SetLastQuery(ByVal tLastQuery As String, Optional ByVal tIndex As Integer)
    For i = tIndex To UBound(mLastQuery)
        mLastQuery(i) = tLastQuery
    Next
End Sub

Sub SetOptions(ByVal tOptions As String)
    mOptions = Split(tOptions, "@")
    ReDim mHeaders(UBound(mOptions))
    ReDim mWidths(UBound(mOptions))
    ReDim mTypes(UBound(mOptions))
    ReDim mfHeaders(UBound(mOptions))
    ReDim mfWidths(UBound(mOptions))
    ReDim mfTypes(UBound(mOptions))
    ReDim mfSQLFields(UBound(mOptions))
    ReDim mSQLs(UBound(mOptions))
    ReDim mLastQuery(UBound(mOptions))
    ReDim mParams(UBound(mOptions))
    ReDim mReportForm(UBound(mOptions))
    For i = 0 To UBound(mOptions)
        If Option1.UBound < i Then
            Load Option1(i)
            Option1(i).Visible = True
        End If
        Option1(i).Caption = mOptions(i)
        FontSize = Option1(i).FontSize
        FontBold = Option1(i).FontBold
        FontName = Option1(i).FontName
        Option1(i).Width = TextWidth(mOptions(i)) + 350
        If i > 0 Then Option1(i).Left = Option1(i - 1).Left + Option1(i - 1).Width + 50
    Next
End Sub

Sub SetParam(ByVal tParams As String, Optional ByVal tIndex As Integer = 0)
    mUseParam = True
    For i = tIndex To UBound(mHeaders)
        mParams(i) = tParams
    Next
End Sub

Sub SetHeader(ByVal tHeaders As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mHeaders)
        mHeaders(i) = tHeaders
    Next
End Sub

Sub SetWidth(ByVal tWidths As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mHeaders)
        mWidths(i) = tWidths
    Next
End Sub

Sub SetType(ByVal tTypes As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mHeaders)
        mTypes(i) = tTypes
    Next
End Sub

Sub SetFilterHeader(ByVal tHeaders As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mHeaders)
        mfHeaders(i) = tHeaders
    Next
End Sub

Sub SetFilterWidth(ByVal tWidths As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mHeaders)
        mfWidths(i) = tWidths
    Next
End Sub

Sub SetFilterType(ByVal tTypes As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mHeaders)
        mfTypes(i) = tTypes
    Next
End Sub

Sub SetFilterDatas(ByVal tFilterDatas As String)
    b = Split(tFilterDatas, "@")
    For i = 0 To UBound(b)
        IText1(i) = b(i)
    Next
End Sub

Sub SetFilterSQLField(ByVal tSQLFields As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mfHeaders)
        mfSQLFields(i) = tSQLFields
    Next
End Sub

Sub SetSQL(ByVal tSQLs As String, Optional ByVal tIndex As Integer = 0)
    For i = tIndex To UBound(mSQLs)
        mSQLs(i) = tSQLs
    Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left - 100
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 100
End Sub

Private Sub fPrint_Click()
Dim res() As Variant
    If mReportForm(mLastOption) <> "" Then
        ReDim res(x.UpperBound(2), x.UpperBound(1))
        For i = 0 To x.UpperBound(1)
            For j = 0 To x.UpperBound(2)
                res(j, i) = x(i, j)
            Next
        Next
        'FormPreview.LoadFromData Me, mReportForm(mLastOption), res
    Else
        TDBGrid1.PrintInfo.SettingsMarginTop = 500
        TDBGrid1.PrintInfo.SettingsMarginLeft = 500
        TDBGrid1.PrintInfo.SettingsMarginRight = 500
        TDBGrid1.PrintInfo.SettingsMarginBottom = 500
        TDBGrid1.PrintInfo.PreviewCaption = Caption
        TDBGrid1.PrintInfo.PageHeader = Caption
        TDBGrid1.PrintInfo.PrintPreview
    End If
End Sub

Private Sub IText1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And IText1(Index).Tag <> "0" Then DoQuery
End Sub

Private Sub Option1_Click(Index As Integer)
    If mPrepared Then
        j = 0
        For i = 0 To UBound(mfSQLField)
            mProp(mfSQLField(i)) = IText1(j)
            j = j + 1
            If mfType(i) = "Date" Or mfType(i) = "Integer" Or mfType(i) = "Decimal" Then
                mProp(mfSQLField(i) & "!") = IText1(j)
                j = j + 1
            End If
        Next
        PrepareControl Index
        j = 0
        For i = 0 To UBound(mfSQLField)
            IText1(j) = mProp(mfSQLField(i))
            j = j + 1
            If mfType(i) = "Date" Or mfType(i) = "Integer" Or mfType(i) = "Decimal" Then
                IText1(j) = mProp(mfSQLField(i) & "!")
                j = j + 1
            End If
        Next
    End If
    DoQuery
End Sub

Sub PrepareControl(ByVal Index As Integer)
On Error Resume Next
    If mOptions(Index) = mOptions(mLastOption) And mPrepared Then Exit Sub
    If mHeaders(Index) <> mHeaders(mLastOption) Or Not mPrepared Then
        mHeader = Split(mHeaders(Index), "@")
        mWidth = Split(mWidths(Index), "@")
        mType = Split(mTypes(Index), "@")
        While col1.Count > 0
            col1.Remove 0
        Wend
        For i = 0 To UBound(mHeader)
            If Left(mHeader(i), 1) = "*" Then
                col1.Add(i).Caption = Mid(mHeader(i), 2)
                col1(i).AllowSizing = False
            Else
                col1.Add(i).Caption = mHeader(i)
                col1(i).Visible = True
                col1(i).Width = mWidth(i)
                col1(i).Tag = mType(i)
            End If
        Next
        TDBGridLoad TDBGrid1
        TDBGrid1.AllowUpdate = False
   End If
   If mfHeaders(Index) <> mfHeaders(mLastOption) Or Not Prepared Then
        mfHeader = Split(mfHeaders(Index), "@")
        mfWidth = Split(mfWidths(Index), "@")
        mfType = Split(mfTypes(Index), "@")
        mfSQLField = Split(mfSQLFields(Index), "@")
        j = Label1.Count
        j1 = 0
        k = IText1.Count
        k1 = 0
        For i = 0 To UBound(mfHeader)
            If i > Label1.Count - 1 Then
                Load Label1(i)
            End If
            Label1(i).Visible = True
            If k1 > IText1.Count - 1 Then
                Load IText1(k1)
                IText1(k1).MaxLength = 0
                IText1(k1).TabIndex = IText1(k1 - 1).TabIndex + 1
            End If
            IText1(k1).Visible = True
            IText1(k1) = ""
            Label1(i).Caption = mfHeader(i)
            Label1(i).Width = mfWidth(i)
            IText1(k1).DataType = tString
            IText1(k1).Width = mfWidth(i)
            If mfType(i) = "Date" Then IText1(k1).DataType = tDate
            If k1 > 0 Then
                IText1(k1).Left = IText1(k1 - 1).Left + IText1(k1 - 1).Width + 50
            End If
            Label1(i).Left = IText1(k1).Left
            IText1(k1).Tag = ""
            k1 = k1 + 1
            If mfType(i) = "Integer" Or mfType(i) = "Decimal" Or mfType(i) = "Date" Then
                If k1 > IText1.Count - 1 Then
                    Load IText1(k1)
                    IText1(k1).MaxLength = 0
                    IText1(k1).TabIndex = IText1(k1 - 1).TabIndex + 1
                End If
                IText1(k1 - 1).Tag = "0"
                IText1(k1).Visible = True
                IText1(k1) = ""
                IText1(k1).DataType = tString
                If mfType(i) = "Date" Then IText1(k1).DataType = tDate
                IText1(k1).Left = IText1(k1 - 1).Left + IText1(k1 - 1).Width + 50
                IText1(k1).Width = mfWidth(i)
                k1 = k1 + 1
            End If
        Next
        For i = UBound(mfHeader) + 1 To Label1.Count - 1
            Label1(i).Visible = False
        Next
        For i = k1 To IText1.Count - 1
            IText1(i).Visible = False
        Next
    End If
    mPrepared = True
    mLastOption = Index
End Sub

Sub DoQuery()
On Error Resume Next
    b = mSQLs(mLastOption)
    mfType = Split(mfTypes(mLastOption), "@")
    a = b & MyFilter & mLastQuery(mLastOption)
    query a
    TDBGridClear TDBGrid1
    If RS.RecordCount > 0 Then
        x.LoadRows RS.GetRows
        For i = 0 To x.UpperBound(1)
            For j = 0 To x.UpperBound(2)
                If col1(j).Tag = "Date" Then
                    x(i, j) = cTanggal2(x(i, j))
                End If
            Next
        Next
    End If
    TDBGrid1.Rebind
End Sub

Function MyFilter() As String
On Error Resume Next
    k1 = 0
    For i = 0 To UBound(mfHeader)
        If mfType(i) = "String" Then
            If Trim(IText1(k1)) <> "" And Left(mfSQLField(i), 1) <> "*" Then MyFilter = MyFilter & " and " & mfSQLField(i) & " like '" & IText1(k1) & "%'"
            k1 = k1 + 1
        ElseIf mfType(i) = "StringExact" Then
            If Trim(IText1(k1)) <> "" And Left(mfSQLField(i), 1) <> "*" Then MyFilter = MyFilter & " and " & mfSQLField(i) & " like '" & IText1(k1) & "'"
            k1 = k1 + 1
        ElseIf mfType(i) = "Date" Then
            If cD(IText1(k1)) <> 0 And Left(mfSQLField(i), 1) <> "*" Then MyFilter = MyFilter & " and " & mfSQLField(i) & ">=" & cD(IText1(k1))
            k1 = k1 + 1
            If cD(IText1(k1)) <> 0 And Left(mfSQLField(i), 1) <> "*" Then MyFilter = MyFilter & " and " & mfSQLField(i) & "<=" & cD(IText1(k1))
            k1 = k1 + 1
        ElseIf mfType(i) = "Decimal" Or mfType(i) = "Integer" Then
            If IText1(k1) <> "" And Left(mfSQLField(i), 1) <> "*" Then MyFilter = MyFilter & " and " & mfSQLField(i) & ">=" & cNum(IText1(k1))
            k1 = k1 + 1
            If IText1(k1) <> "" And Left(mfSQLField(i), 1) <> "*" Then MyFilter = MyFilter & " and " & mfSQLField(i) & "<=" & cNum(IText1(k1))
            k1 = k1 + 1
        End If
    Next
End Function

Sub StartMe()
    Option1(0).Value = True
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If col1(ColIndex).Tag = "NoCents" Then
        Value = cNoCents(Value)
    ElseIf col1(ColIndex).Tag = "Date" Then
        Value = cTanggal(Value)
    ElseIf Left(col1(ColIndex).Tag, 8) = "Zerofill" Then
        Value = zerofill(Value, Mid(col1(ColIndex).Tag, 9))
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    tKey = KeyCode
    TDBGridKeyDown TDBGrid1, KeyCode
    mf.AfterKeyDown tKey
End Sub


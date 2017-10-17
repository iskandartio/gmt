VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.UserControl UsrList 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   ScaleHeight     =   5085
   ScaleWidth      =   11895
   Begin VB.CommandButton fCopy 
      Caption         =   "&COPY"
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Top             =   780
      Width           =   735
   End
   Begin VB.CommandButton fSummary 
      Caption         =   "S"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   395
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton fPrev 
         Appearance      =   0  'Flat
         Caption         =   "<"
         Height          =   375
         Left            =   10
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   10
         Width           =   255
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   265
         TabIndex        =   6
         Top             =   10
         Width           =   3600
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "Command1"
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton fNext 
         Appearance      =   0  'Flat
         Caption         =   ">"
         Height          =   375
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   10
         Width           =   255
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "P"
      Height          =   315
      Left            =   200
      TabIndex        =   3
      Top             =   1320
      Width           =   195
   End
   Begin UsrText.IText Text1 
      Height          =   270
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6588
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.fgcolor=&H80000008&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Times New Roman"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bgcolor=&H8000000E&"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(19)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=13,.parent=3,.bgcolor=&HE1FFFF&,.fgcolor=&HFF0000&"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(32)  =   "Named:id=29:Normal"
      _StyleDefs(33)  =   ":id=29,.parent=0,.valignment=2"
      _StyleDefs(34)  =   "Named:id=30:Heading"
      _StyleDefs(35)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(36)  =   ":id=30,.wraptext=-1"
      _StyleDefs(37)  =   "Named:id=31:Footing"
      _StyleDefs(38)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(39)  =   "Named:id=32:Selected"
      _StyleDefs(40)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(41)  =   "Named:id=33:Caption"
      _StyleDefs(42)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(43)  =   "Named:id=34:HighlightRow"
      _StyleDefs(44)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(45)  =   "Named:id=35:EvenRow"
      _StyleDefs(46)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(47)  =   "Named:id=36:OddRow"
      _StyleDefs(48)  =   ":id=36,.parent=29"
      _StyleDefs(49)  =   "Named:id=39:RecordSelector"
      _StyleDefs(50)  =   ":id=39,.parent=30"
      _StyleDefs(51)  =   "Named:id=42:FilterBar"
      _StyleDefs(52)  =   ":id=42,.parent=29"
   End
   Begin UsrText.IText fFind 
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   840
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Find"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "UsrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mParams As String
Dim mActiveControlIndex As Integer
Dim mSQL As String
Dim mTypesCurrent() As String
Dim mFilterTypes() As String
Dim mLastSQL As String
Dim RS As New ADODB.Recordset
Dim x As New XArrayDB
Dim mActive As String
Dim mLast As Boolean
Dim z As New ClassProperties
Dim FieldIndex As New ClassProperties
Dim mOrder() As Byte
Dim mLeftColumn As Integer
Dim mf As Form
Dim rs1() As Variant
'Event Declarations:
Event Click()
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
'Event Click() 'MappingInfo=Command1(0) ,Command1, 0, Click

Sub query(ByVal QueryString As String)
'on error GoTo err
    If RS.State = 1 Then RS.Close
    QueryString = Replace(QueryString, "~", pTipe)
    RS.Open Trim(QueryString)
err:
End Sub

Function GetFooterText() As String
Dim a As String
    a = ""
    For i = 0 To TDBGrid1.Columns.Count - 1
        If TDBGrid1.Columns(i).FooterText <> "" Then
            a = a & "@" & i & "#" & CDec(TDBGrid1.Columns(i).FooterText)
        End If
    Next
    GetFooterText = Mid(a, 2)
End Function

Function SetFooterText(ByVal tFooterText As String)
Dim a() As String
    a = Split(tFooterText, "@")
Dim b() As String
    For i = 0 To UBound(a)
        b = Split(a(i), "#")
        TDBGrid1.Columns(CLng(b(0))).FooterText = cDecimal(b(1))
    Next
End Function

Sub SetConn(rs1 As Object)
    Set RS = rs1
End Sub

Sub LoadMe(ByVal tCommandTags As String, tf As Form, Optional ByVal tCaption As String = "", Optional ByVal tParams As String = "")
    Dim a() As String
    Dim b() As String
    If tParams <> "" Then
        Dim tCommandTags2 As String
        If tCommandTags = "" Then Exit Sub
        a = Split(tParams, "@")
        j = 0
        Do
            k = InStr(1, tCommandTags, "$")
            If k = 0 Then Exit Do
            tCommandTags2 = tCommandTags2 & Left(tCommandTags, k - 1)
            tCommandTags = Replace(tCommandTags, "$", a(j), k, 1)
            j = j + 1
        Loop
    End If
    tCommandTags = tCommandTags2 & tCommandTags
    a = Split(tCommandTags, "|")
    b = Split(a(0), "#")
    Command1(0).Caption = b(0)
    Command1(0).Width = TextWidth(b(0)) + 1000
    Command1(0).Tag = Mid(a(0), Len(b(0)) + 2)
    Combo1.Clear
    Combo1.List(0) = Command1(0).Caption
    For i = 1 To UBound(a)
        If i > Command1.UBound Then Load Command1(i)
        b = Split(a(i), "#")
        Command1(i).Visible = True
        Command1(i).Left = Command1(i - 1).Left + Command1(i - 1).Width
        Command1(i).Caption = b(0)
        Command1(i).Width = TextWidth(b(0)) + 1000
        Command1(i).Tag = Mid(a(i), Len(b(0)) + 2)
        Combo1.List(i) = Command1(i).Caption
    Next
    If UBound(a) = 0 Then
        Combo1.Visible = False
        Frame2.Visible = False
        Label1(0).Top = Combo1.Top
        Text1(0).Top = Label1(0).Top + Label1(0).Height + 50
        TDBGrid1.Top = Text1(0).Top + Text1(0).Height + 100
        Command2.Top = TDBGrid1.Top + 100
    End If
    Frame1.Width = Command1(i - 1).Left + Command1(i - 1).Width
    For i = UBound(a) + 1 To Command1.UBound
        Unload Command(i)
    Next
    Set mf = tf
    mf.Tag = tCaption
    mLeftColumn = 0
    SetVisibility
'    CommandClick 0
    fPrev.ZOrder 0
    fNext.ZOrder 0
End Sub

Private Sub SetVisibility()
    If mLeftColumn < 0 Then mLeftColumn = 0
    If mLeftColumn = Command1.Count Then mLeftColumn = Command1.Count - 1
    Frame1.Left = 265 - Command1(mLeftColumn).Left
End Sub

Function GetActive()
    GetActive = mActive
End Function

Private Sub SaveCurrentFilterValue()
On Error GoTo err
    For i = 0 To Label1.Count - 1
        z(Label1(i)) = Text1(i)
        If mFilterTypes(i) = "Integer" Or mFilterTypes(i) = "Date" Then
            i = i + 1
            z(Label1(i - 1) & "!") = Text1(i)
        End If
    Next
err:
End Sub

Sub CommandClick(ByVal tIndex As Integer)
'on error Resume Next
    Combo1 = Command1(tIndex).Caption
    Frame3.Visible = False
    MousePointer = vbHourglass
    SaveCurrentFilterValue
    mActive = Command1(tIndex).Caption
    For i = 0 To Command1.UBound
        Command1(i).FontSize = 8
        Command1(i).FontBold = False
    Next
    Command1(tIndex).FontSize = 8
    Command1(tIndex).FontBold = True
    Dim a() As String
    Dim b() As String
    Dim c() As String
    Dim d() As String
    a = Split(Command1(tIndex).Tag, "#")
    mSQL = a(0)
    b = Split(a(1), "@") ' Caption
    c = Split(a(2), "@") ' Field
    d = Split(a(3), "@") ' Widths
    e = Split(a(4), "@") ' Types
    mFilterTypes = Split(a(4), "@") ' Types
    k = 0
    Text1(0).MaxLength = 0
    Text1(0).DataType = 0
    Label1(0).Caption = b(0)
    Text1(0).Width = d(0)
    Label1(0).Tag = c(0)
    Text1(0).Tag = e(0)
    If e(0) = "Date" Then Text1(0).DataType = 3 Else Text1(0).DataType = 0
    Text1(0).Text = IIf(z(Label1(0)) = "", Text1(0).Text, z(Label1(0)))
    If e(0) = "Date" Or e(0) = "Integer" Then
        k = k + 1
        If k > Label1.UBound Then
            Load Label1(k)
            Load Text1(k)
        End If
        Label1(k).Visible = True
        Text1(k).Visible = True
        Text1(k).TabIndex = Text1(k - 1).TabIndex + 1
        Label1(k).Caption = ""
        Label1(k).Tag = c(0)
        Text1(k).MaxLength = 0
        If e(0) = "Date" Then Text1(k).DataType = 3 Else Text1(k).DataType = 1
        Text1(k).Text = IIf(z(Label1(k - 1) & "!") = "", Text1(k).Text, z(Label1(k - 1) & "!"))
        Text1(k).Tag = e(0)
        Text1(k).Width = d(0)
        Text1(k).Left = Text1(k - 1).Left + Text1(k - 1).Width
        Label1(k).Left = Text1(k).Left
    End If
    For i = 1 To UBound(b)
        k = k + 1
        If k > Label1.UBound Then
            Load Label1(k)
            Load Text1(k)
        End If
        Text1(k).MaxLength = 0
        Text1(k).DataType = 0
        Label1(k).Visible = True
        Text1(k).Visible = True
        Label1(k).Caption = b(i)
        Label1(k).Tag = c(i)
        Text1(k).MaxLength = 0
        If e(i) = "Date" Then Text1(k).DataType = 3 Else Text1(k).DataType = 0
        Text1(k).Tag = e(i)
        Text1(k).Text = IIf(z(Label1(k)) = "", Text1(k).Text, z(Label1(k)))
        Text1(k).Width = d(i)
        Text1(k).Left = Text1(k - 1).Left + Text1(k - 1).Width
        Label1(k).Left = Text1(k).Left
        If e(i) = "Date" Or e(i) = "Integer" Then
            k = k + 1
            If k > Label1.UBound Then
                Load Label1(k)
                Load Text1(k)
            End If
            Label1(k).Visible = True
            Text1(k).Visible = True
            Label1(k).Caption = ""
            Label1(k).Tag = c(i)
            Text1(k).MaxLength = 0
            If e(i) = "Date" Then Text1(k).DataType = 3 Else Text1(k).DataType = 1
            Text1(k).Text = IIf(z(Label1(k - 1) & "!") = "", Text1(k).Text, z(Label1(k - 1) & "!"))
            Text1(k).Tag = e(i)
            Label1(k).Visible = False
            Text1(k).Width = d(i)
            Text1(k).Left = Text1(k - 1).Left + Text1(k - 1).Width
            Label1(k).Left = Text1(k).Left
        End If
    Next
    For i = k + 1 To Label1.UBound
        Label1(i).Visible = False
        Text1(i).Visible = False
    Next
    b = Split(a(5), "@")             ' Headers
    c = Split(a(6), "@")             ' Widths
    mTypesCurrent = Split(a(7), "@") ' Types
    mLastSQL = a(8)                  ' LastSQL
    TDBGrid1.HeadingStyle.Alignment = dbgCenter
    FieldIndex.Clear
    For i = 0 To UBound(b)
        TDBGrid1.Columns.Add(i).Visible = True
        TDBGrid1.Columns(i).Caption = b(i)
        'TDBGrid1.Columns(i).ButtonHeader = True
        FieldIndex(b(i)) = i
        TDBGrid1.Columns(i).Width = c(i)
        If mTypesCurrent(i) = "NoCents" Then
            TDBGrid1.Columns(i).Alignment = dbgRight
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
        ElseIf mTypesCurrent(i) = "Integer" Then
            TDBGrid1.Columns(i).Alignment = dbgRight
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
        ElseIf mTypesCurrent(i) = "Decimal" Then
            TDBGrid1.Columns(i).Alignment = dbgRight
            TDBGrid1.Columns(i).NumberFormat = "Standard"
        ElseIf mTypesCurrent(i) = "Date" Then
            TDBGrid1.Columns(i).Tag = "Date"
        ElseIf mTypesCurrent(i) = "YesNo" Then
            TDBGrid1.Columns(i).ValueItems.Presentation = dbgCheckBox
            TDBGrid1.Columns(i).Alignment = dbgCenter
        ElseIf Left(mTypesCurrent(i), 8) = "Zerofill" Then
            TDBGrid1.Columns(i).NumberFormat = "FormatText Event"
        End If
    Next
    ReDim mOrder(UBound(b))
    k = UBound(b) + 1
    l = TDBGrid1.Columns.Count - 1
    For i = k To l
        TDBGrid1.Columns.Remove k
    Next
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
'    TextKeyDown tIndex
    mf.Caption = IIf(mf.Tag = "", "", mf.Tag & ": ") & mActive
    MousePointer = vbDefault
End Sub

Sub TextKeyDown(ByVal tIndex As Integer)
    Frame3.Visible = True
    Frame3.ZOrder 0
Dim MyFilter As String
Dim MyFilter2 As String
    MyFilter = ""
    MyFilter2 = ""
    For i = 0 To Label1.UBound
        If Text1(i) <> "" And Text1(i).Visible Then
            If Left(Label1(i).Tag, 4) <> "sum(" Then
                If Left(Text1(i).Tag, 8) = "Zerofill" Then
                    MyFilter = MyFilter & " and " & Label1(i).Tag & "=" & Text1(i)
                ElseIf Text1(i).Tag = "Integer" Then
                    If Label1(i) = "" Then
                        MyFilter = MyFilter & " and " & Label1(i).Tag & "<=" & Text1(i)
                    Else
                        MyFilter = MyFilter & " and " & Label1(i).Tag & ">=" & Text1(i)
                    End If
                ElseIf Text1(i).Tag = "Date" Then
                    If Label1(i) = "" Then
                        If cD(Text1(i)) <> "A" Then MyFilter = MyFilter & " and " & Label1(i).Tag & "<=" & cD(Text1(i))
                    Else
                        If cD(Text1(i)) <> "A" Then MyFilter = MyFilter & " and " & Label1(i).Tag & ">=" & cD(Text1(i))
                    End If
                ElseIf Text1(i).Tag = "StringExact" Then
                    MyFilter = MyFilter & " and " & Label1(i).Tag & " like '" & Text1(i) & "'"
                ElseIf Text1(i).Tag = "IntegerExact" Then
                    If Text1(i) <> "" Then
                        If Text1(i) = "0" Then
                            MyFilter = MyFilter & " and " & Label1(i).Tag & "=0"
                        Else
                            MyFilter = MyFilter & " and " & Label1(i).Tag & "<>0"
                        End If
                    End If
                Else
                    MyFilter = MyFilter & " and " & Label1(i).Tag & " like '" & Text1(i) & "%'"
                End If
            Else
                If Left(Text1(i).Tag, 8) = "Zerofill" Then
                    MyFilter2 = MyFilter2 & " and " & Label1(i).Tag & "=" & Text1(i)
                ElseIf Text1(i).Tag = "Integer" Then
                    If Label1(i) = "" Then
                        MyFilter2 = MyFilter2 & " and " & Label1(i).Tag & "<=" & Text1(i)
                    Else
                        MyFilter2 = MyFilter2 & " and " & Label1(i).Tag & ">=" & Text1(i)
                    End If
                ElseIf Text1(i).Tag = "Date" Then
                    If Label1(i) = "" Then
                        If cD(Text1(i)) <> "A" Then MyFilter2 = MyFilter2 & " and " & Label1(i).Tag & "<=" & cD(Text1(i))
                    Else
                        If cD(Text1(i)) <> "A" Then MyFilter2 = MyFilter2 & " and " & Label1(i).Tag & ">=" & cD(Text1(i))
                    End If
                ElseIf Text1(i).Tag = "StringExact" Then
                    MyFilter2 = MyFilter2 & " and " & Label1(i).Tag & " like '" & Text1(i) & "'"
                Else
                    MyFilter2 = MyFilter2 & " and " & Label1(i).Tag & " like '" & Text1(i) & "%'"
                End If
            End If
        End If
    Next
    a = mSQL & MyFilter & mLastSQL & MyFilter2
    a = Replace(a, "$", pTipe)
    query a
    
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then
        rs1 = RS.GetRows
        x.LoadRows rs1
    End If
Dim total() As Double
    ReDim total(TDBGrid1.Columns.Count - 1)
    For i = 0 To x.UpperBound(1)
        For j = 0 To TDBGrid1.Columns.Count - 1
            If mTypesCurrent(j) = "NoCents" Or mTypesCurrent(j) = "Decimal" Then
                total(j) = total(j) + x(i, j)
            ElseIf mTypesCurrent(j) = "Date" Then
                x(i, j) = cTanggal(x(i, j))
            End If
        Next
    Next
    For j = 0 To TDBGrid1.Columns.Count - 1
        If mTypesCurrent(j) = "NoCents" Then
            TDBGrid1.Columns(j).FooterText = cNoCents(total(j))
        ElseIf mTypesCurrent(j) = "Decimal" Then
            TDBGrid1.Columns(j).FooterText = cDecimal(total(j))
        End If
    Next
    TDBGrid1.Rebind
    Frame3.Visible = False
    mf.PascaLoading mActive
End Sub

Sub SummaryClick()
'on error GoTo err
    MousePointer = vbHourglass
    Dim curval() As Variant
    ReDim curval(UBound(rs1))
    j = 0
    For i = 0 To x.UpperBound(2)
        'If TDBGrid1.Columns(i).Visible Then
            curval(j) = x(0, i)
            j = j + 1
        'End If
    Next
    j = 0
    m = 0
    For i = 1 To x.UpperBound(1)
        n = j
        Do
            If mTypesCurrent(j) = "NoCents" Or mTypesCurrent(j) = "Decimal" Then
                curval(j) = curval(j) + x(i, j)
                If j = x.UpperBound(2) Then j = -1
                j = j + 1
                If n = j Then Exit Do
            ElseIf TDBGrid1.Columns(j).Visible And x(i, j) <> curval(j) Then
                For k = 0 To x.UpperBound(2)
                    x(m, k) = curval(k)
                    curval(k) = x(i, k)
                Next
                m = m + 1
                Exit Do
            Else
                If j = x.UpperBound(2) Then j = -1
                j = j + 1
                If n = j Then Exit Do
            End If
        Loop
    Next
    For k = 0 To x.UpperBound(2)
        x(m, k) = curval(k)
    Next
    x.ReDim 0, m, 0, x.UpperBound(2)
    TDBGrid1.Rebind
err:
    MousePointer = vbDefault
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
'on error Resume Next
    If KeyCode = 13 Then
        If Combo1.ListIndex = -1 Then Exit Sub
        Command1_Click Combo1.ListIndex
    End If
End Sub
Private Sub Command1_Click(Index As Integer)
    If Not mLast Then RaiseEvent Click
    CommandClick Index
    If mLast Then
        RaiseEvent Click
        mLast = False
    End If
    Text1(0).SetFocus
End Sub

Sub LastEvent()
    mLast = True
End Sub


Private Sub Command2_Click()
    TDBGrid1.PrintInfo.SettingsMarginBottom = 500
    TDBGrid1.PrintInfo.SettingsMarginLeft = 500
    TDBGrid1.PrintInfo.SettingsMarginRight = 500
    TDBGrid1.PrintInfo.SettingsMarginTop = 500
    'TDBGrid1.PrintInfo.PageSetup
    TDBGrid1.PrintInfo.PrintPreview
End Sub

Sub SetPrinterOrientation(ByVal tOrientation As Long)
    TDBGrid1.PrintInfo.SettingsOrientation = tOrientation
End Sub

Private Sub fCopy_Click()
    CopyGrid TDBGrid1
End Sub

Private Sub fFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim cCol As Integer
        cCol = TDBGrid1.Col
        n = TDBGrid1.Bookmark
        m = n + 1
        If m = x.UpperBound(1) + 1 Then m = 0
        Do While m <> n
            If InStr(1, x(m, cCol), fFind, vbTextCompare) <> 0 Then Exit Do
            m = m + 1
            If m = x.UpperBound(1) + 1 Then m = 0
        Loop
        If InStr(1, x(m, cCol), fFind, vbTextCompare) = 0 Then
            MsgBox "Not Found"
        End If
        TDBGrid1.Bookmark = m
        TDBGrid1.SetFocus
        fFind.Cancel = True
    End If
End Sub

Private Sub fPrev_Click()
    mLeftColumn = mLeftColumn - 1
    SetVisibility
End Sub

Private Sub fNext_Click()
    mLeftColumn = mLeftColumn + 1
    SetVisibility
End Sub

Private Sub fSummary_Click()
    SummaryClick
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
'on error Resume Next
    If mTypesCurrent(ColIndex) = "NoCents" Or mTypesCurrent(ColIndex) = "Integer" Then
        Value = cNoCents(Value)
    ElseIf Left(mTypesCurrent(ColIndex), 8) = "Zerofill" Then
        Value = zerofill(Value, Mid(mTypesCurrent(ColIndex), 9))
    End If
End Sub
    
Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim TypeX As XTYPE
    If InStr(mTypesCurrent(ColIndex), "String") <> 0 Then TypeX = XTYPE_STRING Else TypeX = XTYPE_DOUBLE
    x.QuickSort 0, x.UpperBound(1), ColIndex, mOrder(ColIndex), TypeX
    TDBGrid1.Rebind
    If mOrder(ColIndex) = XORDER_ASCEND Then
        mOrder(ColIndex) = XORDER_DESCEND
    Else
        mOrder(ColIndex) = XORDER_ASCEND
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'on error GoTo err
    If KeyCode = 32 And mTypesCurrent(0) = "YesNo" Then
        TDBGrid1.Columns(0).Value = -1 - TDBGrid1.Columns(0).Value
    ElseIf KeyCode = 114 Then
        fFind_KeyDown 13, 0
    ElseIf KeyCode = 117 Then
        TDBGrid1.Font.Size = TDBGrid1.Font.Size + 2
        FontSize = TDBGrid1.Font.Size
        TDBGrid1.RowHeight = TextHeight("I") + 50
        KeyCode = 0
    ElseIf KeyCode = 116 Then
        KeyCode = 0
        If TDBGrid1.Font.Size <= 6 Then Exit Sub
        TDBGrid1.Font.Size = TDBGrid1.Font.Size - 2
        FontSize = TDBGrid1.Font.Size
        TDBGrid1.RowHeight = TextHeight("I") + 50
    End If
    Exit Sub
err:
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    mActiveControlIndex = Index
End Sub

Function GetActiveControl() As Integer
    GetActiveControl = mActiveControlIndex
End Function

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'on error Resume Next
    If KeyCode = 13 Then
        If Label1(Index) <> "" And Text1(Index).DataType <> 0 Then Exit Sub
        TextKeyDown Index
    ElseIf KeyCode = 112 Then
        HelpMe Label1(Index), mf, Label1(Index).Left, 1425
    End If
End Sub

Private Sub UserControl_Initialize()
    Set TDBGrid1.Array = x
End Sub

Private Function cTanggal(ByVal a As Variant) As Variant
    b = zerofill(a, 6)
    cTanggal = Right(b, 2) & "/" & Mid(b, 3, 2) & "/" & Left(b, 2)
End Function

Private Function zerofill(ByVal STR As String, ByVal num As Integer) As String
    If STR = "" Then Exit Function
    zerofill = CStr(STR)
    While Len(zerofill) < num
        zerofill = "0" & zerofill
    Wend
End Function

Private Function cNoCents(ByVal Value As Double)
    cNoCents = Format(Value, "###,###,##0.00")
End Function
Private Function cD(ByVal a As String) As String
On Error GoTo err
    b = Right(a, 2) & Mid(a, 4, 2) & Left(a, 2)
    cD = CLng(b)
    Exit Function
err:
   cD = "A"
End Function
Public Property Get Text() As String
    Text = Text1(GetActiveControl)
End Property
Public Property Get TextIndex(tIndex As Integer) As String
    TextIndex = Text1(tIndex)
End Property
Sub FocusSelect()
    Text1(GetActiveControl).FocusSelect
End Sub
Public Property Let TextIndex(tIndex As Integer, ByVal tVal As String)
    Text1(tIndex) = tVal
End Property

Public Property Get Grid(ByVal tX As Integer, ByVal tY As Integer) As String
    Grid = x(tX, tY)
End Property
Public Property Let Grid(ByVal tX As Integer, ByVal tY As Integer, ByVal tVal As String)
    x(tX, tY) = tVal
End Property
Public Property Get Grid2(ByVal tX As Integer, ByVal tY As String) As String
    Grid2 = x(tX, FieldIndex(tY))
End Property
Public Property Let Grid2(ByVal tX As Integer, ByVal tY As String, ByVal tVal As String)
    x(tX, FieldIndex(tY)) = tVal
End Property
Function UpperBound()
    UpperBound = x.UpperBound(1)
End Function
Sub Update()
    TDBGrid1.Update
End Sub
Sub Rebind()
    TDBGrid1.Rebind
End Sub
Public Property Get FirstRow() As Long
    FirstRow = TDBGrid1.FirstRow
End Property
Public Property Get Bookmark() As Long
'on error Resume Next
    Bookmark = TDBGrid1.Bookmark
End Property
Public Property Let Bookmark(ByVal tBookMark As Long)
    TDBGrid1.Bookmark = tBookMark
End Property
Private Sub UserControl_Resize()
On Error Resume Next
    Frame2.Width = ScaleWidth - Frame2.Left - 100
    fNext.Left = Frame2.Width - fNext.Width
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left * 2
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 120
    Frame3.Width = ScaleWidth - Frame3.Left * 2
    Frame3.Height = ScaleHeight - Frame3.Top - 120
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

VERSION 5.00
Object = "{F6D22ACD-8630-4FE1-97C4-D56AB4AD4DEA}#1.0#0"; "usrtruecombo.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMasterAccount 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CHART OF ACCOUNT"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   Icon            =   "FormMasterAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8730
   Tag             =   "29"
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox fDeskripsi 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   5295
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7858
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO ACCOUNT"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DESKRIPSI"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "SALDO"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NORMAL DK"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   582
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5477"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5398"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=20"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1)._MinWidth=6646905"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=20"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(2)._MinWidth=66529152"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=20"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Times New Roman"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=53,.parent=2,.namedParent=55"
      _StyleDefs(19)  =   "FilterBarStyle:id=56,.parent=1,.namedParent=58"
      _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=54,.parent=53"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=57,.parent=56"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=40,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=37,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=38,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=39,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=48,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=45,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=46,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=47,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=52,.parent=11"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=49,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=50,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=51,.parent=15"
      _StyleDefs(48)  =   "Named:id=29:Normal"
      _StyleDefs(49)  =   ":id=29,.parent=0,.valignment=2"
      _StyleDefs(50)  =   "Named:id=30:Heading"
      _StyleDefs(51)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=30,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=31:Footing"
      _StyleDefs(54)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=32:Selected"
      _StyleDefs(56)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=33:Caption"
      _StyleDefs(58)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(59)  =   "Named:id=34:HighlightRow"
      _StyleDefs(60)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(61)  =   "Named:id=35:EvenRow"
      _StyleDefs(62)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=36:OddRow"
      _StyleDefs(64)  =   ":id=36,.parent=29"
      _StyleDefs(65)  =   "Named:id=55:RecordSelector"
      _StyleDefs(66)  =   ":id=55,.parent=30"
      _StyleDefs(67)  =   "Named:id=58:FilterBar"
      _StyleDefs(68)  =   ":id=58,.parent=29"
   End
   Begin UsrTrueCombo.ITrueCombo fCOA 
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PARENT LEVEL"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHART OF ACCOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FormMasterAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoadedCOA As Boolean
Dim x As New XArrayDB


Private Sub fPrint_Click()
    FormPreview.LoadMe Me, "ChartAccount"
End Sub

Private Sub TDBGrid1_GotFocus()
On Error Resume Next
    If x.UpperBound(1) = -1 Then TDBGrid1.Row = 0
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    For i = 0 To TDBGrid1.Columns.Count - 1
        TDBGrid1.Columns(i).ButtonHeader = True
    Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top
End Sub

Private Sub fSave_Click()
On Error GoTo err
    TDBGrid1.Update
    BeginTransaction
    a = "delete from m_ChartAccount where LevelDiatas='" & fCOA & "'"
    ExecMe a
    If x.UpperBound(1) > -1 Then
        maxlen = Len(x(x.UpperBound(1), 0))
        For i = x.UpperBound(1) To 1 Step -1
            If maxlen < Len(x(i, 0)) Then maxlen = Len(x(i, 0))
        Next
        For i = 0 To x.UpperBound(1)
            noacc = DelZero(fCOA)
            noacclengkap = IIf(fCOA = "", x(i, 0) & "-00000", RZerofill(noacc & zerofill(x(i, 0), maxlen), 7))
            a = "select count(*) from m_ChartAccount where LevelDiatas='" & esc(noacclengkap) & "'"
            query a
            Child = RS.Fields(0).Value
            'noaccpendek = CInt(DelZero(x(i, 0)))
            noaccpendek = x(i, 0)
            level1 = noaccpendek
            level2 = 0
            level3 = 0
            level4 = 0
            level5 = 0
            Level6 = 0
            If fCOA <> "" Then
                a = "select top 1 Level1,Level2,Level3,Level4,Level5,Level6 from m_ChartAccount where NoAccount='" & esc(fCOA) & "'"
                query a
                level1 = RS.Fields(0).Value
                level2 = RS.Fields(1).Value
                level3 = RS.Fields(2).Value
                level4 = RS.Fields(3).Value
                level5 = RS.Fields(4).Value
                Level6 = RS.Fields(5).Value
                If level2 = 0 Then
                    level2 = noaccpendek
                ElseIf level3 = 0 Then
                    level3 = noaccpendek
                ElseIf level4 = 0 Then
                    level4 = noaccpendek
                ElseIf level5 = 0 Then
                    level5 = noaccpendek
                ElseIf Level6 = 0 Then
                    Level6 = noaccpendek
                End If
            End If
            a = "insert into m_ChartAccount(LevelDiatas,NoAccount,Deskripsi,Saldo,Child,Level1,Level2,Level3,Level4,Level5,Level6,NormalDK,Pengupdate) values('" & _
                fCOA & "','" & noacclengkap & "','" & x(i, 1) & "'," & cNum(x(i, 2)) & "," & Child & _
                "," & level1 & "," & level2 & "," & level3 & "," & level4 & "," & level5 & "," & Level6 & ",'" & x(i, 3) & "','" & pUsr & "')"
            ExecMe a
            a = "update t_JasaDetail set DebetAcc='" & noacclengkap & "' where KetAcc ='" & x(i, 1) & "'"
            ExecMe a
            a = "update t_Account set DebetAcc='" & noacclengkap & "' where KetDebetAcc ='" & x(i, 1) & "'"
            ExecMe a
            a = "update t_Account set KreditAcc='" & noacclengkap & "' where KetKreditAcc ='" & x(i, 1) & "'"
            ExecMe a
            a = "update RLSetting set NoAccount='" & noacclengkap & "' where KetAcc ='" & x(i, 1) & "'"
            ExecMe a
            a = "update NeracaSetting set NoAccount='" & noacclengkap & "' where KetAcc ='" & x(i, 1) & "'"
            ExecMe a
            a = "update m_SettingAccounting set NoAccount='" & noacclengkap & "' where KetAcc ='" & x(i, 1) & "'"
            ExecMe a
            'a = "update t_STTPembayaran set DebetAccount='" & noacclengkap & "' where KetAcc ='" & x(i, 1) & "'"
            'ExecMe a
            'a = "update t_PembelianSTTPembayaran set DebetAccount='" & noacclengkap & "' where KetAcc ='" & x(i, 1) & "'"
            'ExecMe a
            'a = "update AllSettingAccount set DebetAcc='" & noacclengkap & "' where KetDebet='" & x(i, 1) & "'"
            'ExecMe a
            'a = "update AllSettingAccount set KreditAcc='" & noacclengkap & "' where KetKredit='" & x(i, 1) & "'"
            'ExecMe a
        Next
        If fCOA <> "" Then
            a = "update m_ChartAccount set Child=" & x.UpperBound(1) + 1 & " where NoAccount='" & fCOA & "'"
            ExecMe a
        End If
    End If
    CommitTransaction
    LoadedCOA = False
    MsgBox "SUKSES"
    DoEvents
    fCOA_LostFocus
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Function DelZero(ByVal tNo As String) As String
    DelZero = tNo
    While Right(DelZero, 1) = "0"
        DelZero = Left(DelZero, Len(DelZero) - 1)
    Wend
End Function
Private Function RZerofill(ByVal tNo As String, ByVal tLen As Byte) As String
    RZerofill = tNo
    While Len(RZerofill) < tLen
        RZerofill = RZerofill & "0"
    Wend
End Function

Private Sub fCOA_LostFocus()
On Error Resume Next
    x.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    x.DeleteRows 0
    If fCOA.ListIndex = -1 Or fCOA = "" Then
        fDeskripsi = ""
        TDBGrid1.Rebind
    Else
        fCOA = fCOA.GetData("No Account")
        fDeskripsi = fCOA.GetData("Deskripsi")
    End If
    If fCOA = "" Then
        a = "select left(NoAccount,len(NoAccount)-6), Level1, Level2, Deskripsi,Saldo,NormalDK from m_ChartAccount where LevelDiatas='' order by Level1, Level2, Level3, Level4, Level5, Level6"
    Else
        tStart = Len(DelZero(fCOA)) + 1
        a = "select Level1,Level2,Level3,Level4,Level5,Level6,Deskripsi,Saldo,NormalDK from m_ChartAccount "
        'If pDatabaseServer = "SQL Server" Then
            'a = "select SubString(NoAccount," & tStart & ", len(NoAccount)-" & tStart & "), Deskripsi, Saldo, NormalDK from m_ChartAccount"
        'ElseIf pDatabaseServer = "Access" Then
        '    a = "select mid(NoAccount," & tStart & "), Deskripsi, Saldo, NormalDK from m_ChartAccount "
        'End If
        a = a & "where LevelDiatas='" & esc(fCOA) & "' order by Level1, Level2, Level3, Level4, Level5, Level6"
    End If
    query a
    If RS.RecordCount > 0 Then x.ReDim 0, RS.RecordCount - 1, 0, TDBGrid1.Columns.Count - 1
    For i = 0 To x.UpperBound(1)
        If RS.Fields("level2").Value = 0 Then
            x(i, 0) = RS.Fields("level1").Value
        ElseIf RS!level3 = 0 Then
            x(i, 0) = RS.Fields("level2").Value
        ElseIf RS!level4 = 0 Then
            x(i, 0) = RS.Fields("level3").Value
        ElseIf RS!level5 = 0 Then
            x(i, 0) = RS.Fields("level4").Value
        ElseIf RS!Level6 = 0 Then
            x(i, 0) = RS.Fields("level5").Value
        Else
            x(i, 0) = RS.Fields("level6").Value
        End If
        x(i, 1) = RS.Fields("Deskripsi").Value
        x(i, 2) = RS.Fields("Saldo").Value
        x(i, 3) = RS.Fields("NormalDK").Value
        RS.MoveNext
    Next
    TDBGrid1.Rebind
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
    fCOA.ZOrder 0
    LoadedCOA = False
    Set TDBGrid1.Array = x
End Sub

Private Sub fCOA_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    If Not LoadedCOA Then
        LoadedCOA = True
        fCOA.SetHeader "NO ACCOUNT@DESKRIPSI"
        fCOA.SetWidth "1500@2500"
        a = "select NoAccount,Deskripsi from m_ChartAccount order by NoAccount"
        query a
        If RS.RecordCount = 0 Then Exit Sub
        Dim rs1() As Variant
        rs1 = RS.GetRows
        fCOA.SetDB rs1
        fCOA.SetType "String"
    End If
err:
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    x.QuickSort 0, x.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_STRING
    TDBGrid1.Rebind
End Sub






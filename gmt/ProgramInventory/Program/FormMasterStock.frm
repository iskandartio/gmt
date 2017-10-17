VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMasterStock 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MASTER STOCK"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "20"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fNamaId 
      Caption         =   "NAMA &ID"
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame FrameFilter 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   1080
      Width           =   9615
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   6
         Left            =   4080
         TabIndex        =   24
         Tag             =   "SatBesar"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton fTambah 
         Caption         =   "+"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   5
         Left            =   8280
         TabIndex        =   20
         Tag             =   "Grade"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   4
         Left            =   6840
         TabIndex        =   18
         Tag             =   "Tube"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   3
         Left            =   5400
         TabIndex        =   16
         Tag             =   "NoWarna"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   15
         Tag             =   "Warna"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Tag             =   "KodeBarang"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Tag             =   "Jenis"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         Height          =   255
         Left            =   4080
         TabIndex        =   25
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No Warna"
         Height          =   255
         Left            =   5400
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
         Height          =   255
         Left            =   8280
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tube"
         Height          =   255
         Left            =   6840
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox fFind 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
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
      _StyleDefs(30)  =   "Named:id=29:Normal"
      _StyleDefs(31)  =   ":id=29,.parent=0"
      _StyleDefs(32)  =   "Named:id=30:Heading"
      _StyleDefs(33)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(34)  =   ":id=30,.wraptext=-1"
      _StyleDefs(35)  =   "Named:id=31:Footing"
      _StyleDefs(36)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(37)  =   "Named:id=32:Selected"
      _StyleDefs(38)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(39)  =   "Named:id=33:Caption"
      _StyleDefs(40)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(41)  =   "Named:id=34:HighlightRow"
      _StyleDefs(42)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(43)  =   "Named:id=35:EvenRow"
      _StyleDefs(44)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(45)  =   "Named:id=36:OddRow"
      _StyleDefs(46)  =   ":id=36,.parent=29"
      _StyleDefs(47)  =   "Named:id=39:RecordSelector"
      _StyleDefs(48)  =   ":id=39,.parent=30"
      _StyleDefs(49)  =   "Named:id=42:FilterBar"
      _StyleDefs(50)  =   ":id=42,.parent=29"
   End
   Begin VB.CommandButton fGrade 
      Caption         =   "&GRADE"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton fTube 
      Caption         =   "&TUBE"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton fNoWarna 
      Caption         =   "&NO WARNA"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton fKode 
      Caption         =   "&KODE"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton fJenis 
      Caption         =   "&JENIS"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton fStock 
      Caption         =   "&STOCK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Find"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "FormMasterStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim mTitle As String
Dim mCancelUpdate As Boolean
Dim col1 As TrueOleDBGrid80.Columns

Private Sub Form_Activate()
On Error Resume Next
    If Not cekValid("MASUK", Tag) Then Unload Me
End Sub


Private Sub fFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim cCol As Integer
        cCol = TDBGrid1.Col
        m = TDBGrid1.Bookmark
        n = m + 1
        If n = x.UpperBound(1) + 1 Then n = 0
        Do While m <> n
            If InStr(1, x(n, cCol), fFind, vbTextCompare) <> 0 Then Exit Do
            n = n + 1
            If n = x.UpperBound(1) + 1 Then n = 0
        Loop
        If InStr(1, x(n, cCol), fFind, vbTextCompare) = 0 Then
            MsgBox "Not Found"
        End If
        TDBGrid1.Bookmark = n
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Caption = Caption & "---" & pTipe
    Set TDBGrid1.Array = x
    Set col1 = TDBGrid1.Columns
    TDBGrid1.FetchRowStyle = True
    FrameFilter.Visible = False
End Sub

Sub DoQuery()
'On Error Resume Next
    If mTitle = "KODE" Then
        fKode_Click
    ElseIf mTitle = "NOWARNA" Then
        fNoWarna_Click
    ElseIf mTitle = "JENIS" Then
        fJenis_Click
    ElseIf mTitle = "STOCK" Then
        fStock_Click
    ElseIf mTitle = "NAMA ID" Then
        fNamaId_Click
    ElseIf mTitle = "TUBE" Then
        fTube_Click
    ElseIf mTitle = "GRADE" Then
        fGrade_Click
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left * 2
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 200
End Sub

Private Sub fStock_Click()
    TDBGrid1.Bookmark = 0
    TDBGrid1.Columns.Clear
    MyFilter = ""
    For i = 0 To Texts.Count - 1
        If Trim(Texts(i)) <> "" Then MyFilter = MyFilter & " and " & Texts(i).Tag & " like '%" & Texts(i) & "%'"
    Next
    a = "select 0 as updated,IsActive,Jenis,Warna,KodeBarang,NoWarna,Tube,Grade,SatBesar,SatKecil,JumlahBox,JumlahKG,IdStock from m_stock~ where 1=1 " & MyFilter & " or jenis='' order by Jenis, IdJenis, IdKodeBarang, d, f, m_stock~.KodeBarang,Idgrade,warnadasar,nowarna,idtube"
    query a
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 0 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1("updated").Visible = False
    col1("JumlahBox").Locked = True
    col1("JumlahKG").Locked = True
    col1("IdStock").Locked = True
    col1("IsActive").Width = 500
    col1("Jenis").Width = 700
    col1("Warna").Width = 700
    col1("NoWarna").Width = 1000
    col1("Tube").Width = 800
    col1("Grade").Width = 700
    col1("SatBesar").Width = 1000
    col1("SatKecil").Width = 1000
    col1("JumlahBox").Width = 900
    col1("JumlahKG").Width = 900
    col1("IdStock").Width = 700
    col1("JumlahBox").Tag = "Integer"
    col1("JumlahKG").Tag = "Decimal"
    col1("IsActive").Tag = "OK?"
    TDBGridLoad TDBGrid1
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    mTitle = "STOCK"
    FrameFilter.Visible = True
End Sub

Private Sub fJenis_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, Jenis as pKey, Jenis, IdJenis from m_stock~ where IsActive=1 order by IdJenis"
    query a
    col1.Clear
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 1 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1(0).Visible = False
    col1(1).Visible = False
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    mTitle = "JENIS"
    FrameFilter.Visible = False
End Sub

Private Sub fTambah_Click()
On Error GoTo err
    TDBGrid1.Update
    a = "select Max(IdStock) from m_stock" & pTipe
    query a
    If Not IsNull(RS.Fields(0).Value) Then IdStock = RS.Fields(0).Value + 1 Else IdStock = 1
    a = "insert into m_stock" & pTipe & "(IdStock, Jenis) values(" & IdStock & ",'')"
    If ExecMe(a) = 0 Then GoTo err
    x.AppendRows
    x(x.UpperBound(1), col1("IdStock").ColIndex) = IdStock
    For i = 0 To col1("JumlahBox").ColIndex - 1
        x(x.UpperBound(1), i) = x(TDBGrid1.Bookmark, i)
    Next
    x(x.UpperBound(1), col1("updated").ColIndex) = 1
    TDBGrid1.Rebind
    TDBGrid1.MoveLast
    TDBGrid1.SetFocus
err:
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
Dim IdJenis As Long
Dim IdTube As Long
Dim IdGrade As Long
Dim d As Long
Dim f As Long
Dim WarnaDasar As Long
Dim KetWarna As String
Dim KonvKG As Double
Dim Jenis As String
Dim Warna As String
Dim Tube As String
Dim Grade As String
Dim Kode As String
Dim NoWarna As String
Dim JumlahBox As Long
Dim jumlahKG As Double
Dim SatBesar As String
Dim SatKecil As String
Dim NamaId As String

    BeginTransaction
    TDBGrid1.Update
    idkodebarang = 0
    For i = 0 To x.UpperBound(1)
        If CLng(x(i, 0)) = 1 Then
            If mTitle = "STOCK" Then
                b = vbYes
                a = "select top 1 NoSPP, IdStock from t_SPPDetail~ where IdStock=" & x(i, col1("IdStock").ColIndex)
                query a
                If RS.RecordCount > 0 Then
                    b = MsgBox("Stock" & x(i, col1("IdStock").ColIndex) & " sudah ada di " & RS.Fields(0).Value & "!!!, Yakin mau diganti?", vbYesNo)
                End If
                a = "select top 1 IdStock from t_InputStock~ where IdStock=" & x(i, col1("IdStock").ColIndex)
                query a
                If RS.RecordCount > 0 Then
                    b = MsgBox("Stock" & x(i, col1("IdStock").ColIndex) & " sudah ada di Input Stock!!!, Yakin mau diganti?", vbYesNo)
                End If
                If b = vbYes Then
                    Jenis = x(i, col1("Jenis").ColIndex)
                    Warna = x(i, col1("Warna").ColIndex) & ""
                    Tube = x(i, col1("Tube").ColIndex) & ""
                    Grade = x(i, col1("Grade").ColIndex) & ""
                    Kode = x(i, col1("KodeBarang").ColIndex) & ""
                    NoWarna = x(i, col1("NoWarna").ColIndex) & ""
                    Kode = x(i, col1("KodeBarang").ColIndex) & ""
                    NoWarna = x(i, col1("NoWarna").ColIndex) & ""
                    JumlahBox = x(i, col1("JumlahBox").ColIndex)
                    jumlahKG = x(i, col1("JumlahKG").ColIndex)
                    SatBesar = x(i, col1("SatBesar").ColIndex) & ""
                    SatKecil = x(i, col1("SatKecil").ColIndex) & ""
                'IdJenis
                    a = "select top 1 IdJenis from m_Stock~ where Jenis='" & esc(x(i, col1("Jenis").ColIndex)) & "'"
                    query a
                    If RS.EOF Then
                        a = "select max(IdJenis)+1 from m_Stock~"
                        query a
                    End If
                    IdJenis = RS.Fields(0).Value
                'D,F, IdKodeBarang, KonvKG, NamaId
                    c = Trim(x(i, col1("KodeBarang").ColIndex))
                    Warna = Trim(x(i, col1("Warna").ColIndex))
                    SatBesar = x(i, col1("SatBesar").ColIndex)
                    a = "select top 1 IdKodeBarang, D, F, KonversiKG, NamaID from m_stock~ where KodeBarang='" & esc(c) & "' and Warna='" & esc(Warna) & "' and SatBesar='" & esc(SatBesar) & "'"
                    query a
                    If RS.EOF Then
                        idkodebarang = 1000
                        Dim c1() As Byte
                        c1 = StrConv(c, vbFromUnicode)
                        b = 0
                        If c <> "" Then
                            b1 = 0
                            While (c1(b) <= 47 Or c1(b) >= 58) And b <= UBound(c1)
                                b1 = b1 + 1
                                b = b + 1
                            Wend
                            While b <= UBound(c1) And c1(b) > 47 And c1(b) < 58
                                b = b + 1
                            Wend
                        End If
                        If b <> 0 Then
                            d = Mid(c, b1 + 1, b - b1)
                            If IsNumeric(Mid(c, b + 1)) Then
                                e = StrConv(Mid(c, b + 1), vbFromUnicode)
                                j = 0
                                g = ""
                                Do While e(j) > 47 And e(j) < 58
                                    g = g & Chr(e(j))
                                    If j = UBound(e) Then Exit Do
                                    j = j + 1
                                Loop
                                f = CLng(g)
                            Else
                                f = 0
                            End If
                        Else
                            h = MsgBox("Format D dan F tidak tepat!!! Lanjutkan?", vbYesNo)
                            If h = vbNo Then GoTo err
                        End If
                    Else
                        d = RS.Fields("D").Value
                        f = RS.Fields("F").Value
                        idkodebarang = RS.Fields("IdKodeBarang").Value
                        KonvKG = 1
                        KonvKG = RS.Fields("KonversiKG").Value
                        NamaId = RS.Fields("NamaID").Value & ""
                    End If
                'Warna Dasar, Ket Warna
                    c = Trim(x(i, col1("NoWarna").ColIndex))
                    If c <> "" Then
                        a = "select top 1 WarnaDasar, KetWarna from m_stock~ where NoWarna='" & esc(c) & "'"
                        query a
                        If RS.EOF Then
                            c1 = StrConv(c, vbFromUnicode)
                            j = 0
                            g = ""
                            Do While c1(j) > 47 And c1(j) < 58
                                g = g & Chr(c1(j))
                                If j = UBound(c1) Then Exit Do
                                j = j + 1
                            Loop
                            If cek_integer(g) Then
                                WarnaDasar = CLng(g)
                            Else
                                h = MsgBox("Warna Dasar tidak tepat!!! Lanjutkan?", vbYesNo)
                                If h = vbNo Then GoTo err
                            End If
                            KetWarna = ""
                        Else
                            WarnaDasar = RS.Fields(0).Value & ""
                            KetWarna = RS.Fields(1).Value & ""
                        End If
                    End If
                'IdTube
                    a = "select top 1 IdTube from m_stock~ where Tube='" & esc(x(i, col1("Tube").ColIndex)) & "'"
                    query a
                    If RS.EOF Then
                        a = "select max(IdTube)+1 from m_stock~"
                        query a
                        If Not IsNull(RS.Fields(0).Value) Then IdTube = RS.Fields(0).Value Else IdTube = 1
                    Else
                        IdTube = RS.Fields(0).Value
                    End If
                'IdGrade
                    a = "select top 1 IdGrade from m_stock~ where Grade='" & esc(x(i, col1("Grade").ColIndex)) & "'"
                    query a
                    If RS.EOF Then
                        a = "select max(IdGrade)+1 from m_stock~"
                        query a
                        If Not IsNull(RS.Fields(0).Value) Then IdGrade = RS.Fields(0).Value Else IdGrade = 1
                    Else
                        IdGrade = RS.Fields(0).Value
                    End If
                    a = "select top 1 jenis from m_Stock~ where SatBesar='" & esc(x(i, col1("SatBesar").ColIndex)) & "' and IdStock<>" & x(i, col1("IdStock").ColIndex) & " and NoWarna='" & esc(NoWarna) & "' and KodeBarang='" & esc(Kode) & "' and Warna='" & esc(Warna) & "' and grade='" & esc(Grade) & "' and tube='" & esc(Tube) & "' and Jenis='" & esc(Jenis) & "' and jenis<>''"
                    query a
                    If RS.RecordCount = 1 Then
                        MsgBox "Nama Stock Sudah ada!!!(" & x(i, col1("IdStock").ColIndex) & ")"
                        GoTo err
                    End If
                    a = "update m_stock~ set IsActive=" & IIf(x(i, col1("IsActive").ColIndex) = 0, 0, 1) & ", NamaId='" & esc(NamaId) & "', KonversiKG=" & cNum(KonvKG, 4) & ",JumlahBox=" & cNum(JumlahBox) & ",JumlahKG=" & cNum(jumlahKG) & ", Jenis='" & esc(Jenis) & "', Warna='" & esc(Warna) & "', IdKodeBarang=" & idkodebarang & ", KodeBarang='" & esc(Kode) & "',NoWarna='" & esc(NoWarna) & "',Tube='" & esc(Tube) & "',Grade='" & esc(Grade) & "',SatBesar='" & esc(SatBesar) & "',SatKecil='" & esc(SatKecil) & "',IdJenis=" & IdJenis & ",D=" & d & ",F=" & f & ",WarnaDasar=" & WarnaDasar & ",KetWarna='" & esc(KetWarna) & "', IdTube=" & IdTube & ",IdGrade=" & IdGrade & " where IdStock=" & x(i, col1("IdStock").ColIndex)
                End If
            ElseIf mTitle = "JENIS" Then
                c = x(i, col1("pKey").ColIndex)
                Jenis = x(i, col1("Jenis").ColIndex)
                IdJenis = x(i, col1("IdJenis").ColIndex)
                If c <> Jenis Then
                    a = "select top 1 IdJenis from m_stock~ where Jenis='" & esc(Jenis) & "'"
                    query a
                    If Not RS.EOF Then
                        b = MsgBox("Jenis sudah Ada!!! Lanjutkan?", vbYesNo)
                        If b = vbNo Then GoTo err
                    End If
                    IdJenis = RS.Fields(0).Value
                End If
                a = "update m_stock~ set Jenis='" & esc(Jenis) & "',IdJenis=" & IdJenis & " where Jenis='" & esc(c) & "" & "'"
            ElseIf mTitle = "TUBE" Then
                c = x(i, col1("pKey").ColIndex)
                Tube = x(i, col1("Tube").ColIndex)
                IdTube = x(i, col1("IdTube").ColIndex)
                If c <> Tube Then
                    a = "select top 1 IdTube from m_stock~ where Tube='" & esc(Tube) & "'"
                    query a
                    If Not RS.EOF Then
                        b = MsgBox("Tube sudah Ada!!! Lanjutkan?", vbYesNo)
                        If b = vbNo Then GoTo err
                        IdTube = RS.Fields(0).Value
                    End If
                End If
                a = "update m_stock~ set Tube='" & esc(Tube) & "',IdTube=" & IdTube & " where Tube='" & esc(c) & "'"
            ElseIf mTitle = "GRADE" Then
                c = x(i, col1("pKey").ColIndex) & ""
                Grade = x(i, col1("Grade").ColIndex) & ""
                IdGrade = x(i, col1("IdGrade").ColIndex)
                If c <> Grade Then
                    a = "select top 1 IdGrade from m_stock~ where Grade='" & esc(Grade) & "'"
                    query a
                    If Not RS.EOF Then
                        b = MsgBox("Grade sudah Ada!!! Lanjutkan?", vbYesNo)
                        If b = vbNo Then GoTo err
                    End If
                    IdGrade = RS.Fields(0).Value
                End If
                a = "update m_stock~ set Grade='" & esc(Grade) & "',IdGrade=" & IdGrade & " where Grade='" & esc(c) & "'"
            ElseIf mTitle = "KODE" Then
                c = x(i, col1("pKey").ColIndex)
                c2 = x(i, col1("pKey2").ColIndex)
                c3 = x(i, col1("pKey3").ColIndex)
                d = x(i, col1("D").ColIndex)
                f = x(i, col1("F").ColIndex)
                Kode = x(i, col1("KodeBarang").ColIndex)
                Warna = x(i, col1("Warna").ColIndex)
                SatBesar = x(i, col1("SatBesar").ColIndex)
                If c <> Kode Then
                    a = "select top 1 d, f from m_Stock~ where KodeBarang='" & esc(Kode) & "' and Warna='" & esc(Warna) & "' and SatBesar='" & esc(SatBesar) & "'"
                    query a
                    If Not RS.EOF Then
                        b = MsgBox("Kode Barang Sudah Ada!!! Lanjutkan?", vbYesNo)
                        If b = vbNo Then GoTo err
                        d = RS.Fields(0).Value
                        f = RS.Fields(1).Value
                    End If
                End If
                Dim konversiKg As Double
                konversiKg = cNum(x(i, col1("KonversiKG").ColIndex), 4)
                If konversiKg = 0 Then konversiKg = 1
                a = "update m_Stock~ set KonversiKG=" & konversiKg & ", SatKecil='" & esc(x(i, col1("SatKecil").ColIndex)) & "', SatBesar='" & esc(x(i, col1("SatBesar").ColIndex)) & "', Warna='" & esc(x(i, col1("Warna").ColIndex)) & "',IdKodeBarang=" & x(i, col1("IdKodeBarang").ColIndex) & ", KodeBarang='" & esc(Kode) & "', d=" & d & ",f=" & f & " where KodeBarang='" & esc(c) & "' and Warna='" & esc(c2) & "' and SatBesar='" & esc(c3) & "'"
            ElseIf mTitle = "NOWARNA" Then
                c = x(i, col1("pKey").ColIndex)
                NoWarna = x(i, col1("NoWarna").ColIndex)
                WarnaDasar = x(i, col1("WarnaDasar").ColIndex)
                KetWarna = x(i, col1("KetWarna").ColIndex) & ""
                If c <> NoWarna Then
                    a = "select top 1 WarnaDasar, KetWarna from m_Stock~ where NoWarna='" & esc(NoWarna) & "'"
                    query a
                    If Not RS.EOF Then
                        b = MsgBox("No Warna sudah ada!! Lanjutkan?", vbYesNo)
                        If b = vbNo Then GoTo err
                        WarnaDasar = RS.Fields(0).Value & ""
                        KetWarna = RS.Fields(1).Value & ""
                    End If
                End If
                a = "update m_Stock~ set NoWarna='" & esc(NoWarna) & "',WarnaDasar=" & WarnaDasar & ", KetWarna='" & esc(KetWarna) & "' where NoWarna='" & esc(c) & "'"
            ElseIf mTitle = "SATUAN" Then
                KonvKG = x(i, col1("KonversiKG").ColIndex)
                a = "update m_Stock~ set KonversiKG=" & cNum(KonvKG) & " where SatKecil='" & esc(x(i, col1("SatKecil").ColIndex)) & "'"
            ElseIf mTitle = "NAMA ID" Then
                a = "update m_Stock~ set NamaId='" & esc(x(i, col1("NamaId").ColIndex)) & "' where IdKodeBarang=" & cNum(x(i, col1("IdKodeBarang").ColIndex))
            End If
            If ExecMe(a) = 0 Then
                GoTo err
            End If
        End If
    Next
    CommitTransaction
    MsgBox "SUKSES"
    DoEvents
    DoQuery
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub fKode_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, KodeBarang as pKey, Warna as pKey2, SatBesar as pKey3, IdKodeBarang, KodeBarang, Warna, D, F, SatBesar, SatKecil, KonversiKG from m_stock~ where IsActive=1 order by IdKodeBarang"
    query a
    col1.Clear
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 3 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    mTitle = "KODE"
    col1("IdKodeBarang").Width = 1000
    col1("Warna").Width = 1000
    col1("D").Width = 1000
    col1("F").Width = 1000
    col1("KonversiKG").Tag = "Decimal"
    FrameFilter.Visible = False
End Sub

Private Sub fNoWarna_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, NoWarna as pKey, NoWarna, WarnaDasar, KetWarna from m_stock~ where IsActive=1 order by WarnaDasar,NoWarna"
    query a
    col1.Clear
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 1 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1(0).Visible = False
    col1(1).Visible = False
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    mTitle = "NOWARNA"
    FrameFilter.Visible = False
End Sub

Private Sub fTube_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, Tube as pKey, Tube, IdTube from m_stock~ where IsActive=1 order by IdTube"
    query a
    col1.Clear
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 1 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1(0).Visible = False
    col1(1).Visible = False
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    mTitle = "TUBE"
    FrameFilter.Visible = False
End Sub

Private Sub fGrade_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, Grade as pKey, Grade, IdGrade from m_stock~ where IsActive=1 order by IdGrade"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
col1.Clear
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 1 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    mTitle = "GRADE"
    FrameFilter.Visible = False
End Sub

Private Sub fNamaId_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, IdKodeBarang, NamaId  from m_stock~ where IsActive=1 order by IdKodeBarang"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    col1.Clear
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 0 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    TDBGridSetVisible TDBGrid1, "updated"
    TDBGridSetLock TDBGrid1, "IdKodeBarang", True
    TDBGridLoad TDBGrid1
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    mTitle = "NAMA ID"
    FrameFilter.Visible = False
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    col1("updated").Value = "1"
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(Bookmark, col1("updated").ColIndex) = "1" Then RowStyle.BackColor = vbYellow
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        fFind_KeyDown 13, 0
    Else
        TDBGridKeyDown TDBGrid1, KeyCode
    End If
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_MouseDown(Button As Integer, Shift As Integer, tX As Single, tY As Single)
On Error Resume Next
    a = ActiveControl.Name
    If a <> "TDBGrid1" Then
        If IsNull(LastRow) Then TDBGrid1.Row = 0
        If x.UpperBound(1) > 0 Then TDBGrid1.SetFocus
    End If
End Sub

Private Sub Texts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DoQuery
    End If
End Sub



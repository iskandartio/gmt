VERSION 5.00
Object = "{81BF1356-E605-4EBC-9514-354F6147449F}#1.0#0"; "usrtext.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FormMasterStockBeli 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MASTER STOCK BELI"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Tag             =   "22"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fGabung 
      Caption         =   "&GABUNG"
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
   Begin UsrText.IText fIdHasil 
      Height          =   270
      Left            =   5160
      TabIndex        =   19
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
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
   Begin UsrText.IText fId 
      Height          =   270
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
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
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton fUpdate 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7011
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
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
      AllowDelete     =   -1  'True
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
   Begin VB.CommandButton fSupplier 
      Caption         =   "SUPPLIER"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame FrameFilter 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   2
         Left            =   4080
         TabIndex        =   12
         Tag             =   "NamaBarang"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Tag             =   "SuppSuggestion"
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton fTambah 
         Caption         =   "+"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Texts 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Tag             =   "Jenis"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox fFind 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton fJenis 
      Caption         =   "&JENIS"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton fStock 
      Caption         =   "&STOCK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gabung Stock"
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label fKet 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   9735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Find"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "FormMasterStockBeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Dim mTitle As String
Dim mKoma As Boolean
Dim col1 As TrueOleDBGrid80.Columns

Private Sub fDelete_Click()
    If mTitle <> "STOCK" Then
        MsgBox "Hapus Lewat Data Stock!!!"
        Exit Sub
    End If
    b = MsgBox("Yakin mau hapus?", vbYesNo)
    If b = vbNo Then Exit Sub
    a = "select top 1 NoPR from t_PR where IdStock=" & col1("IdStock").Value
    query a
    If RS.RecordCount > 0 Then
        MsgBox "Barang sudah ada di No PR : " & RS.Fields(0).Value
        Exit Sub
    Else
        a = "update m_stockBeli set IsActive=0, KodeJenis=0, Jenis='', NamaBarang='', Satuan='', Jumlah=0, TanggalTerakhir=0, HargaTerakhir=0, SuppSuggestion='', MataUang='', WaktuPembayaran='', PPNYesNo=0 where IdStock=" & fId
        If ExecMe(a) = 0 Then Exit Sub
    End If
    TDBGrid1.Delete
End Sub

Private Sub fGabung_Click()
On Error GoTo err
    BeginTransaction
    a = "select Jumlah from m_StockBeli where IdStock=" & fId
    query a
    If RS.RecordCount = 0 Then GoTo err
    n = RS.Fields(0).Value
    a = "update m_stockBeli set Jumlah=Jumlah+" & cNum(n) & " where IdStock=" & fIdHasil
    If ExecMe(a) = 0 Then GoTo err
    a = "update m_stockBeli set IsActive=0, KodeJenis=0, Jenis='', NamaBarang='', Satuan='', Jumlah=0, TanggalTerakhir=0, HargaTerakhir=0, SuppSuggestion='', MataUang='', WaktuPembayaran='', PPNYesNo=0 where IdStock=" & fId
    If ExecMe(a) = 0 Then GoTo err
    a = "update t_PR set IdStock=" & fIdHasil & " where IdStock=" & fId
    ExecMe a
    a = "update t_BTB set IdStock=" & fIdHasil & " where IdStock=" & fId
    ExecMe a
    a = "update t_NRPembelianDetail set IdStock=" & fIdHasil & " where IdStock=" & fId
    ExecMe a
    a = "update m_StockWaste set IdStock=" & fIdHasil & " where IdStock=" & fId
    ExecMe a
    a = "update m_StockProses set IdStock=" & fIdHasil & " where IdStock=" & fId
    ExecMe a
    CommitTransaction
    MsgBox "SUKSES"
    fStock_Click
    Exit Sub
err:
    CommitTransaction
    MsgBox "GAGAL"
End Sub

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
    Set col1 = TDBGrid1.Columns
    Set TDBGrid1.Array = x
    TDBGrid1.AllowAddNew = False
    FrameFilter.Visible = False
    TDBGrid1.FetchRowStyle = True
End Sub

Private Function MyFilter() As String
    MyFilter = ""
    For i = 0 To Texts.Count - 1
        MyFilter = MyFilter & IIf(esc(Texts(i)) = "", "", " and " & Texts(i).Tag & " like '%" & Texts(i) & "%'")
    Next
    If MyFilter <> "" Then MyFilter = " where " & Mid(MyFilter, 6)
End Function

Sub DoQuery()
On Error Resume Next
    a = "select 0, IsActive, IdStock, KodeJenis, Jenis, NamaBarang, Satuan, Jumlah, HargaTerakhir, SuppSuggestion, MataUang, WaktuPembayaran, PPNYesNo from m_stockBeli " & MyFilter & " order by KodeJenis, NamaBarang"
    query a
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    TDBGrid1.MoveFirst
    TDBGrid1.SetFocus
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TDBGrid1.Width = ScaleWidth - TDBGrid1.Left * 2
    fKet.Width = TDBGrid1.Width
    TDBGrid1.Height = ScaleHeight - TDBGrid1.Top - 200
End Sub

Private Sub fStock_Click()
    TDBGrid1.Bookmark = 0
    While col1.Count <> 0
        col1.Remove 0
    Wend
    a = "select 0 as updated, IsActive, IdStock, KodeJenis, Jenis, NamaBarang, Satuan, Jumlah, HargaTerakhir, SuppSuggestion, MataUang, WaktuPembayaran, PPNYesNo from m_stockBeli " & MyFilter & " order by KodeJenis, NamaBarang"
    query a
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 0 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    col1("IdStock").Locked = True
    col1("IsActive").Tag = "OK?"
    col1("Jumlah").Tag = "Integer"
    col1("HargaTerakhir").Tag = "Decimal"
    TDBGridLoad TDBGrid1
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    For i = 0 To col1.Count - 1
        col1(i).AutoSize
    Next
    mTitle = "STOCK"
    FrameFilter.Visible = True
End Sub

Private Sub fJenis_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, KodeJenis as pKey, KodeJenis, Jenis from m_stockBeli where IsActive=1 order by KodeJenis, Jenis"
    query a
    While col1.Count <> 0
        col1.Remove 0
    Wend
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 1 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    For i = 0 To col1.Count - 1
        col1(i).AutoSize
    Next
    mTitle = "JENIS"
    FrameFilter.Visible = False
End Sub


Private Sub fSupplier_Click()
    TDBGrid1.Bookmark = 0
    a = "select distinct 0 as updated, SuppSuggestion as pKey,SuppSuggestion, MataUang, WaktuPembayaran, PPNYesNo from m_stockBeli where IsActive=1 order by SuppSuggestion"
    query a
    While col1.Count <> 0
        col1.Remove 0
    Wend
    For i = 0 To RS.Fields.Count - 1
        col1.Add(i).Caption = RS.Fields(i).Name
        If i > 1 Then col1(i).Visible = True Else col1(i).AllowSizing = False
    Next
    x.ReDim 0, 0, 0, col1.Count - 1
    x.DeleteRows 0
    If Not RS.EOF Then x.LoadRows RS.GetRows
    TDBGrid1.Rebind
    For i = 0 To col1.Count - 1
        col1(i).AutoSize
    Next
    mTitle = "SUPPLIER"
    FrameFilter.Visible = False
End Sub

Private Sub fTambah_Click()
On Error GoTo err
    TDBGrid1.Update
    a = "select Max(IdStock) from m_stockBeli"
    query a
    If Not IsNull(RS.Fields(0).Value) Then IdStock = RS.Fields(0).Value + 1 Else IdStock = 1
    a = "insert into m_stockBeli(IdStock,NamaBarang) values(" & IdStock & ",'')"
    If ExecMe(a) = 0 Then GoTo err
    x.AppendRows
    For i = 0 To x.UpperBound(2)
        x(x.UpperBound(1), i) = x(TDBGrid1.Bookmark, i)
    Next
    x(x.UpperBound(1), col1("IdStock").ColIndex) = IdStock
    x(x.UpperBound(1), col1("updated").ColIndex) = 1
    TDBGrid1.Rebind
    TDBGrid1.MoveLast
err:
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
    col1("updated").Value = 1
End Sub

Private Sub fUpdate_Click()
On Error GoTo err
    BeginTransaction
    TDBGrid1.Update
    b = vbNo
    If mTitle = "STOCK" Then
        iUpdated = col1("updated").ColIndex
        For i = 0 To x.UpperBound(1)
            If x(i, iUpdated) = 1 Then
                a = "select top 1 IdStock from t_PR where IdStock=" & x(i, col1("IdStock").ColIndex)
                query a
                If RS.RecordCount > 0 Then
                    If b = vbNo Then b = MsgBox(RS.Fields(0).Value & " Sudah Pernah ada di Pembelian!!!, Yakin mau Ganti Data tanpa Konfirmasi?", vbQuestion + vbYesNoCancel)
                    If b = vbCancel Then GoTo err
                End If
                a = "update m_stockBeli set " & _
                    "IsActive=" & IIf(x(i, col1("IsActive").ColIndex) = 0, 0, 1) & _
                    ", KodeJenis=" & x(i, col1("KodeJenis").ColIndex) & _
                    ",Jenis='" & esc(x(i, col1("Jenis").ColIndex)) & _
                    "',NamaBarang='" & esc(x(i, col1("NamaBarang").ColIndex)) & _
                    "',Satuan='" & esc(x(i, col1("Satuan").ColIndex)) & _
                    "',Jumlah=" & cNum(x(i, col1("Jumlah").ColIndex)) & _
                    ",HargaTerakhir=" & cNum(x(i, col1("HargaTerakhir").ColIndex)) & _
                    ",SuppSuggestion='" & esc(x(i, col1("SuppSuggestion").ColIndex) & "") & _
                    "',MataUang='" & esc(x(i, col1("MataUang").ColIndex) & "") & _
                    "',WaktuPembayaran='" & esc(x(i, col1("WaktuPembayaran").ColIndex) & "") & _
                    "',PPNYesNo=" & x(i, col1("PPNYesNo").ColIndex) & _
                    " where IdStock=" & x(i, col1("IdStock").ColIndex)
                If ExecMe(a) = 0 Then GoTo err
            End If
        Next
    ElseIf mTitle = "JENIS" Then
        For i = 0 To x.UpperBound(1)
            If x(i, col1("updated").ColIndex) = 1 Then
                a = "update m_stockBeli set KodeJenis=" & x(i, col1("KodeJenis").ColIndex) & ", Jenis='" & esc(x(i, col1("Jenis").ColIndex)) & "' where KodeJenis=" & x(i, col1("pKey").ColIndex)
                ExecMe a
            End If
        Next
    ElseIf mTitle = "SUPPLIER" Then
        For i = 0 To x.UpperBound(1)
            If x(i, col1("updated").ColIndex) = 1 Then
                a = "update m_stockBeli set SuppSuggestion='" & esc(x(i, col1("SuppSuggestion").ColIndex)) & _
                    "', MataUang='" & esc(x(i, col1("MataUang").ColIndex)) & _
                    "', WaktuPembayaran='" & esc(x(i, col1("WaktuPembayaran").ColIndex)) & _
                    "', PPNYesNo=" & x(i, col1("PPNYesNo").ColIndex) & _
                    " where SuppSuggestion='" & esc(x(i, col1("pKey").ColIndex)) & "'"
                ExecMe a
            End If
        Next
    End If
    CommitTransaction
    MsgBox "SUKSES"
    fStock_Click
    Exit Sub
err:
    RollBackTransaction
    MsgBox "GAGAL"
End Sub

Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
    Response = False
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
    If x(Bookmark, col1("updated").ColIndex) = 1 Then RowStyle.BackColor = vbYellow
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    x.QuickSort 0, x.UpperBound(1), ColIndex, XTYPE_STRING, XORDER_ASCEND
    TDBGrid1.Rebind
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then fFind_KeyDown 13, 0
    TDBGridKeyDown TDBGrid1, KeyCode
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    TDBGridKeyPress TDBGrid1, KeyAscii
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    fKet = col1(TDBGrid1.Col).Value
End Sub

Private Sub Texts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DoQuery
End Sub


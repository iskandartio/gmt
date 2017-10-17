VERSION 5.00
Begin VB.Form FormMenu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MENU"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SETUP"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   6570
      Width           =   10455
      Begin VB.CommandButton fPindah 
         Caption         =   "PINDAH"
         Height          =   375
         Left            =   7320
         TabIndex        =   47
         Tag             =   "41"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton fPassword 
         Caption         =   "PASSWORD"
         Height          =   375
         Left            =   6120
         TabIndex        =   46
         Tag             =   "36"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton fGoOffline 
         Caption         =   "GO &OFFLINE"
         Height          =   375
         Left            =   4800
         TabIndex        =   45
         Tag             =   "37"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fSettingRekap 
         Caption         =   "SETTING REKAP"
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Tag             =   "38"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton fPenomoran 
         Caption         =   "PENOMORAN"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Tag             =   "39"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fMarginPrinter 
         Caption         =   "MARGIN PRINTER"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Tag             =   "40"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "STOCK ADJUSTMENT"
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   3600
      Width           =   10395
      Begin VB.CommandButton fDeleteAutoInput 
         Caption         =   "DELETE AUTO INPUT"
         Height          =   375
         Left            =   4320
         TabIndex        =   59
         Tag             =   "53"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton fStockRetur 
         Caption         =   "STOCK RETUR"
         Height          =   375
         Left            =   2760
         TabIndex        =   54
         Tag             =   "49"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton fMutasi 
         Caption         =   "MUTASI"
         Height          =   375
         Left            =   1680
         TabIndex        =   44
         Tag             =   "25"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton fStockOpname 
         Caption         =   "STOCK OPNAME"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Tag             =   "24"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "LAPORAN AKUNTING"
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   5250
      Width           =   10455
      Begin VB.CommandButton fLaporanKeuangan 
         Caption         =   "LAPORAN KEUANGAN"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Tag             =   "48"
         Top             =   660
         Width           =   1935
      End
      Begin VB.CommandButton fKurs 
         Caption         =   "KURS"
         Height          =   375
         Left            =   5640
         TabIndex        =   22
         Tag             =   "35"
         Top             =   660
         Width           =   1095
      End
      Begin VB.CommandButton fKartuPiutang 
         Caption         =   "KARTU PIUTANG"
         Height          =   375
         Left            =   3960
         TabIndex        =   23
         Tag             =   "34"
         Top             =   660
         Width           =   1695
      End
      Begin VB.CommandButton fRekapOP 
         Caption         =   "REKAP OPERASIONAL"
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Tag             =   "33"
         Top             =   660
         Width           =   1935
      End
      Begin VB.CommandButton fSettingAccounting 
         Caption         =   "SETTING ACCOUNTING"
         Height          =   375
         Left            =   5280
         TabIndex        =   40
         Tag             =   "32"
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton fJurnalPembayaran 
         Caption         =   "JURNAL PEMBAYARAN"
         Height          =   375
         Left            =   3240
         TabIndex        =   41
         Tag             =   "31"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton fJurnalUmum 
         Caption         =   "JURNAL UMUM"
         Height          =   375
         Left            =   1800
         TabIndex        =   42
         Tag             =   "30"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton fChartAccount 
         Caption         =   "CHART ACCOUNT"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Tag             =   "29"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TRANSAKSI PEMBELIAN"
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   10335
      Begin VB.CommandButton fPembelianHarian 
         Caption         =   "REKAP BTB"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Tag             =   "46"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton fNPB 
         Caption         =   "NPB"
         Height          =   375
         Left            =   8160
         TabIndex        =   26
         Tag             =   "18"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton fGiroSupplier 
         Caption         =   "GIRO SUPPLIER"
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Tag             =   "17"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton fPelunasanPembelian 
         Caption         =   "PPU"
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Tag             =   "16"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton fJasa 
         Caption         =   "JASA"
         Height          =   375
         Left            =   4920
         TabIndex        =   28
         Tag             =   "15"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton fReturPembelian 
         Caption         =   "RETUR PB"
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Tag             =   "14"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton fSetHargaBeli 
         Caption         =   "HARGA BELI"
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Tag             =   "13"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fBTB 
         Caption         =   "BTB"
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Tag             =   "12"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton fPO 
         Caption         =   "PO"
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Tag             =   "11"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton fPR 
         Caption         =   "PR"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Tag             =   "10"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PROSES STOCK"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   4410
      Width           =   10395
      Begin VB.CommandButton fInputStock 
         Caption         =   "INPUT STOCK GBJ"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Tag             =   "28"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton fInputProses 
         Caption         =   "VALIDASI PROSES"
         Height          =   375
         Left            =   1680
         TabIndex        =   38
         Tag             =   "27"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton fTukarStock 
         Caption         =   "PROSES STOCK"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Tag             =   "26"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "INPUT MASTER"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2730
      Width           =   10395
      Begin VB.CommandButton fKeterangan 
         Caption         =   "KET"
         Height          =   375
         Left            =   9360
         TabIndex        =   57
         Tag             =   "52"
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton fStockProses 
         Caption         =   "STOCK PROSES"
         Height          =   375
         Left            =   7800
         TabIndex        =   55
         Tag             =   "50"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton fStockWaste 
         Caption         =   "STOCK WASTE"
         Height          =   375
         Left            =   6480
         TabIndex        =   50
         Tag             =   "45"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fMasterSupplier 
         Caption         =   "SUPPLIER"
         Height          =   375
         Left            =   5400
         TabIndex        =   35
         Tag             =   "23"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton fStockBeli 
         Caption         =   "STOCK BELI"
         Height          =   375
         Left            =   4080
         TabIndex        =   34
         Tag             =   "22"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fCustomer 
         Caption         =   "CUSTOMER"
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Tag             =   "21"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fMasterStock 
         Caption         =   "STOCK GBJ"
         Height          =   375
         Left            =   1440
         TabIndex        =   32
         Tag             =   "20"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fMataUang 
         Caption         =   "MATA UANG"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Tag             =   "19"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TRANSAKSI PENJUALAN"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   10335
      Begin VB.CommandButton fRekapGiro 
         Caption         =   "REKAP GIRO"
         Height          =   375
         Left            =   8040
         TabIndex        =   58
         Tag             =   "47"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton fSummaryRetur 
         Caption         =   "SUMMARY RETUR"
         Height          =   375
         Left            =   3480
         TabIndex        =   56
         Tag             =   "51"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton fRekapKontrak 
         Caption         =   "REKAP KONTRAK"
         Height          =   375
         Left            =   6360
         TabIndex        =   52
         Tag             =   "47"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton fWaste 
         Caption         =   "WASTE"
         Height          =   375
         Left            =   6720
         TabIndex        =   49
         Tag             =   "44"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton fDP 
         Caption         =   "DP"
         Height          =   375
         Left            =   5880
         TabIndex        =   48
         Tag             =   "43"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton fRetur 
         Caption         =   "REKAP RETUR"
         Height          =   375
         Left            =   2040
         TabIndex        =   31
         Tag             =   "9"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton fLaporanSPP 
         Caption         =   "REKAP SPP"
         Height          =   375
         Left            =   5160
         TabIndex        =   29
         Tag             =   "8"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton fLapPenjualan 
         Caption         =   "LAPORAN PENJUALAN"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Tag             =   "6"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton fGiroCustomer 
         Caption         =   "GIRO CUSTOMER"
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Tag             =   "5"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton fPelunasan 
         Caption         =   "PELUNASAN"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Tag             =   "4"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton fKwitansi 
         Caption         =   "KW"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Tag             =   "3"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton fSuratJalan 
         Caption         =   "SJ"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Tag             =   "2"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton fSPP 
         Caption         =   "SPP"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Tag             =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton fSC 
         Caption         =   "SC"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Tag             =   "0"
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fDeleteAutoInput_Click()
    FormDeleteAutoInput.Show
End Sub

Private Sub fKeterangan_Click()
    FormMasterKeterangan.Show
End Sub

Private Sub fLaporanKeuangan_Click()
    FormLaporanKeuangan.Show
End Sub

Private Sub fLaporanSPP_Click()
    fIntervalTanggal.LoadMe "LaporanSPP", False
End Sub

Private Sub fBTB_Click()
    FormBTB.Show
End Sub

Private Sub fChartAccount_Click()
    FormMasterAccount.Show
End Sub

Private Sub fDP_Click()
    FormDP.Show
End Sub

Private Sub fGiroSupplier_Click()
    FormGiroPembelian.Show
End Sub

'Sub fGoOffline_Click()
'    CopyFileAPI pServerDir & "\a.zip", "c:\a.zip"
'    UnZipDB "a.zip", App.Path
'    FormLogin.Client.Close
'    FormLogin.Timer1.Enabled = False
'    pServerName = App.Path
'    ConnectDatabase
'End Sub

Private Sub fInputProses_Click()
    FormValidasiProses.Show
End Sub

Private Sub fJasa_Click()
    FormPembelianJasa.Show
End Sub

Private Sub fJurnalPembayaran_Click()
    FormJurnalPembayaran.Show
End Sub

Private Sub fJurnalUmum_Click()
    FormJurnalUmum.Show
End Sub

Private Sub fMasterSupplier_Click()
    FormMasterSupplier.Show
End Sub

Private Sub fNPB_Click()
    FormNPB.Show
End Sub

Private Sub Form_Activate()
    Caption = pTipe & "-MENU"
End Sub


Private Sub fCustomer_Click()
   FormMasterCustomer.Show
End Sub

Private Sub fGiroCustomer_Click()
    FormGiroPenjualan.Show
End Sub

Private Sub fInputStock_Click()
    FormInputStock.Show
End Sub

Private Sub fKartuPiutang_Click()
    FormKartuPiutang.Show
End Sub

Private Sub fKurs_Click()
    FormKurs.Show
End Sub

Private Sub fKwitansi_Click()
    FormKW.Show
End Sub

Private Sub fLapPenjualan_Click()
    FormLaporanPenjualan.Show
End Sub

Private Sub fMarginPrinter_Click()
    FormMargin.Show
End Sub

Private Sub fMasterStock_Click()
    FormMasterStock.Show
End Sub

Private Sub fMataUang_Click()
    FormMasterMataUang.Show
End Sub

Private Sub fMutasi_Click()
    FormMutasi.Show
End Sub

Sub ProsesAutoUpdate()
'On Error Resume Next
    a = "select top 1 Tanggal from AutoUpdate where Nama='MENU'"
    query a
    Dim Tgl As Long
    Dim TglOld As Long
    Tgl = cD(pServerDate)
    TglOld = RS.Fields(0).Value
    If Tgl <> TglOld Then
        For i = 0 To Controls.count - 1
            If TypeName(Controls(i)) = "CommandButton" Then
                a = "update m_DaftarMenu set Tag=" & Controls(i).Tag & " where command='" & esc(Controls(i).Caption) & "'"
                If ExecMe(a) = 0 Then
                    a = "insert into m_DaftarMenu(Header, Command, Tag) values('" & _
                        esc(Controls(i).Container) & "','" & esc(Controls(i).Caption) & "'," & Controls(i).Tag & ")"
                    ExecMe a
                    End If
                
            End If
        Next
        a = "update AutoUpdate set Tanggal=" & Tgl & " where Nama='MENU'"
        ExecMe a
    End If
End Sub

Sub CekValidMenuControl()
    For i = 0 To Controls.count - 1
        If TypeName(Controls(i)) = "CommandButton" Then
            Controls(i).Enabled = cekValid("MASUK", Controls(i).Tag, True)
        End If
    Next
    If pOffLineMode Then fGoOffline.Enabled = False
End Sub

Private Sub Form_Load()
    ProsesAutoUpdate
    CekValidMenuControl
    pSettingName = GetSetting("GMT", "MarginPrinter", "Name")
    a = "select top 1 Kiri, Atas from m_MarginPrinter where Nama='" & esc(pSettingName) & "'"
    query a
    If RS.RecordCount > 0 Then
        pLeftMargin = RS.Fields("Kiri").Value
        pTopMargin = RS.Fields("Atas").Value
    End If
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For Each f In Forms
        Unload f
    Next
End Sub

Private Sub fPembelianHarian_Click()
    fIntervalTanggal.LoadMe "LaporanPembelianHarian"
End Sub

Private Sub fPindah_Click()
    For Each Form In Forms
        If Form.Name <> "FormMenu" And Form.Name <> "FormLogin" Then
            Unload Form
        End If
    Next
    FormPindah.Show
End Sub

Private Sub fPelunasan_Click()
    FormPelunasanPenjualan.Show
End Sub

Private Sub fPelunasanPembelian_Click()
    FormPelunasanPembelian.Show
End Sub

Private Sub fPenjualanSummary_Click()
    fIntervalTanggal.LoadMe "SummaryPenjualan"
End Sub

Private Sub fPenomoran_Click()
    FormPenomoran.Show
End Sub

Private Sub fPassword_Click()
    FormPassword.Show
End Sub

Private Sub fPO_Click()
    FormPO.Show
End Sub

Private Sub fPR_Click()
    FormPR.Show
End Sub

Private Sub fRekapGiro_Click()
    fIntervalTanggal.LoadMe "Giro"
End Sub

Private Sub fRekapKontrak_Click()
    fIntervalTanggal.LoadMe "LaporanKontrak"
End Sub

Private Sub fRekapOP_Click()
    FormLihat.LoadMe
End Sub

Private Sub fRetur_Click()
    fIntervalTanggal.LoadMe "LaporanRetur"
End Sub

Private Sub fReturPembelian_Click()
    FormNRPembelian.Show
End Sub

Private Sub fReturPenjualan_Click()
    FormNR.Show
End Sub

Private Sub fSC_Click()
    FormSC.Show
End Sub

Private Sub fSetHarga_Click()
    FormHargaBeli.Show
End Sub

Private Sub fSetHargaBeli_Click()
    FormHargaBeli.Show
End Sub

Private Sub fSettingAccounting_Click()
    FormSettingAccounting.Show
End Sub

Private Sub fSettingRekap_Click()
    FormLihatSetting.Show
End Sub

Private Sub fSPP_Click()
    FormSPP.Show
End Sub

Private Sub fStockBeli_Click()
    FormMasterStockBeli.Show
End Sub

Private Sub fStockOpname_Click()
    FormStockOpname.Show
End Sub

Private Sub fStockProses_Click()
    FormMasterStockProses.Show
End Sub

Private Sub fStockRetur_Click()
    FormRubahStockRetur.Show
End Sub

Private Sub fStockWaste_Click()
    FormMasterStockWaste.Show
End Sub

Private Sub fSummaryRetur_Click()
    fIntervalTanggal.LoadMe "SummaryReturPenjualan"
End Sub

Private Sub fSuratJalan_Click()
    FormSJ.Show
    FormSJ.Caption = pTipe & "-SJ"
End Sub

Private Sub fTukarStock_Click()
    FormProsesStock.Show
End Sub

Private Sub fWaste_Click()
    FormWaste.Show
End Sub



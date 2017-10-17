VERSION 5.00
Begin VB.Form FormMenu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MENU"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "REPORT"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   10395
      Begin VB.CommandButton Command1 
         Caption         =   "ABSENSI"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Tag             =   "58"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdRekapGaji 
         Caption         =   "GAJI"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Tag             =   "55"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TRANSAKSI"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   10395
      Begin VB.CommandButton cmdBulanan 
         Caption         =   "BULANAN"
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Tag             =   "56"
         Top             =   240
         Width           =   1515
      End
      Begin VB.CommandButton cmdTransPegawai 
         Caption         =   "TRANS PEGAWAI"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Tag             =   "54"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdProsesGaji 
         Caption         =   "PROSES GAJI"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Tag             =   "55"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SETUP"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6375
      Begin VB.CommandButton fPassword 
         Caption         =   "PASSWORD"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Tag             =   "36"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton fSettingRekap 
         Caption         =   "SETTING REKAP"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Tag             =   "38"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton fMarginPrinter 
         Caption         =   "MARGIN PRINTER"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Tag             =   "40"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "INPUT MASTER"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10395
      Begin VB.CommandButton cmdDefault 
         Caption         =   "DEFAULT"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Tag             =   "57"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton fMasterPegawai 
         Caption         =   "PEGAWAI"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Tag             =   "53"
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fKeterangan_Click()
    FormMasterKeterangan.Show
End Sub

Private Sub cmdBulanan_Click()
    FormBulanan.Show
End Sub

Private Sub cmdDefault_Click()
    FormDefault.Show
End Sub

Private Sub cmdProsesGaji_Click()
    FormTransGaji.Show
End Sub

Private Sub cmdRekapGaji_Click()
    rptGaji.Show
End Sub

Private Sub cmdTransPegawai_Click()
    FormTransKaryawan.Show
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

Private Sub Command1_Click()
    rptAbsensi.Show
End Sub

Private Sub fMasterPegawai_Click()
    FormMasterKaryawan.Show
End Sub

Private Sub Form_Activate()
    Caption = "MENU---" & pTipe
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
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
On Error Resume Next
    a = "select top 1 Tanggal from AutoUpdate where Nama='MENU'"
    query a
    Dim Tgl As Long
    Dim TglOld As Long
    Tgl = cD(pServerDate)
    TglOld = rs.Fields(0).Value
    If Tgl <> TglOld Then
        For i = 0 To Controls.Count - 1
            If TypeName(Controls(i)) = "CommandButton" Then
                a = "insert into m_DaftarMenu(Header, Command, Tag) values('" & _
                    esc(Controls(i).Container) & "','" & esc(Controls(i).Caption) & "'," & Controls(i).Tag & ")"
                ExecMe a
            End If
        Next
        a = "update AutoUpdate set Tanggal=" & Tgl & " where Nama='MENU'"
        ExecMe a
    End If
End Sub

Sub CekValidMenuControl()
    For i = 0 To Controls.Count - 1
        If TypeName(Controls(i)) = "CommandButton" Then
            Controls(i).Enabled = cekValid("MASUK", Controls(i).Tag, True)
        End If
    Next
    If pOffLineMode Then fGoOffline.Enabled = False
End Sub

Private Sub Form_Load()
    'ProsesAutoUpdate
    'CekValidMenuControl
    pSettingName = GetSetting("GMT", "MarginPrinter", "Name")
    a = "select top 1 Kiri, Atas from m_MarginPrinter where Nama='" & esc(pSettingName) & "'"
    query a
    If rs.RecordCount > 0 Then
        pLeftMargin = rs.Fields("Kiri").Value
        pTopMargin = rs.Fields("Atas").Value
    End If
    a = "select NamaForm, Penomoran from m_Penomoran~ order by NamaForm"
    query a
    For i = 0 To rs.RecordCount - 1
        If rs.Fields("NamaForm").Value = "SC" Then
            pNomorSC = rs.Fields("Penomoran").Value
        ElseIf rs.Fields("NamaForm").Value = "SPP" Then
            pNomorSPP = rs.Fields("Penomoran").Value
        ElseIf rs.Fields("NamaForm").Value = "SJ" Then
            pNomorSJ = rs.Fields("Penomoran").Value
        ElseIf rs.Fields("NamaForm").Value = "KW" Then
            pNomorKW = rs.Fields("Penomoran").Value
        ElseIf rs.Fields("NamaForm").Value = "NI" Then
            pNomorNI = rs.Fields("Penomoran").Value
        ElseIf rs.Fields("NamaForm").Value = "NR" Then
            pNomorNR = rs.Fields("Penomoran").Value
        End If
        rs.MoveNext
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
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> "FormMenu" And Forms(i).Name <> "FormLogin" Then
            b = MsgBox("Ada Form Yang Belum di Tutup!!!, Tutup Semua ?", vbYesNo)
            If b = vbNo Then Exit Sub
            Unload Forms(i)
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
End Sub

Private Sub fTukarStock_Click()
    FormProsesStock.Show
End Sub

Private Sub fWaste_Click()
    FormWaste.Show
End Sub



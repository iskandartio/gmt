Attribute VB_Name = "Start"
'ADODB
Global CN As New ADODB.Connection
Global RS As New ADODB.Recordset
'Crystal Report
'Global Crystal As New CRAXDRT.Application
'Global Rpt As CRAXDRT.Report
'Database Properties
Global pServerName As String
Global pConnectionString As String
Global pProvider As String
Global pDataSource As String
Global pDBPassword As String
Global pDBName As String

'Locale Info
Global pLocal As Boolean
'File location
Global pReportDir As String
Global pServerDir As String
Global pNetWorkPath As String
Global pNetWorkUser As String
Global pNetWorkPassword As String

'Server Time and Name
Global pServerDate As String
Global pDatabaseServer As String
'API
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'User Authentication
Global pBag1 As Long
Global pBagSee1 As Long
Global pBagAdd1 As Long
Global pBagEdit1 As Long
Global pBagDelete1 As Long
Global pBagPrint1 As Long
Global pBag2 As Long
Global pBagSee2 As Long
Global pBagAdd2 As Long
Global pBagEdit2 As Long
Global pBagDelete2 As Long
Global pBagPrint2 As Long
Global pBag3 As Long
Global pBag4 As Long
Global pUsr As String
Global pBosDepartemen As String
Global pUpdateHargaSC As Byte

'Others
Global temp(10) As Variant
Global pActiveForm As String
Global pTipe As String
Global pOffLineMode As Boolean
Global pAddNo As String
Global pAddNoLong As Long
Global pTransState As Byte
Global pSQLLog As String
Global pScript As New ScriptControl
Global pPing As New cPing

'Penomoran
Global pNomorSC As String
Global pNomorSPP As String
Global pNomorSJ As String
Global pNomorKW As String
Global pNomorNI As String
Global pNomorNR As String

'Report Margin
Global pLeftMargin As Single
Global pTopMargin As Single
Global pSettingName As String

Sub Main()
'On Error GoTo err
    'If App.PrevInstance Then ExitProcess 0
    pScript.Language = "VBScript"
    
    MousePointer = vbHourglass
Dim buffer As String * 100
Dim dl As Long
Dim p1 As Boolean
Dim p2 As Boolean
    dl = GetLocaleInfo(0, &HE, buffer, 100) 'Desimal
    If Left(buffer, dl - 1) = "," Then
        p1 = True
    End If
    dl = GetLocaleInfo(0, &HF, buffer, 100) 'Digit Gruping
    If Left(buffer, dl - 1) = "." Then
        p2 = True
    End If
    If p1 And p2 Then
        pLocal = True
    ElseIf Not p1 And Not p2 Then
        pLocal = False
    Else
        MsgBox "Tidak Valid"
        Exit Sub
    End If
    
    ConnectDatabase
    
    MousePointer = vbDefault
    Exit Sub
err:
End Sub

Sub ConnectDatabase()
On Error GoTo err

    allstring = Decrypt(App.Path & "\" & "setting.txt", "Iskandar Tio")
    
    'Encrypt "c:\gmt\gmt\Program\local.txt", "c:\gmt\gmt\Program\setting.txt", "Iskandar Tio", True
    Dim a() As String
    a = Split(allstring, Chr(13) & Chr(10))
    For i = 0 To UBound(a)
        If InStr(a(i), "Provider") = 1 Then
            pProvider = Mid(a(i), Len("Provider") + 2)
        ElseIf InStr(a(i), "DatabasePassword") = 1 Then
            pDBPassword = Mid(a(i), Len("DatabasePassword") + 2)
        ElseIf InStr(a(i), "DatabaseName") = 1 Then
            pDBName = Mid(a(i), Len("DatabaseName") + 2)
        ElseIf InStr(a(i), "ServerDir") = 1 Then
            pServerDir = Mid(a(i), Len("ServerDir") + 2)
        ElseIf InStr(a(i), "ReportDir") = 1 Then
            pReportDir = Mid(a(i), Len("ReportDir") + 2)
        ElseIf InStr(a(i), "NetWorkPath") = 1 Then
            pNetWorkPath = Mid(a(i), Len("NetWorkPath") + 2)
        ElseIf InStr(a(i), "NetWorkUser") = 1 Then
            pNetWorkUser = Mid(a(i), Len("NetWorkUser") + 2)
        ElseIf InStr(a(i), "NetWorkPassword") = 1 Then
            pNetWorkPassword = Mid(a(i), Len("NetWorkPassword") + 2)
        End If
    Next
    'If pServerDir <> "app.path\" Then
    '    MappingNetWork pNetWorkPath, pNetWorkUser, pNetWorkPassword
    'End If
        pConnectionString = "Provider=" & pProvider & ";Data Source=" & pServerDir & "\" & pDBName & ";Persist Security Info=False;Jet OLEDB:Database Password=" & pDBPassword
    pConnectionString = Replace(pConnectionString, "app.path", App.Path)
        Connect pConnectionString
    
        If CN.State = 1 Then
        If pServerName = App.Path Then
            MsgBox "SEKARANG ANDA BERADA DI MODE OFFLINE"
            pOffLineMode = True
        Else
                d = GetFile2("SQLLog.txt")
            If d <> "" Then
                c = MsgBox("Execute Offline Data?", vbYesNo)
                If c = vbYes Then
                    b = Split(d, vbCrLf)
                    CN.BeginTrans
                    v = False
                    For i = 0 To UBound(b)
                        If Len(b(i)) > 1 Then
                            CN.Execute b(i), c
                            If c < 1 Then
                                CN.RollbackTrans
                                v = True
                                Exit For
                            End If
                        End If
                    Next
                    If v Then
                        CN.CommitTrans
                    Else
                        CopyFileAPI "SQLLog.txt", "SQLLog" & Format(date + Time, "ddmmyy-hhnn") & ".txt"
                    End If
                    Kill "SQLLog.txt"
                End If
            End If
        End If
    Else
        End
    End If
    If InStr(pProvider, "SQLOLEDB") <> 0 Then
        pDatabaseServer = "SQL Server"
    ElseIf InStr(pProvider, "Jet.OLEDB") <> 0 Then
        pDatabaseServer = "Access"
    End If
    If pDatabaseServer = "Access" Then
        a1 = "select now()"
    ElseIf pDatabaseServer = "SQL Server" Then
        a1 = "select getdate()"
    End If
    query a1
    If Format("1/12/2006", "dd") = "12" Then
        pServerDate = Format(RS.Fields(0).Value, "mm/dd/yy")
    Else
        pServerDate = Format(RS.Fields(0).Value, "dd/mm/yy")
    End If
    's = "create table mutasi(tgl Long, IdStock long, inBox long, inKg double, outBox long, outKg double, akhirBox long, akhirKg double)"
    'ExecMe s
    's = "create table mutasi_hist(tgl Long, IdStock long, inBox long, inKg double, outBox long, outKg double, akhirBox long, akhirKg double)"
    'ExecMe s
    's = "alter table t_SPPDetailDTY add dtl text"
    'ExecMe s
    's = "alter table t_SPPDetailPE add dtl text"
    'ExecMe s
    'MsgBox IsNull(RS.Fields(0).Value)
    FormLogin.Show
    Exit Sub
err:
    MsgBox err.Description
End Sub

Sub HelpMe(ByVal tHelp As String, tf As Form, Optional ByVal tLeft As Single = 0, Optional ByVal tTop As Single = 0)
    If tHelp = "Nama Customer" Then
        FormHelp.LoadMe "select distinct Nama from m_customer where IsActive=1 order by Nama", tHelp, 3000, "", tf, tLeft, tTop
    ElseIf tHelp = "Nama Customer With Kode" Then
        FormHelp.LoadMe "select distinct Nama, Kode&'@'&Nama from m_customer where IsActive=1 order by Nama", "Nama@*Kode", "3000@700", "", tf, tLeft, tTop, 1
    ElseIf tHelp = "Keterangan" Then
        FormHelp.LoadMe "select IdKet, Keterangan from m_Keterangan order by IdKet", "IdKet@Ket", "700@2500", "", tf, tLeft, tTop, 1
    ElseIf tHelp = "Nama Supplier" Then
        FormHelp.LoadMe "select distinct Nama from m_Supplier order by Nama", tHelp, 3000, "", tf, tLeft, tTop
    ElseIf tHelp = "Nama Bank" Then
        FormHelp.LoadMe "select distinct NamaBank from m_Bank order by NamaBank", tHelp, 3000, "", tf, tLeft, tTop
    ElseIf tHelp = "Status Giro" Then
        FormHelp.LoadMe "select 0,'BELUM CAIR' from t_Updated union all select 1, 'SUDAH CAIR' from t_Updated", "@" & tHelp, "500@3000", "", tf, tLeft, tTop
    ElseIf tHelp = "Cara Bayar" Then
        FormHelp.LoadMe "select distinct CaraBayar from t_STTPembayaran~ order by CaraBayar", tHelp, 3000, "", tf, tLeft, tTop
    ElseIf tHelp = "Keterangan Pembayaran" Then
        FormHelp.LoadMe "select distinct Keterangan from t_STT union all select distinct Keterangan from t_PembelianSTT order by Keterangan", tHelp, 3000, "", tf, tLeft, tTop
    ElseIf tHelp = "Masuk" Then
        FormHelp.LoadMe "select 0, 'YANG KELUAR SAJA' from t_Updated union all select 1, 'YANG MASUK SAJA' from t_Updated", "@" & tHelp, "500@3000", "", tf, tLeft, tTop
    ElseIf tHelp = "Kendaraan" Then
        FormHelp.LoadMe "select distinct ModelKendaraan, MerkKendaraan from m_stock order by ModelKendaraan, MerkKendaraan", "Model Kendaraan@Merk Kendaraan", "3000@3000", "", tf, tLeft, tTop
    ElseIf tHelp = "Mata Uang" Then
        FormHelp.LoadMe "select Kode, Nama, Negara from m_MataUang order by Kode", "Nama@Negara", "3000@3000", "", tf, tLeft, tTop
    ElseIf tHelp = "Status" Then
        FormHelp.LoadMe "select 0,'BELUM POTONG' from t_Updated union all select 1, 'SUDAH POTONG' from t_Updated", "@" & tHelp, "500@3000", "", tf, tLeft, tTop
    ElseIf tHelp = "No Contract Closed" Then
        FormHelp.LoadMe "select distinct NoContract from t_PR where Closed=1 and NoContract<>''", "No Contract", "2000", "", tf, tLeft, tTop
    ElseIf tHelp = "No Contract Not Closed" Then
        FormHelp.LoadMe "select distinct NoContract from t_PR where Closed=0 and NoContract<>''", "No Contract", "2000", "", tf, tLeft, tTop
    ElseIf tHelp = "No Contract All" Then
        FormHelp.LoadMe "select distinct NoContract from t_PR where NoContract<>''", "NoContract", "2000", "", tf, tLeft, tTop
    ElseIf tHelp = "Closed" Then
        FormHelp.LoadMe "select 0,'Not Closed' from t_Updated union all select 1,'Closed' from t_Updated", "@CLOSED", "900@2000", "", tf, tLeft, tTop
    ElseIf tHelp = "No Account" Then
        FormHelp.LoadMe "select NoAccount, Deskripsi from m_ChartAccount where child=0 order by NoAccount", "No Account@Deskripsi", "1500@2000", "", tf, tLeft, tTop
    ElseIf tHelp = "Jenis Barang" Then
        FormHelp.LoadMe "select distinct Jenis, KodeJenis from m_StockBeli order by Jenis", "Jenis@KodeJenis", "2000@700", "", tf, tLeft, tTop
    ElseIf tHelp = "Kode Barang" Then
        FormHelp.LoadMe "select distinct KodeBarang from m_Stock~ where IsActive=1 order by KodeBarang", "KodeBarang", "2000", "", tf, tLeft, tTop
    End If
End Sub

Sub InputTanggal()
    a = "select max(Tanggal) as  v from m_kurs"
    query a
    tY = Left(RS.Fields(0).Value, 2) + 1
    y = tY
    Tgl = y * 10000 + 101
    maxTanggal = 31
    Do
        Nilai = 0
        a = "insert into m_kurs(Tanggal, Nilai) values(" & Tgl & "," & cNum(Nilai) & ")"
        CN.Execute a
        Tgl = Tgl + 1
        If CInt(Right(Tgl, 2)) > maxTanggal Then
            Tgl = (Tgl \ 100 + 1) * 100 + 1
            b = CInt(Mid(Tgl, 2, 2))
            If b = 4 Or b = 6 Or b = 9 Or b = 11 Then
                maxTanggal = 30
            ElseIf b = 2 Then
                maxTanggal = 28
                If y Mod 4 = 0 Then maxTanggal = 29
            Else
                maxTanggal = 31
            End If
            If Tgl > y * 10000 + 1300 Then Exit Do
        End If
    Loop
End Sub



Function IfNull(ByVal tSource) As String
    IfNull = "iif(isnull(" & tSource & "),0," & tSource & ")"
End Function

Sub HitungHPP(ByVal awal As Long)
Dim rsBeli As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim QTYJualSisa As Double
Dim QTYBeliSisa As Double
Dim HPP As Double
    rsTemp.CursorLocation = adUseServer
    a = "select 0, TanggalBTBDetail,t_BTBDetail.QTY-" & IfNull("t_NRPembelianDetail.QTY") & " as QTYBeli,(1+t_BTBDetail.PPNYesNo/10)*(t_BTBDetail.Harga-" & IfNull("t_NRPembelianDetail.Potongan") & "/(t_BTBDetail.QTY-" & IfNull("t_NRPembelianDetail.QTY") & ")) as HargaBeli, [t_BTBDetail].[IdStock] FROM t_BTBDetail LEFT JOIN t_NRPembelianDetail ON (t_BTBDetail.IdStock=t_NRPembelianDetail.IdStock) AND (t_BTBDetail.NoBTB=t_NRPembelianDetail.NoBTB) order by t_BTBDetail.IdStock, TanggalBTBDetail Desc"
    rsBeli.CursorLocation = adUseClient
    rsBeli.Open a, CN, adOpenForwardOnly, adLockReadOnly
    a = "select IdtOut, QTY, IdStock, Tanggal from t_Out where FromNPB>0 and Tanggal>=" & awal & " union all select -1, Jumlah, IdStock, 1000000 from m_StockBeli where StockWaste= 0 and TipeStock=1 and Jumlah>0 order by IdStock, Tanggal Desc"
    query a
    a = "select IdtOut, HPP from temp_HPP"
    rsTemp.Open a, CN, adOpenKeyset, adLockOptimistic
    QTYBeliSisa = rsBeli.Fields("QTYBeli").Value
    For i = 0 To RS.RecordCount - 1
        QTYJualSisa = RS.Fields("QTY").Value
        HPP = 0
        v = 0
        'Cek Kalau Tidak Pernah Dibeli, IdStock Beli sudah lebih besar IdStock Jualnya
        If rsBeli.EOF Then v = 1 Else QTYBeliSisa = rsBeli.Fields("QTYBeli").Value
        If v = 0 Then
            If rsBeli.Fields("IdStock").Value > RS.Fields("IdStock").Value Then v = 1
        End If
        If v = 1 Then
            HPP = -1
            GoTo langsung
        End If
        'Id Stock Beli lebih kecil dari jual cari sampai IdStock-nya sama
        While RS.Fields("IdStock").Value > rsBeli.Fields("IdStock").Value
            rsBeli.MoveNext
            QTYBeliSisa = rsBeli.Fields("QTYBeli").Value
        Wend
        'IdStock Beli sudah lebih besar maka tidak pernah dibeli
        If rsBeli.Fields("IdStock").Value > RS.Fields("IdStock").Value Then
            HPP = -1
            GoTo langsung
        End If
        'Bila IdStock sudah sama
        While QTYJualSisa >= QTYBeliSisa
            HPP = HPP + QTYBeliSisa * rsBeli.Fields("HargaBeli").Value
            QTYJualSisa = QTYJualSisa - QTYBeliSisa
            rsBeli.MoveNext
            'Kalau Sisanya tidak nol
            If QTYJualSisa > 0 Then
                If rsBeli.EOF Then v = 1
                If v = 0 Then
                    If rsBeli.Fields("IdStock").Value > RS.Fields("IdStock").Value Then v = 1
                End If
                If v = 1 Then
                    HPP = -1
                    GoTo langsung
                End If
                QTYBeliSisa = rsBeli("QTYBeli").Value
            End If
        Wend
        If QTYBeliSisa > QTYJualSisa And QTYJualSisa > 0 Then
            HPP = HPP + QTYJualSisa * rsBeli.Fields("HargaBeli").Value
            QTYBeliSisa = QTYBeliSisa - QTYJualSisa
        End If
langsung:
        If RS.Fields("IdtOut").Value > -1 Then
            rsTemp.AddNew
            rsTemp!IdtOut = RS.Fields("IdtOut").Value
            rsTemp!HPP = HPP
        End If
        RS.MoveNext
    Next
    'rsTemp.Update
    CN.BeginTrans
    a = "update temp_HPP left join t_Out on temp_HPP.IdtOut=t_Out.IdtOut set t_Out.HPP=temp_HPP.HPP, t_Out.StatusHPP=0"
    ExecMe a
    a = "delete from temp_HPP"
    ExecMe a
    CN.CommitTrans
End Sub

Sub HitungHPPProses(ByVal awal As Long)
Dim x As New XArrayDB
Dim rsTemp As New ADODB.Recordset
Dim HPPClue As String
Dim sHPP As String
Dim IdtOut As Long
Dim IdHPP As Long
    a = "select IdtOut, HPP from t_Out where FromNPB>0 and StatusHPP=0 order by IdtOut"
    query a
    If RS.RecordCount = 0 Then Exit Sub
    x.LoadRows RS.GetRows
    rsTemp.CursorLocation = adUseServer
    a = "select IdtOut, HPPClue, HPP from t_Out where Tanggal>=" & awal & " and len(HPPClue)>0 order by IdtOut"
    rsTemp.Open a, CN, adOpenKeyset, adLockOptimistic
    For i = 0 To rsTemp.RecordCount - 1
        HPPClue = rsTemp.Fields("HPPClue").Value
        IdtOut = rsTemp.Fields("IdtOut").Value
        j = 0
        sHPP = ""
        Do
            b = InStr(j + 1, HPPClue, "@") + 1
            If b = 1 Then Exit Do
            sHPP = sHPP & Mid(HPPClue, j + 1, b - j - 2)
            j = InStr(b, HPPClue, "@")
            IdHPP = CLng(Mid(HPPClue, b, j - b))
            l = 0
            h = x.UpperBound(1)
            Do While l <= h
                m = (l + h) \ 2
                If x(m, 0) = IdHPP Then
                    sHPP = sHPP & x(m, 1)
                    Exit Do
                ElseIf IdHPP < x(m, 0) Then
                    h = m - 1
                ElseIf IdHPP > x(m, 0) Then
                    l = m + 1
                End If
            Loop
        Loop
        sHPP = sHPP & Mid(HPPClue, j + 1)
        rsTemp.Fields("HPP").Value = pScript.Eval(sHPP)
        rsTemp.MoveNext
    Next
    a = "update t_Out set StatusHPP=1 where Terpakai>=Jumlah"
    CN.Execute a
End Sub

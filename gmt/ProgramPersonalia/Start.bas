Attribute VB_Name = "Start"
'ADODB
Global CN As New ADODB.Connection
Global rs As New ADODB.Recordset

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
Global pNetworkPath As String
Global pNetworkUser As String
Global pNetworkPassword As String

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
'Global pScript As New ScriptControl
'Global pPing As New cPing

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

    If App.PrevInstance Then ExitProcess 0
    
'    pScript.Language = "VBScript"
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
'On Error GoTo err
    allstring = Decrypt(App.Path & "\" & "setting.txt", "Iskandar Tio")
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
            pServerDir = Replace(Mid(a(i), Len("ServerDir") + 2), "$", pServerName)
        ElseIf InStr(a(i), "ReportDir") = 1 Then
            pReportDir = Replace(Mid(a(i), Len("ReportDir") + 2), "$", pServerName)
        ElseIf InStr(a(i), "NetWorkPath") = 1 Then
            pNetworkPath = Replace(Mid(a(i), Len("NetWorkPath") + 2), "$", pServerName)
        ElseIf InStr(a(i), "NetWorkUser") = 1 Then
            pNetworkUser = Replace(Mid(a(i), Len("NetWorkUser") + 2), "$", pServerName)
        ElseIf InStr(a(i), "NetWorkPassword") = 1 Then
            pNetworkPassword = Replace(Mid(a(i), Len("NetWorkPassword") + 2), "$", pServerName)
        End If
    Next
    If pServerDir <> "app.path\" Then
        MappingNetWork pNetworkPath, pNetworkUser, pNetworkPassword
    End If
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
                        CopyFileAPI "SQLLog.txt", "SQLLog" & Format(Date + Time, "ddmmyy-hhnn") & ".txt"
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
        a1 = "select date()"
    ElseIf pDatabaseServer = "SQL Server" Then
        a1 = "select getdate()"
    End If
    query a1
    pServerDate = rs.Fields(0).Value
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
    ElseIf tHelp = "Karyawan" Then
        FormHelp.LoadMe "Select NIK, Nama, Departemen from m_Karyawan order by NIK", "NIK@Nama@Departemen", "1000@2000@1500", "ZFill6@String@String", tf, tLeft, tTop
    End If
End Sub

Sub InputTanggal(ByVal tY As Byte)
    y = tY
    Tgl = y * 10000 + 101
    maxTanggal = 31
    Do
        Nilai = 9200 + Rnd(1) * 250 - 125
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


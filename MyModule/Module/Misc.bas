Attribute VB_Name = "Misc"
Function GetID(ByVal tLong As Variant) As String
On Error GoTo err
Dim s As String
Dim s2() As Byte
Dim s3() As Byte
    s = tLong
    s2 = StrConv(s, vbFromUnicode)
    ReDim s3(UBound(s2))
    j = 0
    For i = 0 To 5
        s3(j) = (s2(i) * 205.57) Mod 26 + 65
        j = j + 1
    Next
    ReDim Preserve s3(j - 1)
    GetID = StrConv(s3, vbUnicode)
err:
End Function
Sub c1(ByRef a As String, ByRef p As Long, ByVal s As String)
    l = Len(s)
    Mid(a, p, l) = s
    p = p + l
End Sub

Private Sub Swap(z() As Single, ByVal l As Long, ByVal r As Long)
    Dim t As Single
    t = z(l)
    z(l) = z(r)
    z(r) = t
End Sub

Private Sub Paint(obj As Object, tLabel As Object)
    obj.FontName = tLabel.FontName
    obj.FontBold = tLabel.FontBold
    obj.FontItalic = tLabel.FontItalic
    obj.ForeColor = tLabel.ForeColor
    obj.FontStrikethru = tLabel.FontStrikethru
    obj.FontUnderline = tLabel.FontUnderline
    obj.FontSize = tLabel.FontSize
End Sub

Sub PaintDetail(ByVal tStr As String, obj As Object, tLabel As Label, ByVal t As Single, tMax As Byte, Optional ByVal tCariMax As Boolean = False, Optional ByVal tAdder As Single = 0)
    If tCariMax Then
        b = Split(WT(tStr, obj, tLabel), vbCrLf)
        If tMax < UBound(b) + 1 Then tMax = UBound(b) + 1
        Exit Sub
    End If
    Paint obj, tLabel
    b = Split(WT(tStr, obj, tLabel), vbCrLf)
    If tMax < UBound(b) + 1 Then tMax = UBound(b) + 1
    For i = 0 To UBound(b)
        obj.CurrentY = i * tLabel.Height + t + pTopMargin
        obj.CurrentX = GetCurrentX(b(i), obj, tLabel) + pLeftMargin + tAdder
        obj.Print b(i)
    Next
End Sub

Sub PaintBox(obj As Object, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    PaintLine obj, x1, y1, x2, y1
    PaintLine obj, x1, y2, x2 + 10, y2
    PaintLine obj, x1, y1, x1, y2
    PaintLine obj, x2, y2, x2, y1
End Sub

Sub PaintHeader(ByVal tStr As String, obj As Object, tLabel As Label, Optional ByVal tPlus As Single = 0)
    Paint obj, tLabel
    b = Split(WT(tStr, obj, tLabel), vbCrLf)
    obj.CurrentY = tLabel.Top + tPlus + pTopMargin
    For i = 0 To UBound(b)
        obj.CurrentX = GetCurrentX(b(i), obj, tLabel) + pLeftMargin
        obj.Print b(i)
    Next
End Sub

Sub PaintLine(obj As Object, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    obj.Line (x1 + pLeftMargin, y1 + pTopMargin)-(x2 + pLeftMargin, y2 + pTopMargin)
End Sub

Sub PaintDot(obj As Object, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If x1 = x2 Then
        For i = y1 To y2 Step 75
            obj.PSet (x1 + pLeftMargin, i + pTopMargin)
        Next
    ElseIf y1 = y2 Then
        For i = x1 To x2 Step 75
            obj.PSet (i + pLeftMargin, y1 + pTopMargin)
        Next
    Else
        m = (y2 - y1) / (x2 - x1)
        For i = x1 To x2 Step 75
            obj.PSet (i + pLeftMargin, m * (i - x1) + y1 + pTopMargin)
        Next
    End If
End Sub

Private Function GetCurrentX(ByVal tStr As String, obj As Object, tLabel As Label) As Single
    GetCurrentX = tLabel.Left
    If tLabel.Alignment = 1 Then
        GetCurrentX = tLabel.Left + tLabel.Width - obj.TextWidth(tStr)
    ElseIf tLabel.Alignment = 2 Then
        GetCurrentX = tLabel.Left + (tLabel.Width - obj.TextWidth(tStr)) / 2
    End If
End Function

Function WT(ByVal tStr As String, obj As Object, tLabel As Object) As String
    WT = ""
    tStr = RTrim(tStr)
    l = tLabel.Width
    If obj Is Printer Then l = l + 30
    numline = obj.TextWidth(tStr) / l
    If numline <= 1 Then
        WT = tStr
        Exit Function
    End If
    Dim b() As Byte
    b = StrConv(tStr, vbFromUnicode)
    c = Round(Len(tStr) / numline + 2)
    d = c
    e = 0
    Do
        While d > UBound(b)
            d = d - 1
        Wend
        d1 = d
        h = False
        g = Mid(tStr, e + 1, d + 1)
        Do While obj.TextWidth(g) > tLabel.Width  'kepanjangan
            While e + d > UBound(b)
                d = d - 1
            Wend
            While b(e + d) <> 32
                If d = 0 Then
                    d = d1
                    While obj.TextWidth(g) > tLabel.Width
                        d = d - 1
                        g = Mid(tStr, e + 1, d + 1)
                    Wend
                    h = True
                    Exit Do
                End If
                d = d - 1
            Wend
            g = Mid(tStr, e + 1, d)
            If b(e + d) = 32 Then d = d - 1
            h = True
        Loop
        If Not h Then
            While obj.TextWidth(Trim(g)) < tLabel.Width  'kependekan
                j = d
                While b(e + d) <> 32
                    If d = UBound(b) Then Exit Do
                    d = d + 1
                Wend
                While b(e + d) = 32
                    d = d + 1
                Wend
                g = Mid(tStr, e + 1, d)
            Wend
            d = j
            g = Mid(tStr, e + 1, d)
        End If
        WT = WT & vbCrLf & RTrim(g)
        e = e + Len(g)
        If e > UBound(b) Then e = UBound(b)
        While b(e) = 32
            e = e + 1
        Wend
        d = c
        g = Mid(tStr, e + 1)
        If obj.TextWidth(g) < tLabel.Width Then Exit Do
    Loop
    WT = Mid(WT, 3) & vbCrLf & g
End Function

Function esc(ByVal tStr As Variant) As String
On Error GoTo err
    esc = ""
    esc = Trim(Replace(tStr, "'", "''"))
err:
End Function

Sub BubbleSort(z() As Single, ByVal l As Long, ByVal r As Long)
    Dim i, j As Long
    Dim v As Variant
    For i = l To r - 1
        For j = i + 1 To r
            If z(i) > z(j) Then Swap z, i, j
        Next
    Next
End Sub

Function RZerofill(ByVal tNo As String, ByVal tLen As Byte) As String
    RZerofill = String(tLen, "0")
    Mid(RZerofill, 1) = tNo
End Function

Function cekValid(ByVal tCommand As String, ByVal tTag As Integer, Optional ByVal tNoMessage As Boolean = False) As Boolean
    If tTag < 31 Then
        tPass = 2 ^ tTag
        If tCommand = "MASUK" Then
            v = pBag1 And tPass
        ElseIf tCommand = "SEE" Then
            v = pBagSee1 And tPass
        ElseIf tCommand = "NEW" Then
            v = pBagAdd1 And tPass
        ElseIf tCommand = "EDIT" Then
            v = pBagEdit1 And tPass
        ElseIf tCommand = "DELETE" Then
            v = pBagDelete1 And tPass
        ElseIf tCommand = "PRINT" Then
            v = pBagPrint1 And tPass
        End If
    Else
        tPass = 2 ^ (tTag - 31)
        If tCommand = "MASUK" Then
            v = pBag2 And tPass
        ElseIf tCommand = "SEE" Then
            v = pBagSee2 And tPass
        ElseIf tCommand = "NEW" Then
            v = pBagAdd2 And tPass
        ElseIf tCommand = "EDIT" Then
            v = pBagEdit2 And tPass
        ElseIf tCommand = "DELETE" Then
            v = pBagDelete2 And tPass
        ElseIf tCommand = "PRINT" Then
            v = pBagPrint2 And tPass
        End If
    End If
    cekValid = IIf(v = 0, False, True)
    If Not cekValid And Not tNoMessage Then MsgBox "TIDAK BERHAK !!!"
End Function

Function Add_Tanggal2(ByVal tTanggal As Long, ByVal tDiff As Integer) As Long
On Error GoTo err
    t = zerofill(tTanggal, 6)
    d = CInt(Right(t, 2))
    m = CInt(Mid(t, 3, 2))
    y = CInt(Left(t, 2))
    If Not cek_tanggal(cTanggal(tTanggal)) Then Exit Function
    d = d + tDiff
    Maximum = cari_max(m, y)
    While d > Maximum
        m = m + 1
        d = d - Maximum
        Maximum = cari_max(m, y)
    Wend
    While d <= 0
        m = m - 1
        Maximum = cari_max(m, y)
        d = d + Maximum
    Wend
    While m > 12
        m = m - 12
        y = y + 1
    Wend
    While m <= 0
        m = m + 12
        y = y - 1
    Wend
    Add_Tanggal2 = CLng(zerofill(y, 2) & zerofill(m, 2) & zerofill(d, 2))
err:
End Function


Sub CampuranSort(z() As Single, ByVal l As Long, ByVal r As Long)
    Dim i, j As Long
    Dim v As Double
    If l >= r Then Exit Sub
    If r < l + 15 Then
        BubbleSort z, l, r
        Exit Sub
    End If
    i = (r + l) \ 2
    If z(l) > z(i) Then Swap z, l, i
    If z(l) > z(r) Then Swap z, l, r
    If z(i) > z(r) Then Swap z, i, r
    Swap z, i, r - 1
    j = r - 2
    i = l + 1
    v = z(r - 1)
    If j < i Then Exit Sub
    Do
        While z(i) < v
            i = i + 1
        Wend
        While z(j) > v
            j = j - 1
        Wend
        If j < i Then Exit Do
        Swap z, i, j
        i = i + 1
        j = j - 1
    Loop
    Swap z, i, r - 1
    CampuranSort z, l, j
    CampuranSort z, i + 1, r
End Sub

Sub QuickSort(z() As Single, ByVal l As Long, ByVal r As Long)
    Dim i, j As Long
    Dim v As Double
    If l >= r Then Exit Sub
    i = (r + l) \ 2
    If z(l) > z(i) Then Swap z, l, i
    If z(l) > z(r) Then Swap z, l, r
    If z(i) > z(r) Then Swap z, i, r
    Swap z, i, r - 1
    j = r - 2
    i = l + 1
    v = z(r - 1)
    If j < i Then Exit Sub
    Do
        While z(i) < v
            i = i + 1
        Wend
        While z(j) > v
            j = j - 1
        Wend
        If j < i Then Exit Do
        Swap z, i, j
        i = i + 1
        j = j - 1
    Loop
    Swap z, i, r - 1
    QuickSort z, l, j
    QuickSort z, i + 1, r
End Sub

Function cek_tanggal(ByVal a As String, Optional ByVal tLong As Boolean)
    On Error GoTo err
    a = Replace(a, "/", "")
    If Format("1/12/2006", "dd") = "12" Then
        bulan = CInt(Left(a, 2))
        tanggal = CInt(Mid(a, 3, 2))
    Else
        tanggal = CInt(Left(a, 2))
        bulan = CInt(Mid(a, 3, 2))
    End If
    If tLong Then
        tahun = CInt(Right(a, 4))
    Else
        tahun = CInt(Right(a, 2))
    End If
    Maximum = cari_max(bulan, tahun)
    If tanggal > 0 And bulan > 0 And tanggal <= Maximum And bulan < 13 Then
        cek_tanggal = True
        Exit Function
    End If
err:
    cek_tanggal = False
End Function

Function cek_currency(ByVal a As String) As Boolean
On Error GoTo err
    cek_number = False
    Dim b() As Byte
    b = StrConv(a, vbFromUnicode)
    c = False
    If b(0) < 48 And b(0) > 57 Then Exit Function
    For i = 1 To UBound(b)
        If (b(i) = 44 Or b(i) = 46) Then
            If c Then GoTo err
            c = True
        End If
    Next
    cek_currency = True
err:
End Function

Function cek_integer(ByVal a As String) As Boolean
On Error GoTo err
    cek_integer = False
    b = CLng(a)
    cek_integer = True
    Exit Function
err:
End Function

Function cTanggal(ByVal a As Variant, Optional ByVal tLong As Boolean = False)
'a dari long
On Error GoTo err
    If a = 0 Then GoTo err
    If Not tLong Then
        b = zerofill(a, 6)
        If Format("1/12/2006", "dd") = "12" Then
            cTanggal = Mid(b, 3, 2) & "/" & Right(b, 2) & "/" & Left(b, 2)
        Else
            cTanggal = Right(b, 2) & "/" & Mid(b, 3, 2) & "/" & Left(b, 2)
        End If
    Else
        b = zerofill(a, 8)
        If Format("1/12/2006", "dd") = "12" Then
            cTanggal = Mid(b, 5, 2) & "/" & Right(b, 2) & "/" & Left(b, 4)
        Else
            cTanggal = Right(b, 2) & "/" & Mid(b, 5, 2) & "/" & Left(b, 4)
        End If
    End If
    Exit Function
err:
    cTanggal = "__/__/__"
End Function

Function cTanggal2(ByVal a As Long)
'a dari long, hasilnya tidak pakai /
On Error GoTo err
    If a = 0 Then GoTo err
    b = zerofill(a, 6)
    If Format("1/12/2006", "dd") = "12" Then
        cTanggal2 = Mid(b, 3, 2) & Right(b, 2) & Left(b, 2)
    Else
        cTanggal2 = Right(b, 2) & Mid(b, 3, 2) & Left(b, 2)
    End If
    Exit Function
err:
    cTanggal2 = "__/__/__"
End Function

Function cTanggal3(ByVal a As Long)
'a dari grid
On Error GoTo err
    If a = 0 Then GoTo err
    b = zerofill(a, 6)
    If Format("1/12/2006", "dd") = "12" Then
        cTanggal3 = Mid(b, 3, 2) & "/" & Left(b, 2) & "/" & Right(b, 2)
    Else
        cTanggal3 = Left(b, 2) & "/" & Mid(b, 3, 2) & "/" & Right(b, 2)
    End If
    Exit Function
err:
    cTanggal3 = "__/__/__"
End Function

Function GetDateSQL()
    If pDatabaseServer = "Access" Then
        GetDateSQL = "now()"
    ElseIf pDatabaseServer = "SQL Server" Then
        GetDateSQL = "getdate()"
    End If
End Function

Function cD(ByVal a As String, Optional ByVal tLong As Boolean = False) As Variant
On Error GoTo err
    a = Replace(a, "/", "")
    If Not cek_tanggal(a, tLong) Then GoTo err
    If Not tLong Then
        If Format("1/12/2006", "dd") = "12" Then
            b = Right(a, 2) & Left(a, 2) & Mid(a, 3, 2)
        Else
            b = Right(a, 2) & Mid(a, 3, 2) & Left(a, 2)
        End If
    Else
        If Format("1/12/2006", "dd") = "12" Then
            b = Right(a, 4) & Left(a, 2) & Mid(a, 3, 2)
        Else
            b = Right(a, 4) & Mid(a, 3, 2) & Left(a, 2)
        End If
    End If
    cD = CLng(b)
    Exit Function
err:
   cD = "A"
End Function

Function zerofill(ByVal tStr As String, ByVal num As Integer) As String
On Error Resume Next
    zerofill = String(num, "0")
    Mid(zerofill, num + 1 - Len(tStr)) = tStr
End Function

Function fFill(ByVal tStr As String, ByVal num As Integer, Optional ByVal tChar As String = " ") As String
    fFill = String(num - Len(tStr), tChar) & tStr
End Function

Function cNoCents(ByVal tValue As Double)
    cNoCents = Format(tValue, "#,##0")
End Function

Function cDecimal(ByVal tValue As Double)
    cDecimal = Format(tValue, "#,##0.00")
End Function

Sub TentukanPenomoran()
    a = "select * from m_penomoran" & pTipe
    query a
    For i = 0 To RS.RecordCount - 1
        If RS.Fields(0).Value = "SC" Then
            pNomorSC = RS.Fields("Penomoran").Value
        ElseIf RS.Fields(0).Value = "SPP" Then
            pNomorSPP = RS.Fields("Penomoran").Value
        ElseIf RS.Fields(0).Value = "SJ" Then
            pNomorSJ = RS.Fields("Penomoran").Value
        ElseIf RS.Fields(0).Value = "KW" Then
            pNomorKW = RS.Fields("Penomoran").Value
        ElseIf RS.Fields(0).Value = "NR" Then
            pNomorNR = RS.Fields("Penomoran").Value
        ElseIf RS.Fields(0).Value = "NI" Then
            pNomorNI = RS.Fields("Penomoran").Value
        End If
        RS.MoveNext
    Next
End Sub

Function cNum(ByVal num As Variant, Optional ByVal Prec As Byte = 2) As String
On Error GoTo err
    cNum = Round(CDec(num), Prec)
    If Format("0.0", ".") = "," Then
        cNum = Replace(cNum, ",", ".", 1, 1)
    End If
    Exit Function
err:
    cNum = 0
End Function

Function add_tanggal(ByVal a As String, ByVal b As Integer) As String
On Error GoTo err
    add_tanggal = ""
    If Not cek_tanggal(a) Then Exit Function
    d = CInt(Left(a, 2))
    m = CInt(Mid(a, 4, 2))
    y = CInt(Right(a, 2))
    d = d + b
    Maximum = cari_max(m, y)
    While d > Maximum
        m = m + 1
        d = d - Maximum
        Maximum = cari_max(m, y)
    Wend
    While d <= 0
        m = m - 1
        If m = 0 Then
            m = 12
            y = y - 1
        End If
        Maximum = cari_max(m, y)
        d = d + Maximum
    Wend
    While m > 12
        m = m - 12
        y = y + 1
    Wend
    While m <= 0
        m = m + 12
        y = y - 1
    Wend
    add_tanggal = zerofill(d, 2) & "/" & zerofill(m, 2) & "/" & zerofill(y, 2)
err:
End Function

Function cari_max(ByVal m As Integer, ByVal y As Integer)
    cari_max = 31
    If m = 4 Or m = 6 Or m = 9 Or m = 11 Then
        cari_max = 30
    ElseIf m = 2 Then
        If y Mod 4 = 0 And (y Mod 100 <> 0 Or y Mod 400 = 0) Then
            cari_max = 29
        Else
            cari_max = 28
        End If
    End If
End Function


Function Encrypt(ByVal FFrom As String, ByVal FileTo As String, Key As String, Optional ByVal FFile As Boolean = False) As String
'On Error Resume Next
    Dim a() As Byte
    a = StrConv(Key, vbFromUnicode)
    Dim b() As Byte
    If FFile Then b = GetFile(FFrom) Else b = StrConv(FFrom, vbFromUnicode)
    hitung = 0
    For i = 0 To UBound(a)
        hitung = hitung + a(i)
    Next
    Dim d() As Byte
    ReDim d(UBound(b) + 100)
    Randomize
    i = 0
    j = 0
    cek = 0
    While cek < hitung
        d(j) = CByte(Rnd * 256)
        cek = cek + d(j)
        j = j + 1
        i = i + 1
        If i = Len(Key) Then i = 0
    Wend
    Last = d(j - 1)
    For k = 0 To UBound(b)
        d(j) = (Last + b(k) + a(i)) Mod 256
        Last = d(j)
        j = j + 1
        i = i + 1
        If i = Len(Key) Then i = 0
    Next
    ReDim Preserve d(j)
    If FileTo = "" Then
        Encrypt = StrConv(d(), vbUnicode)
        Encrypt = Left(Encrypt, Len(Encrypt) - 1)
    Else
        PutFile FileTo, d
    End If
End Function

Function Decrypt(ByVal FileFrom As String, ByVal Key As String, Optional ByVal FileTo As String = "", Optional ByVal NotFile As Boolean = False) As String
'On Error GoTo err
    Decrypt = ""
    Dim a() As Byte
    a = StrConv(Key, vbFromUnicode)
    Dim b() As Byte
    If NotFile Then
        b = StrConv(FileFrom, vbFromUnicode)
    Else
        b = GetFile(FileFrom)
    End If
    hitung = 0
    For i = 0 To UBound(a)
        hitung = hitung + a(i)
    Next
    Dim d() As Byte
    ReDim d(UBound(b))
    Randomize
    i = 0
    j = 0
    k = 0
    While cek < hitung
        cek = cek + b(k)
        k = k + 1
        i = i + 1
        If i = Len(Key) Then i = 0
    Wend
    Last = b(k - 1)
    For k = k To UBound(b)
        d(j) = (512 - Last + b(k) - a(i)) Mod 256
        Last = b(k)
        i = i + 1
        If i = Len(Key) Then i = 0
        j = j + 1
    Next
    ReDim Preserve d(j - 2)
    str1$ = StrConv(d, vbUnicode)
    Decrypt = str1$
    If FileTo <> "" Then PutFile FileTo, d
    Exit Function
err:
End Function

Function Terbilang(ByVal Angka As String, ByVal MataUang As String) As String
    Dim Grup As String
    Grup = "Ribu   Juta   Milyar Trilyun"
    Dim z() As String
    If pLocal Then
        z = Split(Angka, ",")
    Else
        z = Split(Angka, ".")
    End If
    s = z(0)
    l = Len(s)
    If z(0) <> "0" Then
        While l > 3
            a = val(Mid(s, l - 2, 3))
            b = ""
            If a <> 0 Then
                b = Grup3Digit(val(Mid(s, l - 2, 3)))
                If j <> 0 Then b = b & " " & RTrim(Mid(Grup, (j - 1) * 7 + 1, 7)) & " "
            End If
            If b = "Satu Ribu " Then b = "Seribu "
            Terbilang = b & Terbilang
            l = l - 3
            j = j + 1
        Wend
        a = val(Left(s, l))
        b = ""
        If a <> 0 Then
            b = Grup3Digit(val(a))
            If j <> 0 Then b = b & " " & RTrim(Mid(Grup, (j - 1) * 7 + 1, 7)) & " "
        End If
        If b = "Satu Ribu " Then b = "Seribu "
        Terbilang = LTrim(b & Terbilang)
    Else
        Terbilang = "Nol"
    End If
    If UBound(z) = 1 Then
        Dim Satuan As String
        Satuan = "Nol     Satu    Dua     Tiga    Empat   Lima    Enam    Tujuh   Delapan Sembilan"
        Terbilang = Terbilang & " Koma"
        For i = 1 To Len(z(1))
            Terbilang = Terbilang & " " & RTrim(Mid(Satuan, (val(Mid(z(1), i, 1))) * 8 + 1, 8))
        Next
    End If
    Terbilang = Terbilang & " " & MataUang
End Function

Function Grup3Digit(ByVal Grup3 As Integer) As String
    Dim Satuan As String
    Dim Ratusan As String
    Satuan = "Satu    Dua     Tiga    Empat   Lima    Enam    Tujuh   Delapan Sembilan"
    a = Grup3 \ 100
    b = Grup3 Mod 100
    If a <> 0 Then
        If a = 1 Then
            Grup3Digit = "Seratus"
        Else
            Grup3Digit = RTrim(Mid(Satuan, (a - 1) * 8 + 1, 8)) & " Ratus"
        End If
    End If
    a = b \ 10
    b = b Mod 10
    If a <> 0 Then
        If a = 1 Then
            If b = 0 Then
                Grup3Digit = Grup3Digit & " Sepuluh"
            ElseIf b = 1 Then
                Grup3Digit = Grup3Digit & " Sebelas"
            Else
                Grup3Digit = Grup3Digit & " " & RTrim(Mid(Satuan, (b - 1) * 8 + 1, 8)) & " Belas"
            End If
            Grup3Digit = LTrim(Grup3Digit)
            Exit Function
        Else
            Grup3Digit = Grup3Digit & " " & RTrim(Mid(Satuan, (a - 1) * 8 + 1, 8)) & " Puluh"
        End If
    End If
    a = b Mod 10
    If a <> 0 Then Grup3Digit = Grup3Digit & " " & RTrim(Mid(Satuan, (a - 1) * 8 + 1, 8))
    Grup3Digit = LTrim(Grup3Digit)
End Function

Sub BuatNomor(tNo As Object, tTanggal As Object, ByVal tKet As String, tQuick As Object, ByVal tSQLQuery As String)
    If Not cek_integer(Left(tNo, 5)) Then
        query tSQLQuery
        If IsNull(RS.Fields(0).Value) Then b = 1 Else b = Left(RS.Fields(0).Value, 5) + 1
    Else
        b = CLng(Left(tNo, 5))
    End If
    tNo = zerofill(b, 5) & tKet & Mid(tTanggal, 4)
    If tNo.Tag = "" Then tNo.Tag = tNo
    tQuick = Left(tNo, 5) & "/" & Right(tNo, 2)
End Sub

Function SetNomor(tNo As Object, tTanggal As Object, ByVal tSQLQuery As String) As String
    If Not cek_integer(Left(tNo, 5)) Then
        query tSQLQuery
        If IsNull(RS.Fields(0).Value) Then b = 1 Else b = Left(RS.Fields(0).Value, 5) + 1
    Else
        b = CLng(Left(tNo, 5))
    End If
    tNo = zerofill(b, 5) & "/" & Right(tTanggal, 2)
    SetNomor = tNo
End Function

Sub TDBGridKeyDown(TDBGrid1 As TrueOleDBGrid80.TDBGrid, tKey As Integer)
On Error GoTo err
    If tKey = 13 Then
        tKey = 0
        Dim cCol As Integer
        Dim cRow As Integer
        cCol = TDBGrid1.Col
        cRow = TDBGrid1.Row
        Do
            cCol = cCol + 1
            If cCol = TDBGrid1.Columns.count Then
                If TDBGrid1.AddNewMode = 0 Then cRow = cRow + 1
                cCol = 0
            End If
            If TDBGrid1.Columns(cCol).Visible And Not TDBGrid1.Columns(cCol).Locked Then Exit Do
        Loop
        If TDBGrid1.Row <> cRow Then TDBGrid1.Row = cRow
        TDBGrid1.Col = cCol
    ElseIf tKey = 117 Then
        TDBGrid1.Font.Size = TDBGrid1.Font.Size - 2
    ElseIf tKey = 118 Then
        TDBGrid1.Font.Size = TDBGrid1.Font.Size + 2
    ElseIf tKey = 119 Then
        TDBGrid1.Columns(TDBGrid1.Col).AutoSize
    End If
err:
End Sub

Sub TDBGridLoad(TDBGrid1 As TrueOleDBGrid80.TDBGrid)
    Dim col1 As TrueOleDBGrid80.Columns
    Set col1 = TDBGrid1.Columns
    TDBGrid1.HeadingStyle.Alignment = dbgCenter
    For i = 0 To col1.count - 1
        If col1(i).Tag = "Integer" Then
            col1(i).Alignment = dbgRight
        ElseIf col1(i).Tag = "Decimal" Then
            col1(i).Alignment = dbgRight
            col1(i).NumberFormat = "Standard"
        ElseIf Left(col1(i).Tag, 8) = "Zerofill" Then
            col1(i).NumberFormat = "FormatText Event"
        ElseIf col1(i).Tag = "NoCents" Then
            col1(i).Alignment = dbgRight
            col1(i).NumberFormat = "FormatText Event"
        ElseIf col1(i).Tag = "OK?" Then
            col1(i).Alignment = dbgCenter
            col1(i).ValueItems.Presentation = dbgCheckBox
        ElseIf col1(i).Tag = "Date" Then
            col1(i).NumberFormat = "Edit Mask"
            col1(i).EditMask = "##/##/##"
        ElseIf col1(i).Tag = "DateLong" Then
            col1(i).NumberFormat = "Edit Mask"
            col1(i).EditMask = "##/##/####"
        ElseIf Left(col1(i).Tag, 5) = "ZFill" Then
            col1(i).NumberFormat = String(Mid(col1(i).Tag, 6), "0")
        End If
    Next
    'TDBGridClear TDBGrid1
    TDBGrid1.Style.VerticalAlignment = dbgVertCenter
End Sub

Sub TDBGridKeyPress(TDBGrid1 As TrueOleDBGrid80.TDBGrid, KeyAscii As Integer)
    ColIndex = TDBGrid1.Col
    t = TDBGrid1.Columns(ColIndex).Value
    tp = TDBGrid1.Columns(ColIndex).Tag
    If tp = "Integer" Or Left(tp, 8) = "Zerofill" Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then KeyAscii = 0
    ElseIf tp = "Decimal" Then
        If KeyAscii = 46 And pLocal Then
            KeyAscii = 44
            TDBGrid1.Columns(ColIndex).Value = Replace(t, ".", "")
        End If
        If KeyAscii = 44 And Not pLocal Then
            KeyAscii = 46
            TDBGrid1.Columns(ColIndex).Value = Replace(t, ",", "")
        End If
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            If KeyAscii = 45 Then
                If TDBGrid1.SelStart > 0 Then KeyAscii = 0
            ElseIf KeyAscii = 44 Then
                If InStr(t, ",") <> 0 Then KeyAscii = 0
            ElseIf KeyAscii = 46 Then
                If InStr(t, ".") <> 0 Then KeyAscii = 0
            Else
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Sub TDBGridSetVisible(TDBGrid1 As TrueOleDBGrid80.TDBGrid, ByVal tFields As String, Optional ByVal tVal As Boolean = False)
Dim col1 As TrueOleDBGrid80.Columns
Dim tField() As String
    Set col1 = TDBGrid1.Columns
    tField = Split(tFields, "@")
    For i = 0 To UBound(tField)
        If IsNumeric(tField(i)) Then
            col1(CLng(tField(i))).Visible = tVal
            col1(CLng(tField(i))).AllowSizing = tVal
        Else
            col1(tField(i)).Visible = tVal
            col1(tField(i)).AllowSizing = tVal
        End If
    Next
End Sub

Sub TDBGridSetLock(TDBGrid1 As TrueOleDBGrid80.TDBGrid, ByVal tFields As String, ByVal tVal As Boolean)
Dim col1 As TrueOleDBGrid80.Columns
Dim tField() As String
    Set col1 = TDBGrid1.Columns
    tField = Split(tFields, "@")
    For i = 0 To UBound(tField)
        If IsNumeric(tField(i)) Then
            col1(CLng(tField(i))).Locked = tVal
        Else
            col1(tField(i)).Locked = tVal
        End If
    Next
End Sub

Sub TDBGridSetTag(TDBGrid1 As TrueOleDBGrid80.TDBGrid, ByVal tFields As String, ByVal tVal As String)
Dim col1 As TrueOleDBGrid80.Columns
Dim tField() As String
    Set col1 = TDBGrid1.Columns
    tField = Split(tFields, "@")
    For i = 0 To UBound(tField)
        If IsNumeric(tField(i)) Then
            col1(CLng(tField(i))).Tag = tVal
        Else
            col1(tField(i)).Tag = tVal
        End If
    Next
End Sub

Sub TDBGridClear(TDBGrid1 As TrueOleDBGrid80.TDBGrid)
Dim x As XArrayDB
    Set x = TDBGrid1.Array
    x.ReDim 0, 0, 0, TDBGrid1.Columns.count - 1
    x.DeleteRows 0
    TDBGrid1.Rebind
    If TDBGrid1.AllowAddNew Then TDBGrid1.Row = 0
    For i = 0 To x.UpperBound(2)
        If TDBGrid1.Columns(i).Visible Then
            TDBGrid1.Col = i
            Exit For
        End If
    Next
End Sub

Sub CopyGrid(TDBGrid1 As TrueOleDBGrid80.TDBGrid)
Dim x As New XArrayDB
    Set x = TDBGrid1.Array
    TDBGrid1.Update
    Clipboard.Clear
    b = ""
    If TDBGrid1.SelStartCol = -1 Then
        TDBGrid1.SelStartCol = 0
        TDBGrid1.SelEndCol = TDBGrid1.Columns.count - 1
    End If
    For i = 0 To x.UpperBound(1)
        a = ""
        For j = TDBGrid1.SelStartCol To TDBGrid1.SelEndCol
            a = a & vbTab & x(i, j)
        Next
        b = b & vbCrLf & Mid(a, 2)
    Next
    Clipboard.SetText Mid(b, 3)
End Sub

Sub AutoKonversiStock()
    a = "update m_StockGudang left join m_Stock on m_StockGudang.IdStock=m_Stock.IdStock set QTY=QTY+QTYKecil\QTYPerCarton, QTYKecil=QTYKecil-(QTYKecil\QTYPerCarton)*QTYPerCarton where QTYKecil>0 and QTY<0"
    CN.Execute a, b
    a = "update m_StockGudang left join m_Stock on m_StockGudang.IdStock=m_Stock.IdStock set QTY=QTY+QTYKecil\QTYPerCarton-1, QTYKecil=QTYKecil-(QTYKecil\QTYPerCarton-1)*QTYPerCarton where QTYKecil<0 and QTY>0"
    CN.Execute a, b
End Sub

Function EncDec(ByVal tString As String, ByVal tKey As String) As String
Dim a() As Byte
Dim b() As Byte
Dim c() As Byte
    a = StrConv(tString, vbFromUnicode)
    j = Len(tString) / Len(tKey)
    For i = 0 To j
        tKey = tKey & tKey
    Next
    b = StrConv(tKey, vbFromUnicode)
    ReDim c(UBound(a))
    For i = 0 To UBound(a)
        c(i) = a(i) Xor b(i)
    Next
    Enc = StrConv(c, vbUnicode)
End Function

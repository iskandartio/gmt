Public Class sql
    Private connectionString As String
    Public Sub New(ByVal s As String)
        connectionString = s
    End Sub
#Region "Function"
    Private Function esc(ByVal s As String) As String
        Return Replace(s, "'", "''")
    End Function
#End Region
    
#Region "Query"
#Region "  Agent"
    Public Function qMasterAgent(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String, tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and a.AgentCode like '%" & esc(arr(0)) & "%'"
        If arr(1) <> "" Then tFilter &= " and a.AgentName like '%" & esc(arr(1)) & "%'"
        If arr(2) <> "" Then tFilter &= " and b.AgentCode like '%" & esc(arr(2)) & "%'"
        If arr(3) <> "" Then tFilter &= " and c.AgentCode like '%" & esc(arr(3)) & "%'"
        If arr.Count = 5 Then
            Dim chk As CheckBox = arr(4)
            If chk.Checked Then
                tFilter &= " and b.AgentCode is null"
            End If
        End If
        s = "Select a.AgentID, a.AgentCode, a.AgentName, a.PTKPID, PTKPCode, a.NPWP, mPAM.SalesOfficeID, SalesOfficeCode, a.JabatanID, JabatanCode, b.AgentID as UplineID, b.AgentCode as UplineCode, b.AgentName as UplineName, c.AgentID as PAMID, c.AgentCode as PAMCode, c.AgentName as PAMName from mAgent  a" & _
        " left join mAgent b on b.AgentID=a.UplineID" & _
        " left join mAgent c on c.AgentID=a.PAMID" & _
        " left join mPAM on mPAM.AgentID=a.PAMID" & _
        " left join mSalesOffice on mSalesOffice.SalesOfficeID=mPAM.SalesOfficeID" & _
        " left join mJabatan on mJabatan.JabatanID=a.JabatanID" & _
        " left join mPTKP on mPTKP.PTKPID=a.PTKPID" & _
        " where 1=1" & tFilter
        Dim a As New db(connectionString)
        dt = a.doQuery(s)
        Return dt
    End Function
    Public Function qGetAgentID(ByVal tCode As String) As Integer
        Dim s As String, arr As New ArrayList
        Dim dt As New DataTable
        Dim a As New db(connectionString)
        s = "select AgentID from mAgent where AgentCode=@Code"
        arr.Add(New param("@Code", tCode))
        dt = a.doQuery(s, arr)
        Return dt.Rows(0)(0)
    End Function
#End Region
#Region "  SalesOffice"
    Public Function qMasterSalesOffice(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String, tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and SalesOfficeCode like '%" & esc(arr(0)) & "%'"
        If arr(1) <> "" Then tFilter &= " and CityCode like '%" & esc(arr(1)) & "%'"
        s = "Select SalesOfficeID, SalesOfficeCode, SalesOfficeName, mSalesOffice.CityID, CityCode from mSalesOffice" & _
        " left join mCity on mCity.CityID=mSalesOffice.CityID" & _
        " where 1=1" & tFilter
        Dim a As New db(connectionString)
        dt = a.doQuery(s)
        Return dt
    End Function
#End Region
#Region "  Jabatan"
    Public Function qMasterJabatan(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String, tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and JabatanCode like '%" & esc(arr(0)) & "%'"
        If arr(1) <> "" Then tFilter &= " and TipeJabatanCode like '%" & esc(arr(1)) & "%'"
        s = "Select JabatanID, JabatanCode, JabatanName, mTipeJabatan.TipeJabatanID, TipeJabatanCode from mJabatan" & _
        " left join mTipeJabatan on mTipeJabatan.TipeJabatanID=mJabatan.TipeJabatanID" & _
        " where 1=1" & tFilter
        Dim a As New db(connectionString)
        dt = a.doQuery(s)
        Return dt
    End Function
#End Region
#Region "  City"
    Public Function qMasterCity(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String, tJoin As String = "", tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and CityCode like '%" & esc(arr(0)) & "%'"
        s = "Select CityID, CityCode, CityName from mCity" & _
        " where 1=1" & tFilter
        Dim a As New db(connectionString)
        dt = a.doQuery(s)
        Return dt
    End Function
#End Region
#Region "  TipeJabatan"
    Public Function qMasterTipeJabatan(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String, tJoin As String = "", tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and TipeJabatanCode like '%" & esc(arr(0)) & "%'"
        s = "Select TipeJabatanID, TipeJabatanCode, TipeJabatanName from mTipeJabatan" & _
        " where 1=1" & tFilter
        Dim a As New db(connectionString)
        dt = a.doQuery(s)
        Return dt
    End Function
#End Region
#Region "  PTKP"
    Public Function qMasterPTKP(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String, tJoin As String = "", tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and PTKPCode like '%" & esc(arr(0)) & "%'"
        s = "Select PTKPID, PTKPCode, PTKPName, Tahunan, Bulanan from mPTKP" & _
        " where 1=1" & tFilter
        Dim a As New db(connectionString)
        dt = a.doQuery(s)
        Return dt
    End Function
#End Region
#Region "  TanggalBayar"
    Public Function qMasterTanggalBayar(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String, tJoin As String = "", tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and TglBayar = '" & Format(arr(0), "yyyy/MM/dd") & "'"
        s = "Select TglBayarID, TglBayar from transTglBayar" & _
        " where 1=1" & tFilter & " order by TglBayar"
        Dim a As New db(connectionString)
        dt = a.doQuery(s)
        Return dt
    End Function

    Public Function qMaxTanggalBayar() As DateTime
        Dim s As String = "select max(TglBayar) from transTglBayar"
        Dim a As New db(connectionString)
        Return a.doQuery(s).Rows(0)(0)
    End Function
#End Region
#Region "  TransPajak"
    Public Function qTransPajak(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable, arr2 As New ArrayList
        Dim s As String, tJoin As String = "", tFilter As String = ""
        If arr(0) <> "" Then tFilter &= " and a.AgentCode like '%" & esc(arr(0)) & "%'"
        If arr(1) <> "" Then tFilter &= " and a.AgentName like '%" & esc(arr(1)) & "%'"
        If arr(2) <> "" Then tFilter &= " and b.AgentCode like '%" & esc(arr(2)) & "%'"
        If arr(3) <> "" Then tFilter &= " and c.AgentCode like '%" & esc(arr(3)) & "%'"
        arr2.Clear()
        arr2.Add(New param("@TglBayar", arr(4)))
        s = "select transSummary.AgentID, transSummary.TglBayar, a.AgentCode, a.AgentName, mJabatan.JabatanCode, PTKPCode, a.NPWP" & _
" , AkumulasiPKP, AkumulasiPajak, AkumulasiNPWP" & _
" , Pjk1, Pjk2, Pjk3, Pjk4, AddNPWP, Pjk" & _
" , Pajak1, Pajak2, Pajak3, Pajak4, AdditionalNPWP, Pajak" & _
" , TotalNilai, Additional, Deduction" & _
" from transSummary" & _
" left join mAgent a on a.AgentID=transSummary.AgentID" & _
" left join mAgent b on b.AgentID=a.UplineID" & _
" left join mAgent c on c.AgentID=a.PAMID" & _
" left join mJabatan on mJabatan.JabatanID=a.JabatanID" & _
" left join mPTKP on mPTKP.PTKPID=a.PTKPID" & _
" where TransSummary.TglBayar=@TglBayar"
        Dim a As New db(connectionString)
        dt = a.doQuery(s, arr2)
        Return dt
    End Function

#End Region
#Region "  TransDetailTax"
    Public Function qTransDetailTax(Optional ByVal arr As ArrayList = Nothing) As DataTable
        Dim dt As New DataTable
        Dim s As String
        s = "select TransSummaryID, TransDetailTax.KodeTransID, NamaTransaksi, Nilai from TransDetailTax " & _
" left join mTax on mTax.KodeTransID=TransDetailTax.KodeTransID where TransSummaryID = @TransSummaryID"
        Dim a As New db(connectionString)
        dt = a.doQuery(s, arr)
        Return dt
    End Function

#End Region
#End Region

#Region "SP"
#Region "  Agent"
    Public Sub updatePAM()
        Dim s As String
        s = "exec updatePAM"
        Dim a As New db(connectionString)
        a.execMe(s)
    End Sub
#End Region
    
#End Region
#Region "Update"
#Region "  Agent"
    Public Sub saveAgent(ByRef mdt As DataTable)
        Dim a As New db(connectionString)
        Dim arr As New ArrayList
        Dim dt = mdt.GetChanges
        Dim s As String = ""
        If dt Is Nothing Then Exit Sub
        For Each row As DataRow In dt.Rows
            If row.RowState = DataRowState.Deleted Then Continue For
            arr.Clear()
            arr.Add(New param("@AgentID", row("AgentID")))
            arr.Add(New param("@AgentCode", row("AgentCode")))
            arr.Add(New param("@AgentName", row("AgentName")))
            arr.Add(New param("@UplineID", row("UplineID")))
            arr.Add(New param("@PAMID", row("PAMID")))
            arr.Add(New param("@SalesOfficeID", row("SalesOfficeID")))
            arr.Add(New param("@JabatanID", row("JabatanID")))
            arr.Add(New param("@PTKPID", row("PTKPID")))
            arr.Add(New param("@NPWP", row("NPWP")))
            If row.RowState = DataRowState.Modified Then
                s = "update mAgent set AgentCode=@AgentCode, AgentName=@AgentName, UplineID=@UplineID, PAMID=@PAMID, SalesOfficeID=@SalesOfficeID, JabatanID=@JabatanID, PTKPID=@PTKPID, NPWP=@NPWP  where AgentID=@AgentID"
                a.execMe(s, arr)
            ElseIf row.RowState = DataRowState.Added Then
                s = "insert into mAgent(AgentID, AgentCode, AgentName, UplineID, PAMID, SalesOfficeID, JabatanID, PTKPID, NPWP) values(@AgentID, @AgentCode, @AgentName, @UplineID, @PAMID, @SalesOfficeID, @JabatanID, @PTKPID, @NPWP)"
                a.execMe(s, arr)
            End If
        Next
        mdt.AcceptChanges()
    End Sub
    Public Sub deleteAgent(ByVal tKey As Integer)
        Dim s As String
        s = "delete mAgent where AgentID=" & tKey
        Dim a As New db(connectionString)
        a.execMe(s)
    End Sub
#End Region
#Region "  SalesOffice"
    Public Sub saveSalesOffice(ByRef mdt As DataTable)
        Dim a As New db(connectionString)
        Dim arr As New ArrayList
        Dim dt = mdt.GetChanges
        Dim s As String = ""
        If dt Is Nothing Then Exit Sub
        For Each row As DataRow In dt.Rows
            If row.RowState = DataRowState.Deleted Then Continue For
            arr.Clear()
            arr.Add(New param("@SalesOfficeID", row("SalesOfficeID")))
            arr.Add(New param("@SalesOfficeName", row("SalesOfficeName")))
            arr.Add(New param("@SalesOfficeCode", row("SalesOfficeCode")))
            arr.Add(New param("@CityID", row("CityID")))
            If row.RowState = DataRowState.Modified Then
                s = "update mSalesOffice set SalesOfficeName=@SalesOfficeName, SalesOfficeCode=@SalesOfficeCode, CityID=@CityID  where SalesOfficeID=@SalesOfficeID"
                a.execMe(s, arr)
            ElseIf row.RowState = DataRowState.Added Then
                s = "insert into mSalesOffice(SalesOfficeName, SalesOfficeCode, CityID) values(@SalesOfficeName, @SalesOfficeCode, @CityID)"
                a.execMe(s, arr)
            End If
        Next
        mdt.AcceptChanges()
    End Sub
    Public Sub deleteSalesOffice(ByVal tKey As Integer)
        Dim s As String
        s = "delete mSalesOffice where SalesOfficeID=" & tKey
        Dim a As New db(connectionString)
        a.execMe(s)
    End Sub
#End Region
#Region "  Jabatan"
    Public Sub saveJabatan(ByRef mdt As DataTable)
        Dim a As New db(connectionString)
        Dim arr As New ArrayList
        Dim dt = mdt.GetChanges
        Dim s As String = ""
        If dt Is Nothing Then Exit Sub
        For Each row As DataRow In dt.Rows
            If row.RowState = DataRowState.Deleted Then Continue For
            arr.Clear()
            arr.Add(New param("@JabatanID", row("JabatanID")))
            arr.Add(New param("@JabatanName", row("JabatanName")))
            arr.Add(New param("@JabatanCode", row("JabatanCode")))
            arr.Add(New param("@TipeJabatanID", row("TipeJabatanID")))
            If row.RowState = DataRowState.Modified Then
                s = "update mJabatan set JabatanName=@JabatanName, JabatanCode=@JabatanCode, TipeJabatanID=@TipeJabatanID  where JabatanID=@JabatanID"
            ElseIf row.RowState = DataRowState.Added Then
                s = "insert into mJabatan(JabatanName, JabatanCode, TipeJabatanID) values(@JabatanName, @JabatanCode, @TipeJabatanID)"
            End If
            a.execMe(s, arr)
        Next
        mdt.AcceptChanges()
    End Sub
    Public Sub deleteJabatan(ByVal tKey As Integer)
        Dim s As String
        s = "delete mJabatan where JabatanID=" & tKey
        Dim a As New db(connectionString)
        a.execMe(s)
    End Sub
#End Region
#Region "  City"
    Public Sub saveCity(ByRef mdt As DataTable)
        Dim a As New db(connectionString)
        Dim arr As New ArrayList
        Dim dt = mdt.GetChanges
        Dim s As String = ""
        If dt Is Nothing Then Exit Sub
        For Each row As DataRow In dt.Rows
            If row.RowState = DataRowState.Deleted Then Continue For
            arr.Clear()
            arr.Add(New param("@CityID", row("CityID")))
            arr.Add(New param("@CityName", row("CityName")))
            arr.Add(New param("@CityCode", row("CityCode")))
            If row.RowState = DataRowState.Modified Then
                s = "update mCity set CityName=@CityName, CityCode=@CityCode where CityID=@CityID"
            ElseIf row.RowState = DataRowState.Added Then
                s = "insert into mCity(CityID, CityName, CityCode) values(@CityID, @CityName, @CityCode)"
            End If
            a.execMe(s, arr)
        Next
        mdt.AcceptChanges()
    End Sub
    Public Sub deleteCity(ByVal tKey As Integer)
        Dim s As String
        s = "delete mCity where CityID=" & tKey
        Dim a As New db(connectionString)
        a.execMe(s)
    End Sub
#End Region
#Region "  PTKP"
    Public Sub savePTKP(ByRef mdt As DataTable)
        Dim a As New db(connectionString)
        Dim arr As New ArrayList
        Dim dt = mdt.GetChanges
        Dim s As String = ""
        If dt Is Nothing Then Exit Sub
        For Each row As DataRow In dt.Rows
            If row.RowState = DataRowState.Deleted Then Continue For
            arr.Clear()
            arr.Add(New param("@PTKPID", row("PTKPID")))
            arr.Add(New param("@PTKPName", row("PTKPName")))
            arr.Add(New param("@PTKPCode", row("PTKPCode")))
            arr.Add(New param("@Tahunan", row("Tahunan")))
            arr.Add(New param("@Bulanan", row("Bulanan")))
            If row.RowState = DataRowState.Modified Then
                s = "update mPTKP set PTKPName=@PTKPName, PTKPCode=@PTKPCode, Tahunan=@Tahunan, Bulanan=@Bulanan where PTKPID=@PTKPID"
            ElseIf row.RowState = DataRowState.Added Then
                s = "insert into mPTKP(PTKPID, PTKPName, PTKPCode, Tahunan, Bulanan) values(@PTKPID, @PTKPName, @PTKPCode, @Tahunan, @Bulanan)"
            End If
            a.execMe(s, arr)
        Next
        mdt.AcceptChanges()
    End Sub
    Public Sub deletePTKP(ByVal tKey As Integer)
        Dim s As String
        s = "delete mPTKP where PTKPID=" & tKey
        Dim a As New db(connectionString)
        a.execMe(s)
    End Sub
#End Region
#Region "  TipeJabatan"
    Public Sub saveTipeJabatan(ByRef mdt As DataTable)
        Dim a As New db(connectionString)
        Dim arr As New ArrayList
        Dim dt = mdt.GetChanges
        Dim s As String = ""
        If dt Is Nothing Then Exit Sub
        For Each row As DataRow In dt.Rows
            If row.RowState = DataRowState.Deleted Then Continue For
            arr.Clear()
            arr.Add(New param("@TipeJabatanID", row("TipeJabatanID")))
            arr.Add(New param("@TipeJabatanName", row("TipeJabatanName")))
            arr.Add(New param("@TipeJabatanCode", row("TipeJabatanCode")))
            If row.RowState = DataRowState.Modified Then
                s = "update mTipeJabatan set TipeJabatanName=@TipeJabatanName, TipeJabatanCode=@TipeJabatanCode where TipeJabatanID=@TipeJabatanID"
            ElseIf row.RowState = DataRowState.Added Then
                s = "insert into mTipeJabatan(TipeJabatanName, TipeJabatanCode, TipeTipeJabatanID) values(@TipeJabatanName, @TipeJabatanCode, @TipeTipeJabatanID)"
            End If
            a.execMe(s, arr)
        Next
        mdt.AcceptChanges()
    End Sub
    Public Sub deleteTipeJabatan(ByVal tKey As Integer)
        Dim s As String
        s = "delete mTipeJabatan where TipeJabatanID=" & tKey
        Dim a As New db(connectionString)
        a.execMe(s)
    End Sub
#End Region

#End Region
    
End Class

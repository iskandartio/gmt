Imports System.Windows.Forms
Imports System.Net.Sockets

Public Class MDIForm

    Private Sub MDIParent1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim f As New frmLogin
        If f.ShowDialog(Me) <> Windows.Forms.DialogResult.Yes Then
            Me.Dispose()
        End If
        gUser = f.gUser
        Me.SetupMenu(menuMaster, gUser, "M")
        Me.SetupMenu(menuTransaction, gUser, "T")
        Me.SetupMenu(menuReport, gUser, "R")
        Me.MenuStrip.MdiWindowListItem = Me.WindowToolStripMenuItem
        Dim s As String
        Dim s2() As String
        's = "NoSC TanggalSC Kode NamaCustomer MataUang LamaKontrak WaktuPembayaran NilaiKontrak Keterangan Status Pengupdate WaktuUpdate Disetujui ShortSC DP NamaMarketing NamaCustomerSC"
        's = """" & Replace(s, " ", """,""") & """"
        ''s = Replace(s, " ", ", ")
        ''s = "@" & Replace(s, " ", ", @")
        's2 = Split(s, " ")
        's = ""
        'For i As Integer = 0 To s2.Count - 1
        '    s = s & ", " & s2(i) & "=@" & s2(i)
        'Next
        'Dim b As New SQLRefreshView
        'b.RefreshView()
    End Sub

    Private Sub SetupMenu(ByVal m As ToolStripMenuItem, ByVal u As String, ByVal t As String)
        Dim s As String
        Dim a As New db
        Dim dt As DataTable
        
        s = "select c.* from mUserModule a" & _
" left join mModuleForm b on a.ModuleID=b.ModuleID" & _
" left join mForm c on c.FormID=b.FormID" & _
" where a.UID='" & u & "' and c.FormType='" & t & "'"
        dt = a.doQuery(s)
        For i As Integer = 0 To dt.Rows.Count - 1
            m.DropDownItems.Add(dt.Rows(i)("FormName"))
            m.DropDownItems(i).Tag = dt.Rows(i)("FormID")
        Next
    End Sub

    Private Sub menu_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles menuMaster.DropDownItemClicked, menuTransaction.DropDownItemClicked, menuReport.DropDownItemClicked
        Dim FormName As String
        FormName = e.ClickedItem.Tag
        Dim T As Type = Type.GetType([GetType].Namespace & "." & FormName, False)
        If T Is Nothing Then
            For Each T In System.Reflection.Assembly.Load(FormName).GetTypes
                If T.Name = FormName Then Exit For
            Next
        End If
        If T Is Nothing Then
            MsgBox("Form not found")
            Exit Sub
        End If

        Dim f As Form = CType(Activator.CreateInstance(T), Form)
        Dim i As Integer
        For i = 0 To My.Application.OpenForms.Count - 1
            If f.Name = My.Application.OpenForms(i).Name Then
                My.Application.OpenForms(i).BringToFront()
                Exit Sub
            End If
        Next
        f.StartPosition = FormStartPosition.Manual
        f.Left = (My.Application.OpenForms.Count - 1) * 25
        f.Top = (My.Application.OpenForms.Count - 1) * 25
        f.MdiParent = Me
        f.Text = e.ClickedItem.Text
        f.Show()
    End Sub



    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next

    End Sub


    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Dispose()
    End Sub
End Class

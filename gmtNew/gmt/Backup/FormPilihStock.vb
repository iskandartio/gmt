﻿Public Class FormPilihStock
    Dim mdt As DataTable
    Dim mDefaultFilter As String
    Dim mFields As String
    Private Sub FindMe(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, btnFind.Click
        Find()
    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        ModControl.ClearControl(grpFilter)
    End Sub
    Private Sub _KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyDown
        If e.KeyCode = Keys.Down OrElse e.KeyCode = Keys.Up Then
            dg.Focus()
        End If
    End Sub

    Private Sub Find()
        Dim z As New List(Of Param)

        z.Add(New Param("KodeBarang", txtCode.Text))
        z.Add(New Param("NoWarna", txtNoWarna.Text))

        mdt = New SQLmStock().qStock(z, mDefaultFilter, mFields)
        dg.DataSource = mdt
        If mdt.Rows.Count = 0 Then
            MsgBox("No Record Found")
        End If
        ModControl.setGridAutoInvisible(dg)
        ModControl.setGridStyle(dg, 0)
    End Sub

    Overrides Sub Choose()
        Me.DialogResult = Windows.Forms.DialogResult.OK
        row = mdt.Rows(dg.CurrentCell.RowIndex)
        Me.Hide()
    End Sub

    Public Sub New(Optional ByVal zDefaultFilter As List(Of Param) = Nothing, Optional ByVal tFields As String = "")

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        mDefaultFilter = setFilter(zDefaultFilter)
        mFields = tFields
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
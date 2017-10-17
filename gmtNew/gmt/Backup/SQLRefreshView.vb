Public Class SQLRefreshView
    Dim a As New db
    Public Function RefreshView() As Boolean
        Dim s As String
        s = "alter view vmPenomoran as select a.* from mPenomoran a inner join mCurrentTipe b on a.Tipe=b.Tipe"
        a.execMe(s)
        s = "alter view vmStock as select a.* from mStock a inner join mCurrentTipe b on a.Tipe=b.Tipe"
        a.execMe(s)
        s = "alter view vtSC as select a.* from tSC a inner join mCurrentTipe b on a.Tipe=b.Tipe"
        a.execMe(s)
        s = "alter view vtSPP as select a.* from tSPP a inner join mCurrentTipe b on a.Tipe=b.Tipe"
        a.execMe(s)
        's = "alter view vtSTT as select a.* from tSTT a inner join mCurrentTipe b on a.Tipe=b.Tipe"
        'a.execMe(s)
        Return True
    End Function
End Class

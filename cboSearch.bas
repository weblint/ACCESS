Option Compare Database
Option Explicit


Private Sub Kombinationsfeld9_AfterUpdate()
 ' Den mit dem Steuerelement übereinstimmenden Datensatz suchen.
    Dim rs As Recordset

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[KategorieID] = " & Str(Me![Kombinationsfeld9])
    If rs.NoMatch Then
            Me!Kombinationsfeld9 = Null
    End If
        Me.Bookmark = rs.Bookmark
    Set rs = Nothing
End Sub

Attribute VB_Name = "RemoveHyperlinksFromSelection"
Sub RemoveHyperlinksFromSelection()
    Dim cell As Range
    For Each cell In Selection
        If cell.Hyperlinks.Count > 0 Then
            cell.Hyperlinks.Delete
        End If
    Next cell
End Sub

Attribute VB_Name = "HideAllExceptActiveSheet"
Sub HideAllExceptActiveSheet()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> ActiveSheet.Name Then
            ws.Visible = xlSheetHidden
        End If
    Next ws

End Sub

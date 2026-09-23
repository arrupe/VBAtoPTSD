Attribute VB_Name = "UnhideALlSheets"
Sub UnhideALlSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub

Attribute VB_Name = "UnprotectAllSheets"
Sub UnprotectAllSheets()

    Dim ws As Worksheet
    Dim Password As String

    Password = "Test123"

    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=Password
    Next ws

End Sub

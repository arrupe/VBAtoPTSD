Attribute VB_Name = "ProtectAllSheets"
'This code will protect all worksheets at once
Sub ProtectAllSheets()

    Dim ws As Worksheet
    Dim Password As String

    Password = "Test123"    'Replace Test123 with your desired password

    For Each ws In ThisWorkbook.Worksheets
        ws.Protect _
            Password:=Password, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True

    Next ws

End Sub


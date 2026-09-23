Attribute VB_Name = "LockFormulaCellsAllSheets"
'This code will lock all formula cells on all worksheets
Sub LockFormulaCellsAllSheets()

    Dim ws As Worksheet
    Dim Password As String

    Password = "Test123"    'Replace with your desired password

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each ws In ThisWorkbook.Worksheets

        ws.Unprotect Password:=Password

        'Unlock all cells
        ws.Cells.Locked = False

        'Lock only formula cells
        On Error Resume Next
        ws.Cells.SpecialCells(xlCellTypeFormulas).Locked = True
        On Error GoTo 0

        'Protect sheet while allowing common user actions
        ws.Protect _
            Password:=Password, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True

    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "All formula cells have been locked and all sheets protected.", vbInformation

End Sub


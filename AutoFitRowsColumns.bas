Attribute VB_Name = "AutoFitRowsColumns"
'This code will AutoFit all rows and columns on the active worksheet
Sub AutoFitRowsColumns()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    With ActiveSheet
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

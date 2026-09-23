Attribute VB_Name = "RefreshEverything"
'This code will refresh all workbook data, queries, tables, pivots, and calculations
Sub RefreshEverything()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "Refreshing workbook..."

    On Error Resume Next

    'Refresh all queries, connections, tables, pivots, and Power Pivot model
    ThisWorkbook.RefreshAll

    'Wait for async queries to finish
    Application.CalculateUntilAsyncQueriesDone

    'Refresh all pivot caches
    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc

    'Recalculate workbook
    Application.CalculateFullRebuild

    On Error GoTo 0

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Workbook refresh completed successfully.", vbInformation

End Sub

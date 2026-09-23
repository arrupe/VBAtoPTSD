Attribute VB_Name = "UnhideRowsColumns"
'This code will unhide all rows and columns on the active worksheet
Sub UnhideRowsColumns()

    Application.ScreenUpdating = False

    With ActiveSheet
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
        .Range("A1").Select
    End With

    Application.ScreenUpdating = True

End Sub

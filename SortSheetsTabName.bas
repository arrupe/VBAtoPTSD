Attribute VB_Name = "SortSheetsTabName"
'This code will sort the worksheets alphabetically
Sub SortSheetsTabName()

    Dim ShCount As Long
    Dim i As Long
    Dim j As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ShCount = Sheets.Count

    For i = 1 To ShCount - 1
        For j = i + 1 To ShCount

            If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
                Sheets(j).Move Before:=Sheets(i)
            End If

        Next j
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

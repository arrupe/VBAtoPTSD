Attribute VB_Name = "ConvertAllFormulasToValues"
Sub ConvertAllFormulasToValues()
    Dim ws As Worksheet
    
    ' Speed up macro execution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through every worksheet in the active workbook
    For Each ws In ActiveWorkbook.Worksheets
        ' Reassign the used range to its own values to remove formulas instantly
        ws.UsedRange.Value = ws.UsedRange.Value
    Next ws
    
    ' Restore Excel settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "All formulas have been converted to values on every sheet!", vbInformation
End Sub

Attribute VB_Name = "UnhideAndResetWorkbookView"
Sub UnhideAndResetWorkbookView()

    Dim ws As Worksheet
    Dim StartingSheet As Worksheet
    
    'Remember current sheet
    Set StartingSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    
    'Unhide all sheets and reset view
    For Each ws In ThisWorkbook.Worksheets
        
        'Unhide sheet (including VeryHidden)
        ws.Visible = xlSheetVisible
        
        'Activate sheet to change zoom and selection
        ws.Activate
        
        'Set zoom to 100%
        ActiveWindow.Zoom = 100
        
        'Move cursor to A1
        ws.Range("A1").Select
        
    Next ws
    
    'Return to original sheet
    StartingSheet.Activate
    StartingSheet.Range("A1").Select
    ActiveWindow.Zoom = 100
    
    Application.ScreenUpdating = True

End Sub

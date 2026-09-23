Attribute VB_Name = "ClearWorkbookErrorHighlights"
'Removes highlighting applied by AuditWorkbookErrors
Sub ClearWorkbookErrorHighlights()

    Const AUDIT_STYLE As String = "Workbook Audit Error"

    Dim ws As Worksheet
    Dim Cell As Range
    Dim ClearedCount As Long

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Removing workbook audit highlighting..."

    For Each ws In ThisWorkbook.Worksheets

        Application.StatusBar = _
            "Checking worksheet: " & ws.Name

        For Each Cell In ws.UsedRange.Cells
            If Cell.Style = AUDIT_STYLE Then
                Cell.Style = "Normal"
                ClearedCount = ClearedCount + 1
            End If
        Next Cell

    Next ws

    MsgBox _
        Format(ClearedCount, "#,##0") & _
        " audit highlights were removed.", _
        vbInformation, _
        "Clear Workbook Audit"

CleanExit:

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:

    MsgBox _
        "The audit highlighting could not be completely removed." & _
        vbCrLf & _
        "Error " & err.Number & ": " & err.Description, _
        vbCritical, _
        "Clear Workbook Audit"

    Resume CleanExit

End Sub


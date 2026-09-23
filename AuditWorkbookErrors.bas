Attribute VB_Name = "AuditWorkbookErrors"
Option Explicit

'Highlights Excel error cells across all worksheets
'and checks the workbook for broken defined names
Sub AuditWorkbookErrors()

    Const AUDIT_STYLE As String = "Workbook Audit Error"

    Dim ws As Worksheet
    Dim ErrorCells As Range
    Dim Cell As Range
    Dim WorkbookName As Name
    Dim ErrorCount As Long
    Dim FormulaErrorCount As Long
    Dim ConstantErrorCount As Long
    Dim BrokenNameCount As Long
    Dim ProtectedSheetCount As Long
    Dim ErrorSummary As String
    Dim OriginalCalculation As XlCalculation
    Dim OriginalSheet As Object
    Dim AuditStyle As Style

    On Error GoTo ErrorHandler

    Set OriginalSheet = ActiveSheet
    OriginalCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Auditing workbook for errors..."

    'Create or retrieve the reusable audit style
    On Error Resume Next
    Set AuditStyle = ThisWorkbook.Styles(AUDIT_STYLE)
    On Error GoTo ErrorHandler

    If AuditStyle Is Nothing Then
        Set AuditStyle = ThisWorkbook.Styles.Add(AUDIT_STYLE)

        With AuditStyle
            .IncludeFont = True
            .IncludeBorder = True
            .IncludePatterns = True
            .IncludeAlignment = False
            .IncludeNumber = False
            .IncludeProtection = False

            .Font.Color = RGB(156, 0, 6)
            .Font.Bold = True
            .Interior.Color = RGB(255, 199, 206)

            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = RGB(156, 0, 6)
                .Weight = xlThin
            End With

            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = RGB(156, 0, 6)
                .Weight = xlThin
            End With

            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(156, 0, 6)
                .Weight = xlThin
            End With

            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = RGB(156, 0, 6)
                .Weight = xlThin
            End With
        End With
    End If

    For Each ws In ThisWorkbook.Worksheets

        Application.StatusBar = _
            "Auditing worksheet: " & ws.Name

        If ws.ProtectContents Then
            ProtectedSheetCount = ProtectedSheetCount + 1
        End If

        Set ErrorCells = Nothing

        'Find errors produced by formulas
        On Error Resume Next
        Set ErrorCells = ws.UsedRange.SpecialCells( _
            Type:=xlCellTypeFormulas, _
            Value:=xlErrors)
        On Error GoTo ErrorHandler

        If Not ErrorCells Is Nothing Then
            For Each Cell In ErrorCells.Cells
                Cell.Style = AUDIT_STYLE
                FormulaErrorCount = FormulaErrorCount + 1
            Next Cell
        End If

        Set ErrorCells = Nothing

        'Find error values entered or stored as constants
        On Error Resume Next
        Set ErrorCells = ws.UsedRange.SpecialCells( _
            Type:=xlCellTypeConstants, _
            Value:=xlErrors)
        On Error GoTo ErrorHandler

        If Not ErrorCells Is Nothing Then
            For Each Cell In ErrorCells.Cells
                Cell.Style = AUDIT_STYLE
                ConstantErrorCount = ConstantErrorCount + 1
            Next Cell
        End If

    Next ws

    'Check workbook-level and worksheet-level defined names
    For Each WorkbookName In ThisWorkbook.Names
        If InStr(1, WorkbookName.RefersTo, "#REF!", _
            vbTextCompare) > 0 Then

            BrokenNameCount = BrokenNameCount + 1
        End If
    Next WorkbookName

    ErrorCount = FormulaErrorCount + ConstantErrorCount

    ErrorSummary = _
        "Workbook audit completed." & vbCrLf & vbCrLf & _
        "Error cells: " & Format(ErrorCount, "#,##0") & vbCrLf & _
        "   Formula errors: " & _
            Format(FormulaErrorCount, "#,##0") & vbCrLf & _
        "   Constant error values: " & _
            Format(ConstantErrorCount, "#,##0") & vbCrLf & _
        "Broken defined names: " & _
            Format(BrokenNameCount, "#,##0") & vbCrLf & _
        "Protected worksheets reviewed: " & _
            Format(ProtectedSheetCount, "#,##0")

    If ErrorCount = 0 And BrokenNameCount = 0 Then
        MsgBox ErrorSummary & vbCrLf & vbCrLf & _
               "No workbook errors were found.", _
               vbInformation, _
               "Workbook Audit"
    Else
        MsgBox ErrorSummary & vbCrLf & vbCrLf & _
               "Error cells were highlighted in light red.", _
               vbExclamation, _
               "Workbook Audit"
    End If

CleanExit:

    On Error Resume Next

    If Not OriginalSheet Is Nothing Then
        OriginalSheet.Activate
    End If

    Application.StatusBar = False
    Application.Calculation = OriginalCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    On Error GoTo 0
    Exit Sub

ErrorHandler:

    MsgBox _
        "The workbook audit could not be completed." & vbCrLf & _
        "Error " & err.Number & ": " & err.Description, _
        vbCritical, _
        "Workbook Audit"

    Resume CleanExit

End Sub


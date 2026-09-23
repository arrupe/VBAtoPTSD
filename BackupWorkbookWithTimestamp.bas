Attribute VB_Name = "BackupWorkbookWithTimestamp"
Sub BackupWorkbookWithTimestamp()
    Dim wb As Workbook
    Dim fPath As String, fName As String, fExt As String
    Dim baseName As String, timeStamp As String
    
    Set wb = ThisWorkbook ' Use ActiveWorkbook if running from a personal macro workbook
    
    ' Check if workbook has been saved before
    If wb.Path = "" Then
        MsgBox "Please save the main workbook first.", vbExclamation
        Exit Sub
    End If
    
    ' Extract file path, name, and extension
    fPath = wb.Path & "\"
    fName = wb.Name
    fExt = Mid(fName, InStrRev(fName, "."))
    baseName = Left(fName, Len(fName) - Len(fExt))
    
    ' Create timestamp (e.g., 2026-09-23_1340)
    timeStamp = Format(Now, "yyyy-mm-dd_hhnn")
    
    ' Save a copy with timestamp
    wb.SaveCopyAs fPath & baseName & "_" & timeStamp & fExt
    
    MsgBox "Backup created successfully!", vbInformation, "Backup Complete"
End Sub


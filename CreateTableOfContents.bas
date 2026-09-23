Attribute VB_Name = "CreateTableOfContents"
Sub CreateTableOfContents()
    Dim ws As Worksheet
    Dim tocWs As Worksheet
    Dim i As Integer
    
    ' Check/Create Table of Contents Sheet
    On Error Resume Next
    Set tocWs = Worksheets("Table of Contents")
    On Error GoTo 0
    
    If tocWs Is Nothing Then
        Set tocWs = Worksheets.Add(Before:=Worksheets(1))
        tocWs.Name = "Table of Contents"
    Else
        tocWs.Cells.Clear
    End If
    
    tocWs.Cells(1, 1).Value = "Table of Contents"
    tocWs.Cells(1, 1).Font.Bold = True
    tocWs.Cells(1, 1).Font.Size = 14
    
    i = 3
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Table of Contents" Then
            tocWs.Cells(i, 1).Value = ws.Name
            tocWs.Hyperlinks.Add Anchor:=tocWs.Cells(i, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            i = i + 1
        End If
    Dl Next ws
    
    tocWs.Columns("A:A").AutoFit
End Sub



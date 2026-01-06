Option Explicit

Public Sub WriteModelSummary(wb As Workbook, root As String)

    Dim ws As Worksheet, cCount As Long, fCount As Long
    Dim chCount As Long, vbaCount As Long
    Dim rng As Range

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Not rng Is Nothing Then fCount = fCount + rng.Cells.Count
        chCount = chCount + ws.ChartObjects.Count
        Set rng = Nothing
        On Error GoTo 0
    Next

    vbaCount = wb.VBProject.VBComponents.Count

    Dim f As Integer
    f = FreeFile
    Open root & "\Model_Summary.txt" For Output As #f

    Print #f, "MODEL SUMMARY"
    Print #f, "Worksheets: " & wb.Worksheets.Count
    Print #f, "Formula Cells: " & fCount
    Print #f, "Charts: " & chCount
    Print #f, "VBA Components: " & vbaCount

    Close #f

End Sub


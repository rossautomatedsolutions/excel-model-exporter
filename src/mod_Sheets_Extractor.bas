Option Explicit

Public Sub ExportSheetLayouts(wb As Workbook, outPath As String)

    Dim ws As Worksheet, f As Integer

    For Each ws In wb.Worksheets
        f = FreeFile
        Open outPath & "\" & CleanName(ws.Name) & "_Layout.txt" For Output As #f
        Print #f, "Sheet: " & ws.Name
        Print #f, "UsedRange: " & ws.UsedRange.Address
        Print #f, "Rows: " & ws.UsedRange.Rows.Count
        Print #f, "Columns: " & ws.UsedRange.Columns.Count
        Close #f
    Next

End Sub

Public Sub ExportFormulas(wb As Workbook, outPath As String)

    Dim ws As Worksheet, rng As Range, c As Range, f As Integer

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextSheet

        f = FreeFile
        Open outPath & "\" & CleanName(ws.Name) & "_Formulas.txt" For Output As #f
        For Each c In rng.Cells
            Print #f, c.Address(False, False) & " = " & c.Formula
        Next
        Close #f

NextSheet:
        Set rng = Nothing
    Next

End Sub


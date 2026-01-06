Option Explicit

Public Sub ExportCharts(wb As Workbook, outPath As String)

    Dim ws As Worksheet, co As ChartObject, ch As Chart
    Dim s As Series, f As Integer

    For Each ws In wb.Worksheets
        If ws.ChartObjects.Count = 0 Then GoTo NextSheet

        f = FreeFile
        Open outPath & "\" & CleanName(ws.Name) & "_Charts.txt" For Output As #f

        For Each co In ws.ChartObjects
            Set ch = co.Chart
            Print #f, "Chart: " & co.Name
            Print #f, "Type: " & ch.ChartType
            For Each s In ch.SeriesCollection
                Print #f, "  Series: " & s.Name
            Next
            Print #f, ""
        Next

        Close #f
NextSheet:
    Next

End Sub


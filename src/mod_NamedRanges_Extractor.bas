Option Explicit

Public Sub ExportNamedRanges(wb As Workbook, rootFolder As String)

    Dim f As Integer
    Dim n As Name
    Dim outPath As String
    Dim hasAny As Boolean

    outPath = rootFolder & "\Named_Ranges.txt"

    f = FreeFile
    Open outPath For Output As #f

    Print #f, "=== NAMED RANGES ==="
    Print #f, "Workbook: " & wb.Name
    Print #f, ""

    On Error Resume Next
    For Each n In wb.Names
        hasAny = True
        ' Name = RefersTo (formula or range)
        Print #f, n.Name & " = " & n.RefersTo
        If Err.Number <> 0 Then
            Print #f, n.Name & " = [ERROR reading RefersTo: " & Err.Description & "]"
            Err.Clear
        End If
    Next n
    On Error GoTo 0

    Close #f

    ' If no named ranges, delete the file to keep output clean
    If Not hasAny Then
        On Error Resume Next
        Kill outPath
        On Error GoTo 0
    End If

End Sub


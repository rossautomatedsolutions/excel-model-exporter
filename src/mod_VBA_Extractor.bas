Option Explicit

Public Sub ExportVBAModules(wb As Workbook, outPath As String)

    Dim comp As Object
    Dim f As Integer
    Dim lineCount As Long
    Dim root As String

    root = Left(outPath, InStrRev(outPath, "\") - 1)

    On Error GoTo VBABlocked

    For Each comp In wb.VBProject.VBComponents

        lineCount = comp.CodeModule.CountOfLines

        f = FreeFile
        Open outPath & "\" & CleanName(comp.Name) & ".txt" For Output As #f

        Print #f, "' VBA COMPONENT: " & comp.Name
        Print #f, "' Type: " & comp.Type
        Print #f, ""

        If lineCount > 0 Then
            Print #f, comp.CodeModule.Lines(1, lineCount)
        Else
            Print #f, "' <No code in this module>"
        End If

        Close #f
 
    Next comp

    Exit Sub

VBABlocked:
    LogMessage root, "VBA export skipped or partially failed: " & Err.Description
    Err.Clear

End Sub


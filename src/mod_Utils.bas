Option Explicit

Public Function PickWorkbook() As Workbook

    Dim fd As FileDialog
    Dim path As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .AllowMultiSelect = False
        .Title = "Select Excel File to Export"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        If .Show <> -1 Then Exit Function
        path = .SelectedItems(1)
    End With

    Set PickWorkbook = Workbooks.Open(path, ReadOnly:=True)

End Function

Public Function CleanName(s As String) As String

    Dim badChars As Variant, c As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    CleanName = s
    For Each c In badChars
        CleanName = Replace(CleanName, c, "_")
    Next

End Function


Option Explicit

Public Function CreateExportRootFolder(wb As Workbook) As String

    Dim fso As Object
    Dim basePath As String
    Dim root As String
    Dim cleanWbName As String
    Dim dotPos As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ALWAYS use a local, writable path
    basePath = GetLocalExportBasePath

    If Not fso.FolderExists(basePath) Then
        fso.CreateFolder basePath
    End If

    dotPos = InStrRev(wb.Name, ".")
    If dotPos > 0 Then
        cleanWbName = CleanName(Left(wb.Name, dotPos - 1))
    Else
        cleanWbName = CleanName(wb.Name)
    End If

    root = basePath & "\" & cleanWbName & "_Export"

    EnsureFolder fso, root
    EnsureFolder fso, root & "\VBA"
    EnsureFolder fso, root & "\Sheets"
    EnsureFolder fso, root & "\Formulas"
    EnsureFolder fso, root & "\Charts"

    CreateExportRootFolder = root

End Function


Public Function GetLocalExportBasePath() As String
    GetLocalExportBasePath = Environ$("USERPROFILE") & "\Documents\Excel_Model_Exports"
End Function

Public Sub EnsureFolder(fso As Object, path As String)

    Debug.Print "EnsureFolder called with path:"
    Debug.Print "[" & path & "]"
    Debug.Print "Length: " & Len(path)
    Debug.Print "EndsWithSlash: " & (Right(path, 1) = "\")
    Debug.Print "Exists?: " & fso.FolderExists(path)
    Debug.Print "----"

    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If

End Sub



Public Sub LogMessage(root As String, msg As String)

    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set f = fso.OpenTextFile(root & "\Run_Log.txt", 8, True)
    f.WriteLine Format(Now(), "yyyy-mm-dd hh:nn:ss") & " - " & msg
    f.Close

End Sub

Public Sub WriteReadmeTemplate(root As String)

    Dim f As Integer
    f = FreeFile

    Open root & "\README_HOW_TO_USE.txt" For Output As #f

    Print #f, "EXCEL MODEL EXPORT — INTERPRETATION GUIDE"
    Print #f, ""
    Print #f, "Purpose:"
    Print #f, "This folder contains a structural and logical extraction of an Excel model."
    Print #f, "It enables full understanding without access to the original workbook."
    Print #f, ""
    Print #f, "Folder Overview:"
    Print #f, "/Sheets   ? worksheet layout summaries"
    Print #f, "/Formulas ? all cell-level formulas"
    Print #f, "/Charts   ? chart metadata and data bindings"
    Print #f, "/VBA      ? VBA modules (optional)"
    Print #f, ""
    Print #f, "Interpretation Order:"
    Print #f, "1) Sheets"
    Print #f, "2) Formulas"
    Print #f, "3) Charts"
    Print #f, "4) VBA (if present)"

    Close #f

End Sub


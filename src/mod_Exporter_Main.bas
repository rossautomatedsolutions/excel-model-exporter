Option Explicit

Public Sub Run_Excel_Model_Export()

    Dim wb As Workbook
    Dim exportRoot As String
    Dim startTime As Date

    Set wb = PickWorkbook()
    If wb Is Nothing Then Exit Sub

    startTime = Now
    exportRoot = CreateExportRootFolder(wb)

    WriteReadmeTemplate exportRoot
    LogMessage exportRoot, "Export started"

    ExportVBAModules wb, exportRoot & "\VBA"
    ExportSheetLayouts wb, exportRoot & "\Sheets"
    ExportFormulas wb, exportRoot & "\Formulas"
    ExportCharts wb, exportRoot & "\Charts"
    ExportNamedRanges wb, exportRoot
    WriteModelSummary wb, exportRoot

    LogMessage exportRoot, "Export completed"
    LogMessage exportRoot, "Elapsed seconds: " & Format((Now - startTime) * 86400, "0.00")

    wb.Close SaveChanges:=False
    MsgBox "Export complete:" & vbCrLf & exportRoot, vbInformation

End Sub


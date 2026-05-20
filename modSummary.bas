Option Explicit

Public Sub SummaryRefresh()
    Dim summaryService As clsSummaryService
    Dim errorMessage As String
    Dim success As Boolean

    Set summaryService = New clsSummaryService

    success = summaryService.RefreshSummary(ThisWorkbook, errorMessage)

    If success Then
        MsgBox "Summary refreshed successfully.", vbInformation, "Summary"
    Else
        MsgBox errorMessage, vbExclamation, "Summary"
    End If
End Sub
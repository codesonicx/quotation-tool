Option Explicit

' =========================================================
' Text helpers
' =========================================================

' Normalizes text for safer comparisons and validations.
Public Function NormalizeText(ByVal value As String) As String
    NormalizeText = Trim$(value)
End Function


' =========================================================
' General worksheet helpers
' =========================================================

' Returns the last used row on the worksheet.
' Returns 0 when the sheet has no content.
Public Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastUsed As Range

    Set lastUsed = ws.Cells.Find(What:="*", _
                                 LookIn:=xlFormulas, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious)

    If lastUsed Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = lastUsed.row
    End If
End Function


' Returns the worksheet currently active in Excel.
Public Function GetTargetWorksheet() As Worksheet
    Set GetTargetWorksheet = ActiveSheet
End Function


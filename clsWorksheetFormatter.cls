Option Explicit

' =========================================================
' clsWorksheetFormatter
'
' Purpose:
' Handles worksheet visual formatting and row grouping.
'
' Responsibilities:
' - apply assembly row styles
' - clear inherited formatting
' - apply row grouping
' =========================================================

Public Sub ResetInsertedRows(ByVal targetRows As Range)

    targetRows.ClearFormats
    targetRows.ClearOutline

End Sub

Public Sub FormatAssemblyHeaderRow(ByVal ws As Worksheet, _
                                   ByVal rowIndex As Long)

    With ws.Rows(rowIndex).Interior
        .Pattern = xlSolid
        .Color = RGB( _
            COLOR_ASSEMBLY_GRAY_R, _
            COLOR_ASSEMBLY_GRAY_G, _
            COLOR_ASSEMBLY_GRAY_B)
    End With

End Sub

Public Sub GroupAssemblyChildRows(ByVal ws As Worksheet, _
                                  ByVal firstChildRow As Long, _
                                  ByVal lastChildRow As Long)

    If lastChildRow >= firstChildRow Then
        ws.Rows(firstChildRow & ":" & lastChildRow).Group
    End If

End Sub

Public Sub FormatCustomItemRow(ByVal targetRow As Range)

    targetRow.ClearFormats
    targetRow.ClearOutline
    targetRow.Interior.Pattern = xlNone

End Sub
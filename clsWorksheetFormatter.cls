Option Explicit

' =========================================================
' clsWorksheetFormatter
'
' Purpose:
' Handles worksheet visual formatting, number formats,
' and row grouping.
'
' Responsibilities:
' - reset inserted rows
' - apply section row styles
' - apply assembly row styles
' - apply custom item row styles
' - apply estimate number formats
' - apply row grouping
' =========================================================

Public Sub ResetInsertedRows(ByVal targetRows As Range)

    targetRows.ClearFormats
    targetRows.ClearOutline

End Sub

Public Sub FormatSectionHeaderRow(ByVal ws As Worksheet, _
                                  ByVal rowIndex As Long)

    With ws.Rows(rowIndex).Interior
        .Pattern = xlSolid
        .ColorIndex = COLOR_INDEX_SECTION_BLUE
    End With

    ApplyEstimateNumberFormats ws, rowIndex

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

    ApplyEstimateNumberFormats ws, rowIndex

End Sub

Public Sub FormatCustomItemRow(ByVal targetRow As Range)

    targetRow.ClearFormats
    targetRow.ClearOutline
    targetRow.Interior.Pattern = xlNone

    ApplyEstimateNumberFormats targetRow.Worksheet, targetRow.Row

End Sub

Public Sub GroupAssemblyChildRows(ByVal ws As Worksheet, _
                                  ByVal firstChildRow As Long, _
                                  ByVal lastChildRow As Long)

    If lastChildRow >= firstChildRow Then
        ws.Rows(firstChildRow & ":" & lastChildRow).Group
    End If

End Sub

Public Sub RefreshAssemblyChildGrouping(ByVal ws As Worksheet, _
                                        ByVal firstChildRow As Long, _
                                        ByVal lastChildRow As Long)

    If lastChildRow < firstChildRow Then Exit Sub

    ws.Rows(firstChildRow & ":" & lastChildRow).ClearOutline
    ws.Rows(firstChildRow & ":" & lastChildRow).Group

End Sub

Public Sub ApplyEstimateNumberFormats(ByVal ws As Worksheet, _
                                      ByVal rowIndex As Long)

    ' Assembly quantity columns
    ws.Range( _
        ws.Cells(rowIndex, COL_ITEM_DES), _
        ws.Cells(rowIndex, COL_ITEM_BUILD) _
    ).NumberFormat = FORMAT_INTEGER

    ' Component quantity columns
    ws.Range( _
        ws.Cells(rowIndex, COL_COMP_LIST), _
        ws.Cells(rowIndex, COL_SECT_BUILD) _
    ).NumberFormat = FORMAT_INTEGER

    ' Markup percentages
    ws.Range( _
        ws.Cells(rowIndex, COL_MATERIAL_MARKUP), _
        ws.Cells(rowIndex, COL_LABOR_MARKUP) _
    ).NumberFormat = FORMAT_PERCENT

    ' Hours
    ws.Range( _
        ws.Cells(rowIndex, COL_DESIGN_HOURS), _
        ws.Cells(rowIndex, COL_STARTUP_HOURS) _
    ).NumberFormat = FORMAT_DECIMAL

    ' Material rate
    ws.Cells(rowIndex, COL_MATERIAL_RATE).NumberFormat = FORMAT_CURRENCY

    ' Budgeted costs
    ws.Range( _
        ws.Cells(rowIndex, COL_BUDGET_DESIGN), _
        ws.Cells(rowIndex, COL_BUDGET_TOTAL) _
    ).NumberFormat = FORMAT_CURRENCY

    ' Costs with markup
    ws.Range( _
        ws.Cells(rowIndex, COL_MARKUP_DESIGN), _
        ws.Cells(rowIndex, COL_MARKUP_TOTAL) _
    ).NumberFormat = FORMAT_CURRENCY

    ' Percentage column
    ws.Cells(rowIndex, COL_PERCENT).NumberFormat = FORMAT_PERCENT

End Sub
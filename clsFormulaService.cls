Option Explicit

' =========================================================
' clsFormulaService
'
' Purpose:
' Writes formulas into custom component rows.
' =========================================================

Public Sub WriteCustomComponentFormulas(ByVal ws As Worksheet, _
                                        ByVal rowIndex As Long, _
                                        ByVal customData As Collection)

    WriteComponentQuantities ws, rowIndex
    WriteMarkupValues ws, rowIndex
    WriteHourFormulas ws, rowIndex, customData
    WriteBudgetCostFormulas ws, rowIndex
    WriteMarkupCostFormulas ws, rowIndex
    WriteTotals ws, rowIndex

End Sub

Private Sub WriteComponentQuantities(ByVal ws As Worksheet, _
                                     ByVal rowIndex As Long)

    ws.Cells(rowIndex, COL_COMP_LIST).Formula = "=" & _
        ws.Cells(rowIndex - 1, COL_COMP_LIST).Address(False, False) & "+1"

    ws.Cells(rowIndex, COL_SECT_DES).Formula = "=" & _
        ws.Cells(rowIndex, COL_COMP_DES).Address(False, False) & _
        "*OFFSET(" & _
        ws.Cells(rowIndex, COL_SECT_DES).Address(False, False) & _
        ",-" & ws.Cells(rowIndex, COL_COMP_LIST).Address(False, False) & _
        ",-6)"

    ws.Cells(rowIndex, COL_SECT_BUILD).Formula = "=" & _
        ws.Cells(rowIndex, COL_COMP_BUILD).Address(False, False) & _
        "*OFFSET(" & _
        ws.Cells(rowIndex, COL_SECT_BUILD).Address(False, False) & _
        ",-" & ws.Cells(rowIndex, COL_COMP_LIST).Address(False, False) & _
        ",-6)"

End Sub

Private Sub WriteMarkupValues(ByVal ws As Worksheet, _
                              ByVal rowIndex As Long)

    ws.Cells(rowIndex, COL_MATERIAL_MARKUP).Formula = "=Drivers!$B$11"
    ws.Cells(rowIndex, COL_LABOR_MARKUP).Formula = "=Drivers!$B$12"

End Sub

Private Sub WriteHourFormulas(ByVal ws As Worksheet, _
                              ByVal rowIndex As Long, _
                              ByVal customData As Collection)

    ws.Cells(rowIndex, COL_DESIGN_HOURS).Formula = "=" & _
        ws.Cells(rowIndex, COL_SECT_DES).Address(False, False) & _
        "*" & CDbl(customData("des"))

    ws.Cells(rowIndex, COL_CONSTRUCTION_HOURS).Formula = "=" & _
        ws.Cells(rowIndex, COL_SECT_BUILD).Address(False, False) & _
        "*" & CDbl(customData("constr"))

    ws.Cells(rowIndex, COL_DEBUG_HOURS).Formula = "=" & _
        ws.Cells(rowIndex, COL_SECT_BUILD).Address(False, False) & _
        "*" & CDbl(customData("debug"))

    ws.Cells(rowIndex, COL_INSTALLATION_HOURS).Formula = "=" & _
        ws.Cells(rowIndex, COL_SECT_BUILD).Address(False, False) & _
        "*" & CDbl(customData("inst"))

    ws.Cells(rowIndex, COL_STARTUP_HOURS).Formula = "=" & _
        ws.Cells(rowIndex, COL_SECT_BUILD).Address(False, False) & _
        "*" & CDbl(customData("start"))

End Sub

Private Sub WriteBudgetCostFormulas(ByVal ws As Worksheet, _
                                    ByVal rowIndex As Long)

    ws.Cells(rowIndex, COL_BUDGET_DESIGN).Formula = "=" & _
        ws.Cells(rowIndex, COL_DESIGN_HOURS).Address(False, False) & "*Drivers!$D$4"

    ws.Cells(rowIndex, COL_BUDGET_CONSTRUCTION).Formula = "=" & _
        ws.Cells(rowIndex, COL_CONSTRUCTION_HOURS).Address(False, False) & "*Drivers!$D$4"

    ws.Cells(rowIndex, COL_BUDGET_DEBUG).Formula = "=" & _
        ws.Cells(rowIndex, COL_DEBUG_HOURS).Address(False, False) & "*Drivers!$D$4"

    ws.Cells(rowIndex, COL_BUDGET_INSTALLATION).Formula = "=" & _
        ws.Cells(rowIndex, COL_INSTALLATION_HOURS).Address(False, False) & "*Drivers!$D$4"

    ws.Cells(rowIndex, COL_BUDGET_STARTUP).Formula = "=" & _
        ws.Cells(rowIndex, COL_STARTUP_HOURS).Address(False, False) & "*Drivers!$D$4"

    ws.Cells(rowIndex, COL_BUDGET_MATERIAL).Formula = "=" & _
        ws.Cells(rowIndex, COL_MATERIAL_RATE).Address(False, False) & _
        "*" & _
        ws.Cells(rowIndex, COL_SECT_BUILD).Address(False, False)

End Sub

Private Sub WriteMarkupCostFormulas(ByVal ws As Worksheet, _
                                    ByVal rowIndex As Long)

    ws.Cells(rowIndex, COL_MARKUP_DESIGN).Formula = "=" & _
        ws.Cells(rowIndex, COL_BUDGET_DESIGN).Address(False, False) & _
        "*(" & ws.Cells(rowIndex, COL_LABOR_MARKUP).Address(False, False) & "+1)"

    ws.Cells(rowIndex, COL_MARKUP_CONSTRUCTION).Formula = "=" & _
        ws.Cells(rowIndex, COL_BUDGET_CONSTRUCTION).Address(False, False) & _
        "*(" & ws.Cells(rowIndex, COL_LABOR_MARKUP).Address(False, False) & "+1)"

    ws.Cells(rowIndex, COL_MARKUP_DEBUG).Formula = "=" & _
        ws.Cells(rowIndex, COL_BUDGET_DEBUG).Address(False, False) & _
        "*(" & ws.Cells(rowIndex, COL_LABOR_MARKUP).Address(False, False) & "+1)"

    ws.Cells(rowIndex, COL_MARKUP_INSTALLATION).Formula = "=" & _
        ws.Cells(rowIndex, COL_BUDGET_INSTALLATION).Address(False, False) & _
        "*(" & ws.Cells(rowIndex, COL_LABOR_MARKUP).Address(False, False) & "+1)"

    ws.Cells(rowIndex, COL_MARKUP_STARTUP).Formula = "=" & _
        ws.Cells(rowIndex, COL_BUDGET_STARTUP).Address(False, False) & _
        "*(" & ws.Cells(rowIndex, COL_LABOR_MARKUP).Address(False, False) & "+1)"

    ws.Cells(rowIndex, COL_MARKUP_MATERIAL).Formula = "=" & _
        ws.Cells(rowIndex, COL_BUDGET_MATERIAL).Address(False, False) & _
        "*(" & ws.Cells(rowIndex, COL_MATERIAL_MARKUP).Address(False, False) & "+1)"

End Sub

Private Sub WriteTotals(ByVal ws As Worksheet, _
                        ByVal rowIndex As Long)

    ws.Cells(rowIndex, COL_BUDGET_TOTAL).Formula = "=SUM(" & _
        ws.Cells(rowIndex, COL_BUDGET_DESIGN).Address(False, False) & ":" & _
        ws.Cells(rowIndex, COL_BUDGET_MATERIAL).Address(False, False) & ")"

    ws.Cells(rowIndex, COL_MARKUP_TOTAL).Formula = "=SUM(" & _
        ws.Cells(rowIndex, COL_MARKUP_DESIGN).Address(False, False) & ":" & _
        ws.Cells(rowIndex, COL_MARKUP_MATERIAL).Address(False, False) & ")"

End Sub

Public Sub WriteSectionTotalFormulas(ByVal ws As Worksheet, _
                                     ByVal sectionRow As Long, _
                                     ByRef assemblyRows() As Long)

    Dim colIndex As Long
    Dim i As Long
    Dim formulaText As String
    Dim columnLetter As String

    For colIndex = COL_SECTION_TOTAL_START To COL_SECTION_TOTAL_END

        columnLetter = Split(ws.Cells(1, colIndex).Address(False, False), "1")(0)

        formulaText = "=SUM("

        For i = LBound(assemblyRows) To UBound(assemblyRows)

            If i > LBound(assemblyRows) Then
                formulaText = formulaText & ","
            End If

            formulaText = formulaText & columnLetter & assemblyRows(i)

        Next i

        formulaText = formulaText & ")"

        ws.Cells(sectionRow, colIndex).Formula = formulaText

    Next colIndex

End Sub

Public Sub WriteAssemblyTotalFormulas(ByVal ws As Worksheet, _
                                      ByVal assemblyRow As Long, _
                                      ByVal firstChildRow As Long, _
                                      ByVal lastChildRow As Long)

    Dim colIndex As Long
    Dim columnLetter As String

    If lastChildRow < firstChildRow Then Exit Sub

    For colIndex = COL_SECTION_TOTAL_START To COL_SECTION_TOTAL_END
        columnLetter = Split(ws.Cells(1, colIndex).Address(False, False), "1")(0)

        ws.Cells(assemblyRow, colIndex).Formula = "=SUM(" & _
            columnLetter & firstChildRow & ":" & _
            columnLetter & lastChildRow & ")"
    Next colIndex

End Sub
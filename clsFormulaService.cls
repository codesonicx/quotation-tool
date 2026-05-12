Option Explicit

' =========================================================
' clsFormulaService
'
' Purpose:
' Writes formulas and calculated values into target worksheet rows.
' =========================================================

Public Sub WriteDesignFormula(ByVal wsTarget As Worksheet, _
                              ByVal targetRow As Long, _
                              ByVal targetCol As Long, _
                              ByVal sourceCol As Long, _
                              ByVal multiplier As Double)

    If targetCol > 0 And sourceCol > 0 Then
        wsTarget.Cells(targetRow, targetCol).Formula = "=" & _
            wsTarget.Cells(targetRow, sourceCol).Address(False, False) & "*" & multiplier
    End If

End Sub

Public Sub WriteBuildFormula(ByVal wsTarget As Worksheet, _
                             ByVal targetRow As Long, _
                             ByVal targetCol As Long, _
                             ByVal sourceCol As Long, _
                             ByVal multiplier As Double)

    If targetCol > 0 And sourceCol > 0 Then
        wsTarget.Cells(targetRow, targetCol).Formula = "=" & _
            wsTarget.Cells(targetRow, sourceCol).Address(False, False) & "*" & multiplier
    End If

End Sub

Public Sub WriteComponentListNumber(ByVal wsTarget As Worksheet, _
                                    ByVal targetRow As Long, _
                                    ByVal compListCol As Long)

    If targetRow > 1 And IsNumeric(wsTarget.Cells(targetRow - 1, compListCol).Value) Then
        wsTarget.Cells(targetRow, compListCol).Value = CLng(wsTarget.Cells(targetRow - 1, compListCol).Value) + 1
    Else
        wsTarget.Cells(targetRow, compListCol).Value = 1
    End If

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
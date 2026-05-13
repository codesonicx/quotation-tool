Option Explicit

' =========================================================
' clsGeneratedAssemblyService
'
' Purpose:
' Inserts generated assembly blocks directly from VBA-defined data.
'
' This service does NOT copy from Assembly List.
'
' Responsibilities:
' - insert one gray assembly row
' - insert child component rows
' - format rows
' - write base child values
' - write child formulas
' - write assembly total formulas
' - refresh child row grouping
' =========================================================

Public Function InsertGeneratedAssembly(ByVal wsTarget As Worksheet, _
                                        ByVal insertRow As Long, _
                                        ByVal assemblyName As String, _
                                        ByVal itemDes As Double, _
                                        ByVal itemBuild As Double, _
                                        ByVal children As Collection, _
                                        ByRef assemblyRow As Long, _
                                        ByRef firstChildRow As Long, _
                                        ByRef lastChildRow As Long, _
                                        ByRef errorMessage As String) As Boolean
    Dim worksheetFormatter As clsWorksheetFormatter
    Dim formulaService As clsFormulaService

    Dim rowsToInsert As Long
    Dim childIndex As Long
    Dim targetRow As Long

    errorMessage = vbNullString

    Set worksheetFormatter = New clsWorksheetFormatter
    Set formulaService = New clsFormulaService

    rowsToInsert = children.Count + 1

    assemblyRow = insertRow
    firstChildRow = assemblyRow + 1
    lastChildRow = assemblyRow + children.Count

    wsTarget.Rows(assemblyRow & ":" & lastChildRow).Insert Shift:=xlShiftDown

    worksheetFormatter.ResetInsertedRows wsTarget.Rows(assemblyRow & ":" & lastChildRow)
    worksheetFormatter.FormatAssemblyHeaderRow wsTarget, assemblyRow

    wsTarget.Cells(assemblyRow, COL_ITEM).Value = assemblyName
    wsTarget.Cells(assemblyRow, COL_ITEM_DES).Value = itemDes
    wsTarget.Cells(assemblyRow, COL_ITEM_BUILD).Value = itemBuild

    For childIndex = 1 To children.Count
        targetRow = firstChildRow + childIndex - 1

        worksheetFormatter.FormatCustomItemRow wsTarget.Rows(targetRow)

        WriteGeneratedChildRow _
            wsTarget, _
            targetRow, _
            children(childIndex)
    Next childIndex

    formulaService.WriteAssemblyTotalFormulas _
        wsTarget, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow

    worksheetFormatter.RefreshAssemblyChildGrouping _
        wsTarget, _
        firstChildRow, _
        lastChildRow

    InsertGeneratedAssembly = True
End Function

Public Function CreateChild(ByVal componentName As String, _
                            ByVal compDes As Double, _
                            ByVal compBuild As Double, _
                            ByVal designHours As Double, _
                            ByVal constructionHours As Double, _
                            ByVal debugHours As Double, _
                            ByVal installHours As Double, _
                            ByVal startHours As Double, _
                            ByVal materialRate As Double) As Collection
    Dim child As Collection
    Set child = New Collection

    child.Add componentName, "name"

    child.Add compDes, "compDes"
    child.Add compBuild, "compBuild"

    child.Add designHours, "des"
    child.Add constructionHours, "constr"
    child.Add debugHours, "debug"
    child.Add installHours, "inst"
    child.Add startHours, "start"

    child.Add materialRate, "mat"

    Set CreateChild = child
End Function

Public Function CreateDesignOnlyChild(ByVal componentName As String, _
                                      ByVal designHours As Double) As Collection

    Set CreateDesignOnlyChild = CreateChild( _
        componentName, _
        1, _
        1, _
        designHours, _
        0, _
        0, _
        0, _
        0, _
        0)

End Function

Public Function CreateMaterialOnlyChild(ByVal componentName As String, _
                                        ByVal materialRate As Double) As Collection

    Set CreateMaterialOnlyChild = CreateChild( _
        componentName, _
        1, _
        1, _
        0, _
        0, _
        0, _
        0, _
        0, _
        materialRate)

End Function

Public Function CreateStartupOnlyChild(ByVal componentName As String, _
                                       ByVal startupHours As Double, _
                                       ByVal materialRate As Double) As Collection

    Set CreateStartupOnlyChild = CreateChild( _
        componentName, _
        1, _
        1, _
        0, _
        0, _
        0, _
        0, _
        startupHours, _
        materialRate)

End Function

Private Sub WriteGeneratedChildRow(ByVal wsTarget As Worksheet, _
                                   ByVal targetRow As Long, _
                                   ByVal child As Collection)
    Dim formulaService As clsFormulaService

    Set formulaService = New clsFormulaService

    wsTarget.Cells(targetRow, COL_COMPONENT).Value = CStr(child("name"))
    wsTarget.Cells(targetRow, COL_COMP_DES).Value = CDbl(child("compDes"))
    wsTarget.Cells(targetRow, COL_COMP_BUILD).Value = CDbl(child("compBuild"))
    wsTarget.Cells(targetRow, COL_MATERIAL_RATE).Value = CDbl(child("mat"))

    formulaService.WriteCustomComponentFormulas _
        wsTarget, _
        targetRow, _
        child
End Sub
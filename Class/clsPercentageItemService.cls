Option Explicit

' =========================================================
' clsPercentItemService
'
' Purpose:
' Inserts a percentage-based child item into an existing assembly.
'
' Behavior:
' - Adds a child row at the end of the selected assembly.
' - Calculates percentage amounts per selected budget column.
' - Writes each result into its own budget column:
'       S = Design
'       T = Construction
'       U = Debug/Tryout
'       V = Installation
'       X = Materia
' - Writes Y = SUM of S up to X column
' - Refreshes assembly totals, grouping, section totals, and grand totals.
' =========================================================

Public Function InsertPercentItemIntoAssembly(ByVal wsTarget As Worksheet, _
                                              ByVal sectionName As String, _
                                              ByVal assemblyName As String, _
                                              ByVal componentName As String, _
                                              ByVal percentAmount As Double, _
                                              ByVal includeDesign As Boolean, _
                                              ByVal includeConstruction As Boolean, _
                                              ByVal includeDebug As Boolean, _
                                              ByVal includeMaterial As Boolean, _
                                              ByVal includeInstallation As Boolean, _
                                              ByRef errorMessage As String) As Boolean
    Dim assemblyFinder As clsAssemblyFinder
    Dim sectionService As clsSectionService
    Dim worksheetFormatter As clsWorksheetFormatter
    Dim formulaService As clsFormulaService

    Dim insertRow As Long
    Dim assemblyRow As Long
    Dim assemblyEndRow As Long
    Dim firstChildRow As Long
    Dim sourceLastChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString

    Set assemblyFinder = New clsAssemblyFinder
    Set sectionService = New clsSectionService
    Set worksheetFormatter = New clsWorksheetFormatter
    Set formulaService = New clsFormulaService

    assemblyRow = assemblyFinder.FindAssemblyHeaderRow(wsTarget, sectionName, assemblyName)
    If assemblyRow = 0 Then
        errorMessage = "Assembly '" & assemblyName & "' was not found in section '" & sectionName & "'."
        InsertPercentItemIntoAssembly = False
        Exit Function
    End If

    insertRow = assemblyFinder.FindCustomComponentInsertRow(wsTarget, sectionName, assemblyName)
    If insertRow = 0 Then
        errorMessage = "Could not determine insert row for assembly '" & assemblyName & "'."
        InsertPercentItemIntoAssembly = False
        Exit Function
    End If

    firstChildRow = assemblyRow + 1
    sourceLastChildRow = insertRow - 1

    wsTarget.Rows(insertRow).Insert Shift:=xlShiftDown
    worksheetFormatter.FormatCustomItemRow wsTarget.Rows(insertRow)

    WritePercentItemRow _
        wsTarget, _
        insertRow, _
        firstChildRow, _
        sourceLastChildRow, _
        componentName, _
        percentAmount, _
        includeDesign, _
        includeConstruction, _
        includeDebug, _
        includeMaterial, _
        includeInstallation

    assemblyEndRow = assemblyFinder.FindAssemblyEndRow(wsTarget, sectionName, assemblyName)
    lastChildRow = assemblyEndRow

    formulaService.WriteAssemblyTotalFormulas _
        wsTarget, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow

    worksheetFormatter.RefreshAssemblyChildGrouping _
        wsTarget, _
        firstChildRow, _
        lastChildRow

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertPercentItemIntoAssembly = False
        Exit Function
    End If

    InsertPercentItemIntoAssembly = True
End Function

Private Sub WritePercentItemRow(ByVal ws As Worksheet, _
                                ByVal rowIndex As Long, _
                                ByVal firstSourceRow As Long, _
                                ByVal lastSourceRow As Long, _
                                ByVal componentName As String, _
                                ByVal percentAmount As Double, _
                                ByVal includeDesign As Boolean, _
                                ByVal includeConstruction As Boolean, _
                                ByVal includeDebug As Boolean, _
                                ByVal includeMaterial As Boolean, _
                                ByVal includeInstallation As Boolean)

    ws.Cells(rowIndex, COL_COMP_LIST).value = GetNextComponentListNumber(ws, rowIndex)
    ws.Cells(rowIndex, COL_COMPONENT).Value = BuildPercentComponentName(percentAmount, componentName)

    ws.Cells(rowIndex, COL_COMP_DES).value = 1
    ws.Cells(rowIndex, COL_COMP_BUILD).value = 1
    ws.Cells(rowIndex, COL_SECT_DES).value = 0
    ws.Cells(rowIndex, COL_SECT_BUILD).value = 0

    ws.Cells(rowIndex, COL_MATERIAL_MARKUP).Formula = "=Drivers!$B$11"
    ws.Cells(rowIndex, COL_LABOR_MARKUP).Formula = "=Drivers!$B$12"

    ws.Cells(rowIndex, COL_DESIGN_HOURS).value = 0
    ws.Cells(rowIndex, COL_CONSTRUCTION_HOURS).value = 0
    ws.Cells(rowIndex, COL_DEBUG_HOURS).value = 0
    ws.Cells(rowIndex, COL_INSTALLATION_HOURS).value = 0
    ws.Cells(rowIndex, COL_STARTUP_HOURS).value = 0
    ws.Cells(rowIndex, COL_MATERIAL_RATE).value = 0

    ' These are the only budget columns we override.
    ws.Cells(rowIndex, COL_BUDGET_DESIGN).Formula = BuildSingleBudgetPercentFormula( _
        ws, firstSourceRow, lastSourceRow, COL_BUDGET_DESIGN, percentAmount, includeDesign)

    ws.Cells(rowIndex, COL_BUDGET_CONSTRUCTION).Formula = BuildSingleBudgetPercentFormula( _
        ws, firstSourceRow, lastSourceRow, COL_BUDGET_CONSTRUCTION, percentAmount, includeConstruction)

    ws.Cells(rowIndex, COL_BUDGET_DEBUG).Formula = BuildSingleBudgetPercentFormula( _
        ws, firstSourceRow, lastSourceRow, COL_BUDGET_DEBUG, percentAmount, includeDebug)

    ws.Cells(rowIndex, COL_BUDGET_INSTALLATION).Formula = BuildSingleBudgetPercentFormula( _
        ws, firstSourceRow, lastSourceRow, COL_BUDGET_INSTALLATION, percentAmount, includeInstallation)

    ws.Cells(rowIndex, COL_BUDGET_MATERIAL).Formula = BuildSingleBudgetPercentFormula( _
    ws, firstSourceRow, lastSourceRow, COL_BUDGET_MATERIAL, percentAmount, includeMaterial)

    ' Leave the remaining budget buckets as normal zero values.
    ws.Cells(rowIndex, COL_BUDGET_STARTUP).value = 0

    WritePercentItemCostFormulas ws, rowIndex

End Sub

Private Function GetNextComponentListNumber(ByVal ws As Worksheet, _
                                            ByVal rowIndex As Long) As Long
    If rowIndex > 1 And IsNumeric(ws.Cells(rowIndex - 1, COL_COMP_LIST).value) Then
        GetNextComponentListNumber = CLng(ws.Cells(rowIndex - 1, COL_COMP_LIST).value) + 1
    Else
        GetNextComponentListNumber = 1
    End If
End Function

Private Function BuildSingleBudgetPercentFormula(ByVal ws As Worksheet, _
                                                 ByVal firstSourceRow As Long, _
                                                 ByVal lastSourceRow As Long, _
                                                 ByVal sourceColumn As Long, _
                                                 ByVal percentAmount As Double, _
                                                 ByVal includeColumn As Boolean) As String
    Dim columnLetter As String
    Dim percentValue As Double

    If Not includeColumn Then
        BuildSingleBudgetPercentFormula = "=0"
        Exit Function
    End If

    If lastSourceRow < firstSourceRow Then
        BuildSingleBudgetPercentFormula = "=0"
        Exit Function
    End If

    percentValue = percentAmount / 100
    columnLetter = Split(ws.Cells(1, sourceColumn).Address(False, False), "1")(0)

    BuildSingleBudgetPercentFormula = "=SUM(" & _
        columnLetter & firstSourceRow & ":" & _
        columnLetter & lastSourceRow & ")*" & CStr(percentValue)
End Function

Private Sub WritePercentItemCostFormulas(ByVal ws As Worksheet, _
                                         ByVal rowIndex As Long)

    ' Normal budget total behavior.
    ws.Cells(rowIndex, COL_BUDGET_TOTAL).Formula = "=SUM(" & _
        ws.Cells(rowIndex, COL_BUDGET_DESIGN).Address(False, False) & ":" & _
        ws.Cells(rowIndex, COL_BUDGET_MATERIAL).Address(False, False) & ")"

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

    ' Normal markup total behavior.
    ws.Cells(rowIndex, COL_MARKUP_TOTAL).Formula = "=SUM(" & _
        ws.Cells(rowIndex, COL_MARKUP_DESIGN).Address(False, False) & ":" & _
        ws.Cells(rowIndex, COL_MARKUP_MATERIAL).Address(False, False) & ")"

    ' ws.Cells(rowIndex, COL_PERCENT).Formula = "=(" & _
    '     ws.Cells(rowIndex, COL_MARKUP_TOTAL).Address(False, False) & "-" & _
    '     ws.Cells(rowIndex, COL_BUDGET_TOTAL).Address(False, False) & ")/+" & _
    '     ws.Cells(rowIndex, COL_MARKUP_TOTAL).Address(False, False)

End Sub

Private Function BuildPercentComponentName(ByVal percentAmount As Double, _
                                           ByVal componentName As String) As String
    BuildPercentComponentName = "[" & CStr(percentAmount) & "%] - " & componentName
End Function
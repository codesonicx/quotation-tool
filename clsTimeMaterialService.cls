Option Explicit

' =========================================================
' clsTimeMaterialService
'
' Purpose:
' Inserts the T&M section with Travel and Expense and Labor
' assemblies.
' =========================================================

Private Const SECTION_BASE_NAME As String = "T&M"
Private Const ASSEMBLY_TRAVEL As String = "Travel and Expense"
Private Const ASSEMBLY_LABOR As String = "Labor"

Public Function InsertTimeMaterialSection(ByVal wsTarget As Worksheet, _
                                          ByVal data As Collection, _
                                          ByRef errorMessage As String) As Boolean
    Dim sectionService As clsSectionService
    Dim sectionFinder As clsSectionFinder

    Dim sectionName As String
    Dim sectionRow As Long
    Dim insertRow As Long

    errorMessage = vbNullString

    Set sectionService = New clsSectionService
    Set sectionFinder = New clsSectionFinder

    sectionName = sectionService.GetUniqueSectionName(wsTarget, SECTION_BASE_NAME)

    If Not sectionService.AddSection(wsTarget, sectionName, errorMessage) Then
        InsertTimeMaterialSection = False
        Exit Function
    End If

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Could not locate created section '" & sectionName & "'."
        InsertTimeMaterialSection = False
        Exit Function
    End If

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    InsertTravelAssembly wsTarget, insertRow, data

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    InsertLaborAssembly wsTarget, insertRow, data

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertTimeMaterialSection = False
        Exit Function
    End If

    InsertTimeMaterialSection = True
End Function

Private Sub InsertTravelAssembly(ByVal ws As Worksheet, _
                                 ByVal insertRow As Long, _
                                 ByVal data As Collection)
    Dim worksheetFormatter As clsWorksheetFormatter
    Dim formulaService As clsFormulaService

    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long
    Dim rowIndex As Long

    Set worksheetFormatter = New clsWorksheetFormatter
    Set formulaService = New clsFormulaService

    assemblyRow = insertRow

    ws.Rows(assemblyRow & ":" & assemblyRow + 6).Insert Shift:=xlShiftDown

    ' Reset inherited section formatting/grouping.
    ws.Rows(assemblyRow & ":" & assemblyRow + 6).ClearFormats
    ws.Rows(assemblyRow & ":" & assemblyRow + 6).ClearOutline

    worksheetFormatter.FormatAssemblyHeaderRow ws, assemblyRow

    ws.Cells(assemblyRow, COL_ITEM).Value = ASSEMBLY_TRAVEL
    ws.Cells(assemblyRow, COL_ITEM_DES).Value = 1
    ws.Cells(assemblyRow, COL_ITEM_BUILD).Value = 1

    firstChildRow = assemblyRow + 1
    lastChildRow = assemblyRow + 6

    For rowIndex = firstChildRow To lastChildRow
        worksheetFormatter.FormatCustomItemRow ws.Rows(rowIndex)
    Next rowIndex

    WriteTravelChildRow ws, firstChildRow, 1, "Per Diem", CDbl(data("perDiem"))
    WriteTravelChildRow ws, firstChildRow + 1, 2, "Hotel", CDbl(data("hotelRate"))
    WriteTravelChildRow ws, firstChildRow + 2, 3, "Car Rent", CDbl(data("carRentRate"))
    WriteTravelChildRow ws, firstChildRow + 3, 4, "Fuel", CDbl(data("fuelRate"))
    WriteTravelChildRow ws, firstChildRow + 4, 5, "Mileage", CDbl(data("mileageRate"))
    WriteTravelChildRow ws, firstChildRow + 5, 6, "Air", CDbl(data("flightRate"))

    formulaService.WriteAssemblyTotalFormulas ws, assemblyRow, firstChildRow, lastChildRow
    worksheetFormatter.RefreshAssemblyChildGrouping ws, firstChildRow, lastChildRow
End Sub

Private Sub InsertLaborAssembly(ByVal ws As Worksheet, _
                                ByVal insertRow As Long, _
                                ByVal data As Collection)
    Dim worksheetFormatter As clsWorksheetFormatter
    Dim formulaService As clsFormulaService

    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long
    Dim rowIndex As Long

    Set worksheetFormatter = New clsWorksheetFormatter
    Set formulaService = New clsFormulaService

    assemblyRow = insertRow

    ws.Rows(assemblyRow & ":" & assemblyRow + 2).Insert Shift:=xlShiftDown

    ' Reset inherited section formatting/grouping.
    ws.Rows(assemblyRow & ":" & assemblyRow + 2).ClearFormats
    ws.Rows(assemblyRow & ":" & assemblyRow + 2).ClearOutline

    worksheetFormatter.FormatAssemblyHeaderRow ws, assemblyRow

    ws.Cells(assemblyRow, COL_ITEM).Value = ASSEMBLY_LABOR
    ws.Cells(assemblyRow, COL_ITEM_DES).Value = 1
    ws.Cells(assemblyRow, COL_ITEM_BUILD).Value = 1

    firstChildRow = assemblyRow + 1
    lastChildRow = assemblyRow + 2

    For rowIndex = firstChildRow To lastChildRow
        worksheetFormatter.FormatCustomItemRow ws.Rows(rowIndex)
    Next rowIndex

    WriteLaborChildRow ws, firstChildRow, 1, "Standard Hours", CDbl(data("standardHours")), CDbl(data("standardRate"))
    WriteLaborChildRow ws, firstChildRow + 1, 2, "Overtime Hours", CDbl(data("overtimeHours")), CDbl(data("overtimeRate"))

    formulaService.WriteAssemblyTotalFormulas ws, assemblyRow, firstChildRow, lastChildRow
    worksheetFormatter.RefreshAssemblyChildGrouping ws, firstChildRow, lastChildRow
End Sub

Private Sub WriteTravelChildRow(ByVal ws As Worksheet, _
                                ByVal rowIndex As Long, _
                                ByVal compListNumber As Long, _
                                ByVal componentName As String, _
                                ByVal materialRate As Double)
    Dim formulaService As clsFormulaService
    Dim customData As Collection

    Set formulaService = New clsFormulaService
    Set customData = New Collection

    customData.Add 0, "des"
    customData.Add 0, "constr"
    customData.Add 0, "debug"
    customData.Add 0, "inst"
    customData.Add 0, "start"

    ws.Cells(rowIndex, COL_COMP_LIST).Value = compListNumber
    ws.Cells(rowIndex, COL_COMPONENT).Value = componentName
    ws.Cells(rowIndex, COL_COMP_DES).Value = 1
    ws.Cells(rowIndex, COL_COMP_BUILD).Value = 1
    ws.Cells(rowIndex, COL_MATERIAL_RATE).Value = materialRate

    formulaService.WriteCustomComponentFormulas ws, rowIndex, customData
End Sub

Private Sub WriteLaborChildRow(ByVal ws As Worksheet, _
                               ByVal rowIndex As Long, _
                               ByVal compListNumber As Long, _
                               ByVal componentName As String, _
                               ByVal startUpHours As Double, _
                               ByVal materialRate As Double)
    Dim formulaService As clsFormulaService
    Dim customData As Collection

    Set formulaService = New clsFormulaService
    Set customData = New Collection

    customData.Add 0, "des"
    customData.Add 0, "constr"
    customData.Add 0, "debug"
    customData.Add 0, "inst"
    customData.Add startUpHours, "start"

    ws.Cells(rowIndex, COL_COMP_LIST).Value = compListNumber
    ws.Cells(rowIndex, COL_COMPONENT).Value = componentName
    ws.Cells(rowIndex, COL_COMP_DES).Value = 1
    ws.Cells(rowIndex, COL_COMP_BUILD).Value = 1
    ws.Cells(rowIndex, COL_MATERIAL_RATE).Value = materialRate

    formulaService.WriteCustomComponentFormulas ws, rowIndex, customData
End Sub
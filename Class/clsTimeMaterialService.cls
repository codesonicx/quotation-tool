Option Explicit

' =========================================================
' clsTimeMaterialService
'
' Purpose:
' Inserts the T&M section directly.
'
' Drivers is the source of truth for live Excel formulas.
' =========================================================

Private Const SECTION_BASE_NAME As String = "T&M"

Private Const ASSEMBLY_TRAVEL As String = "Travel and Expense"
Private Const ASSEMBLY_LABOR As String = "Labor"

Private Const ITEM_PER_DIEM As String = "Per Diem"
Private Const ITEM_HOTEL As String = "Hotel"
Private Const ITEM_CAR_RENT As String = "Car Rent"
Private Const ITEM_FUEL As String = "Fuel"
Private Const ITEM_MILEAGE As String = "Mileage"
Private Const ITEM_AIR As String = "Air"

Private Const ITEM_STANDARD_HOURS As String = "Standard Hours"
Private Const ITEM_OVERTIME_HOURS As String = "Overtime Hours"

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
    InsertTravelAssembly wsTarget, insertRow

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)
    InsertLaborAssembly wsTarget, insertRow

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertTimeMaterialSection = False
        Exit Function
    End If

    InsertTimeMaterialSection = True
End Function

Private Sub InsertTravelAssembly(ByVal wsTarget As Worksheet, _
                                 ByVal insertRow As Long)
    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    assemblyRow = insertRow
    firstChildRow = assemblyRow + 1
    lastChildRow = assemblyRow + 6

    InsertAssemblyBlock wsTarget, assemblyRow, lastChildRow

    WriteAssemblyHeader wsTarget, assemblyRow, ASSEMBLY_TRAVEL

    WriteTravelChild wsTarget, firstChildRow, ITEM_PER_DIEM, "=Drivers!B24", "Drivers!B15*Drivers!B13"
    WriteTravelChild wsTarget, firstChildRow + 1, ITEM_HOTEL, "=Drivers!B19", "Drivers!B15*Drivers!B13"
    WriteTravelChild wsTarget, firstChildRow + 2, ITEM_CAR_RENT, "=Drivers!B26", "Drivers!B25*Drivers!B14"
    WriteTravelChild wsTarget, firstChildRow + 3, ITEM_FUEL, "=Drivers!B27", "Drivers!B14"
    WriteTravelChild wsTarget, firstChildRow + 4, ITEM_MILEAGE, "=Drivers!B29", "Drivers!B28*Drivers!B13*Drivers!B15"
    WriteTravelChild wsTarget, firstChildRow + 5, ITEM_AIR, "=Drivers!B23", "Drivers!B22*2"

    FinalizeAssembly wsTarget, assemblyRow, firstChildRow, lastChildRow
End Sub

Private Sub InsertLaborAssembly(ByVal wsTarget As Worksheet, _
                                ByVal insertRow As Long)
    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    assemblyRow = insertRow
    firstChildRow = assemblyRow + 1
    lastChildRow = assemblyRow + 2

    InsertAssemblyBlock wsTarget, assemblyRow, lastChildRow

    WriteAssemblyHeader wsTarget, assemblyRow, ASSEMBLY_LABOR

    WriteLaborChild wsTarget, firstChildRow, ITEM_STANDARD_HOURS, "=Drivers!B16*5*Drivers!B14*Drivers!B13",  "Drivers!B20"

    WriteLaborChild wsTarget, firstChildRow + 1, ITEM_OVERTIME_HOURS, "=(Drivers!B17*5+Drivers!B18)*Drivers!B14*Drivers!B13",  "Drivers!B21"

    FinalizeAssembly wsTarget, assemblyRow, firstChildRow, lastChildRow
End Sub

Private Sub InsertAssemblyBlock(ByVal wsTarget As Worksheet, _
                                ByVal assemblyRow As Long, _
                                ByVal lastChildRow As Long)
    Dim formatter As clsWorksheetFormatter
    Dim rowIndex As Long

    Set formatter = New clsWorksheetFormatter

    wsTarget.Rows(assemblyRow & ":" & lastChildRow).Insert Shift:=xlShiftDown

    formatter.ResetInsertedRows wsTarget.Rows(assemblyRow & ":" & lastChildRow)
    formatter.FormatAssemblyHeaderRow wsTarget, assemblyRow

    For rowIndex = assemblyRow + 1 To lastChildRow
        formatter.FormatCustomItemRow wsTarget.Rows(rowIndex)
    Next rowIndex
End Sub

Private Sub WriteAssemblyHeader(ByVal wsTarget As Worksheet, _
                                ByVal assemblyRow As Long, _
                                ByVal assemblyName As String)
    wsTarget.Cells(assemblyRow, COL_ITEM).value = assemblyName
    wsTarget.Cells(assemblyRow, COL_ITEM_DES).value = 1
    wsTarget.Cells(assemblyRow, COL_ITEM_BUILD).value = 1
End Sub

Private Sub WriteTravelChild(ByVal wsTarget As Worksheet, _
                             ByVal rowIndex As Long, _
                             ByVal componentName As String, _
                             ByVal materialRateFormula As String, _
                             ByVal materialCostMultiplierFormula As String)
    Dim formulaService As clsFormulaService

    Set formulaService = New clsFormulaService

    wsTarget.Cells(rowIndex, COL_COMPONENT).value = componentName

    wsTarget.Cells(rowIndex, COL_COMP_DES).value = 1
    wsTarget.Cells(rowIndex, COL_COMP_BUILD).value = 1

    wsTarget.Cells(rowIndex, COL_DESIGN_HOURS).value = 0
    wsTarget.Cells(rowIndex, COL_CONSTRUCTION_HOURS).value = 0
    wsTarget.Cells(rowIndex, COL_DEBUG_HOURS).value = 0
    wsTarget.Cells(rowIndex, COL_INSTALLATION_HOURS).value = 0
    wsTarget.Cells(rowIndex, COL_STARTUP_HOURS).value = 0

    wsTarget.Cells(rowIndex, COL_MATERIAL_RATE).Formula = materialRateFormula

    formulaService.WriteTimeMaterialComponentFormulas wsTarget, rowIndex

    ' Override T&M material cost calculation.
    wsTarget.Cells(rowIndex, COL_BUDGET_MATERIAL).Formula = "=" & _
        wsTarget.Cells(rowIndex, COL_MATERIAL_RATE).Address(False, False) & _
        "*" & materialCostMultiplierFormula

End Sub

Private Sub WriteLaborChild(ByVal wsTarget As Worksheet, _
                            ByVal rowIndex As Long, _
                            ByVal componentName As String, _
                            ByVal startupHoursFormula As String, _
                            ByVal laborRateFormula As String)
    Dim formulaService As clsFormulaService

    Set formulaService = New clsFormulaService

    wsTarget.Cells(rowIndex, COL_COMPONENT).value = componentName

    wsTarget.Cells(rowIndex, COL_COMP_DES).value = 1
    wsTarget.Cells(rowIndex, COL_COMP_BUILD).value = 1

    wsTarget.Cells(rowIndex, COL_DESIGN_HOURS).value = 0
    wsTarget.Cells(rowIndex, COL_CONSTRUCTION_HOURS).value = 0
    wsTarget.Cells(rowIndex, COL_DEBUG_HOURS).value = 0
    wsTarget.Cells(rowIndex, COL_INSTALLATION_HOURS).value = 0

    wsTarget.Cells(rowIndex, COL_STARTUP_HOURS).Formula = startupHoursFormula
    wsTarget.Cells(rowIndex, COL_MATERIAL_RATE).value = 0

    formulaService.WriteTimeMaterialComponentFormulas wsTarget, rowIndex

    ' Override T&M labor cost calculation in column X.
    wsTarget.Cells(rowIndex, COL_BUDGET_MATERIAL).Formula = "=" & _
        wsTarget.Cells(rowIndex, COL_STARTUP_HOURS).Address(False, False) & _
        "*" & laborRateFormula

End Sub

Private Sub FinalizeAssembly(ByVal wsTarget As Worksheet, _
                             ByVal assemblyRow As Long, _
                             ByVal firstChildRow As Long, _
                             ByVal lastChildRow As Long)
    Dim formulaService As clsFormulaService
    Dim formatter As clsWorksheetFormatter

    Set formulaService = New clsFormulaService
    Set formatter = New clsWorksheetFormatter

    formulaService.WriteAssemblyTotalFormulas _
        wsTarget, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow

    formatter.RefreshAssemblyChildGrouping _
        wsTarget, _
        firstChildRow, _
        lastChildRow
End Sub



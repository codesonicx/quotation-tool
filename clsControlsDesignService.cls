Option Explicit

' =========================================================
' clsControlsDesignService
'
' Purpose:
' Inserts the fixed Controls Design Hours assembly
' and its child hour rows into a selected section.
' =========================================================

Private Const ASSEMBLY_NAME As String = "CONTROLS DESIGN HOURS"

Private Const ITEM_HARDWARE As String = "Hardware Hours"
Private Const ITEM_PLC As String = "PLC Hours"
Private Const ITEM_HMI As String = "HMI Hours"
Private Const ITEM_SCADA As String = "Scada Hours"

Public Function InsertControlsDesignHoursIntoSection(ByVal wsTarget As Worksheet, _
                                                     ByVal sectionName As String, _
                                                     ByVal hardwareHours As Double, _
                                                     ByVal plcHours As Double, _
                                                     ByVal hmiHours As Double, _
                                                     ByVal scadaHours As Double, _
                                                     ByRef errorMessage As String) As Boolean
    Dim sectionFinder As clsSectionFinder
    Dim sectionService As clsSectionService
    Dim worksheetFormatter As clsWorksheetFormatter
    Dim formulaService As clsFormulaService

    Dim sectionRow As Long
    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString

    Set sectionFinder = New clsSectionFinder
    Set sectionService = New clsSectionService
    Set worksheetFormatter = New clsWorksheetFormatter
    Set formulaService = New clsFormulaService

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        InsertControlsDesignHoursIntoSection = False
        Exit Function
    End If

    assemblyRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    wsTarget.Rows(assemblyRow & ":" & assemblyRow + 4).Insert Shift:=xlShiftDown

    worksheetFormatter.FormatAssemblyHeaderRow wsTarget, assemblyRow

    wsTarget.Cells(assemblyRow, COL_ITEM).Value = ASSEMBLY_NAME
    wsTarget.Cells(assemblyRow, COL_ITEM_DES).Value = 1
    wsTarget.Cells(assemblyRow, COL_ITEM_BUILD).Value = 1

    firstChildRow = assemblyRow + 1
    lastChildRow = assemblyRow + 4

    WriteControlsChildRow wsTarget, firstChildRow, 1, ITEM_HARDWARE, hardwareHours
    WriteControlsChildRow wsTarget, firstChildRow + 1, 2, ITEM_PLC, plcHours
    WriteControlsChildRow wsTarget, firstChildRow + 2, 3, ITEM_HMI, hmiHours
    WriteControlsChildRow wsTarget, firstChildRow + 3, 4, ITEM_SCADA, scadaHours

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
        InsertControlsDesignHoursIntoSection = False
        Exit Function
    End If

    InsertControlsDesignHoursIntoSection = True
End Function

Private Sub WriteControlsChildRow(ByVal ws As Worksheet, _
                                  ByVal rowIndex As Long, _
                                  ByVal compListNumber As Long, _
                                  ByVal componentName As String, _
                                  ByVal designHours As Double)
    Dim customData As Collection
    Dim formulaService As clsFormulaService

    Set customData = New Collection
    Set formulaService = New clsFormulaService

    customData.Add designHours, "des"
    customData.Add 0, "constr"
    customData.Add 0, "debug"
    customData.Add 0, "inst"
    customData.Add 0, "start"
    customData.Add 0, "mat"

    ws.Cells(rowIndex, COL_COMP_LIST).Value = compListNumber
    ws.Cells(rowIndex, COL_COMPONENT).Value = componentName
    ws.Cells(rowIndex, COL_COMP_DES).Value = 1
    ws.Cells(rowIndex, COL_COMP_BUILD).Value = 1
    ws.Cells(rowIndex, COL_MATERIAL_RATE).Value = 0

    formulaService.WriteCustomComponentFormulas ws, rowIndex, customData
End Sub
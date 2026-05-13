Option Explicit

' =========================================================
' clsCustomItemService
'
' Purpose:
' Inserts custom/manual assemblies and components.
' =========================================================

Public Function InsertCustomAssemblyIntoSection(ByVal wsTarget As Worksheet, _
                                                ByVal sectionName As String, _
                                                ByVal customData As Collection, _
                                                ByRef errorMessage As String) As Boolean
    Dim sectionFinder As clsSectionFinder
    Dim sectionService As clsSectionService
    Dim worksheetFormatter As clsWorksheetFormatter
    Dim sectionRow As Long
    Dim insertRow As Long

    errorMessage = vbNullString

    Set sectionFinder = New clsSectionFinder
    Set sectionService = New clsSectionService
    Set worksheetFormatter = New clsWorksheetFormatter

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        InsertCustomAssemblyIntoSection = False
        Exit Function
    End If

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    wsTarget.Rows(insertRow).Insert Shift:=xlShiftDown
    worksheetFormatter.FormatAssemblyHeaderRow wsTarget, insertRow

    If Not PopulateInsertedCustomAssemblyRow(wsTarget, insertRow, customData, errorMessage) Then
        InsertCustomAssemblyIntoSection = False
        Exit Function
    End If

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertCustomAssemblyIntoSection = False
        Exit Function
    End If

    InsertCustomAssemblyIntoSection = True
End Function

Public Function InsertCustomComponentIntoAssembly(ByVal wsTarget As Worksheet, _
                                                  ByVal sectionName As String, _
                                                  ByVal assemblyName As String, _
                                                  ByVal customData As Collection, _
                                                  ByRef errorMessage As String) As Boolean
    Dim assemblyFinder As clsAssemblyFinder
    Dim sectionService As clsSectionService
    Dim worksheetFormatter As clsWorksheetFormatter
    Dim formulaService As clsFormulaService

    Dim insertRow As Long
    Dim assemblyRow As Long
    Dim assemblyEndRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    Dim insertedRow As Range

    errorMessage = vbNullString

    Set assemblyFinder = New clsAssemblyFinder
    Set sectionService = New clsSectionService
    Set worksheetFormatter = New clsWorksheetFormatter
    Set formulaService = New clsFormulaService

    insertRow = assemblyFinder.FindCustomComponentInsertRow(wsTarget, sectionName, assemblyName)

    If insertRow = 0 Then
        errorMessage = "Could not determine insert row for assembly '" & assemblyName & "' in section '" & sectionName & "'."
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    wsTarget.Rows(insertRow).Insert Shift:=xlShiftDown
    Set insertedRow = wsTarget.Rows(insertRow)

    worksheetFormatter.FormatCustomItemRow insertedRow

    If Not PopulateInsertedCustomRow(wsTarget, insertRow, customData, errorMessage) Then
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    assemblyRow = assemblyFinder.FindAssemblyHeaderRow(wsTarget, sectionName, assemblyName)
    assemblyEndRow = assemblyFinder.FindAssemblyEndRow(wsTarget, sectionName, assemblyName)

    firstChildRow = assemblyRow + 1
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
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    InsertCustomComponentIntoAssembly = True
End Function

Private Function PopulateInsertedCustomAssemblyRow(ByVal wsTarget As Worksheet, _
                                                   ByVal targetRow As Long, _
                                                   ByVal customData As Collection, _
                                                   ByRef errorMessage As String) As Boolean
    Dim headerFinder As clsHeaderFinder
    Dim itemCol As Long
    Dim itemDesCol As Long
    Dim itemBuildCol As Long

    errorMessage = vbNullString
    Set headerFinder = New clsHeaderFinder

    itemCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_ITEM)
    itemDesCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_ITEM_DES)
    itemBuildCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_ITEM_BUILD)

    If itemCol = 0 Then
        errorMessage = "Column '" & HEADER_ITEM & "' was not found in the target sheet."
        PopulateInsertedCustomAssemblyRow = False
        Exit Function
    End If

    wsTarget.Cells(targetRow, itemCol).Value = CStr(customData("name"))

    If itemDesCol > 0 Then wsTarget.Cells(targetRow, itemDesCol).Value = CLng(customData("desNum"))
    If itemBuildCol > 0 Then wsTarget.Cells(targetRow, itemBuildCol).Value = CLng(customData("buildNum"))

    PopulateInsertedCustomAssemblyRow = True
End Function

Private Function PopulateInsertedCustomRow(ByVal wsTarget As Worksheet, _
                                           ByVal targetRow As Long, _
                                           ByVal customData As Collection, _
                                           ByRef errorMessage As String) As Boolean

    Dim formulaService As clsFormulaService

    errorMessage = vbNullString

    Set formulaService = New clsFormulaService

    wsTarget.Cells(targetRow, COL_COMPONENT).Value = _
        "[" & CStr(customData("domain")) & "] " & CStr(customData("name"))

    wsTarget.Cells(targetRow, COL_COMP_DES).Value = _
        CLng(customData("desNum"))

    wsTarget.Cells(targetRow, COL_COMP_BUILD).Value = _
        CLng(customData("buildNum"))

    wsTarget.Cells(targetRow, COL_MATERIAL_RATE).Value = _
        CDbl(customData("mat"))

    formulaService.WriteCustomComponentFormulas _
        wsTarget, _
        targetRow, _
        customData

    PopulateInsertedCustomRow = True

End Function
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
    Dim insertedRow As Range

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
    Set insertedRow = wsTarget.Rows(insertRow)

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
    Dim wsLib As Worksheet

    Dim assemblyFinder As clsAssemblyFinder
    Dim libraryFinder As clsLibraryFinder
    Dim sectionService As clsSectionService
    Dim worksheetFormatter As clsWorksheetFormatter

    Dim insertRow As Long
    Dim templateRow As Long
    Dim lastCol As Long

    Dim src As Range
    Dim dst As Range
    Dim insertedRow As Range

    errorMessage = vbNullString

    Set assemblyFinder = New clsAssemblyFinder
    Set libraryFinder = New clsLibraryFinder
    Set sectionService = New clsSectionService
    Set worksheetFormatter = New clsWorksheetFormatter

    Set wsLib = ThisWorkbook.Worksheets(SHEET_ASSEMBLY_LIST)

    insertRow = assemblyFinder.FindCustomComponentInsertRow( _
        wsTarget, _
        sectionName, _
        assemblyName)

    If insertRow = 0 Then
        errorMessage = "Could not determine insert row for assembly '" & assemblyName & "' in section '" & sectionName & "'."
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    templateRow = libraryFinder.FindManualEntryTemplateRow(wsLib, errorMessage)

    If templateRow = 0 Then
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    ' Determine the copy range on the template row from column B to the last used column.
    lastCol = wsLib.Cells(templateRow, wsLib.Columns.Count).End(xlToLeft).Column

    If lastCol < COL_COPY_START Then
        lastCol = COL_COPY_START
    End If

    Set src = wsLib.Range( _
        wsLib.Cells(templateRow, COL_COPY_START), _
        wsLib.Cells(templateRow, lastCol))

    ' Insert one row in the target sheet and reset inherited formatting/grouping.
    wsTarget.Rows(insertRow).Insert Shift:=xlShiftDown
    Set insertedRow = wsTarget.Rows(insertRow)

    ' Copy the template row into the inserted target row.
    Set dst = wsTarget.Range( _
        wsTarget.Cells(insertRow, COL_COPY_START), _
        wsTarget.Cells(insertRow, lastCol))

    src.Copy Destination:=dst
    Application.CutCopyMode = False

    worksheetFormatter.FormatCustomItemRow insertedRow

    If Not PopulateInsertedCustomRow(wsTarget, insertRow, customData, errorMessage) Then
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

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

    If itemDesCol > 0 Then
        wsTarget.Cells(targetRow, itemDesCol).Value = CLng(customData("desNum"))
    End If

    If itemBuildCol > 0 Then
        wsTarget.Cells(targetRow, itemBuildCol).Value = CLng(customData("buildNum"))
    End If

    PopulateInsertedCustomAssemblyRow = True
End Function

Private Function PopulateInsertedCustomRow(ByVal wsTarget As Worksheet, _
                                           ByVal targetRow As Long, _
                                           ByVal customData As Collection, _
                                           ByRef errorMessage As String) As Boolean
    Dim formulaService As clsFormulaService
    Dim headerFinder As clsHeaderFinder

    Dim compCol As Long
    Dim compDesCol As Long
    Dim compBuildCol As Long
    Dim sectDesCol As Long
    Dim sectBuildCol As Long
    Dim designCol As Long
    Dim constrCol As Long
    Dim debugCol As Long
    Dim installCol As Long
    Dim startCol As Long
    Dim matCol As Long
    Dim compListCol As Long

    errorMessage = vbNullString

    Set formulaService = New clsFormulaService
    Set headerFinder = New clsHeaderFinder

    compCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_COMPONENT)
    compDesCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_COMP_DES)
    compBuildCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_COMP_BUILD)
    sectDesCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_SECT_DES)
    sectBuildCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_SECT_BUILD)
    designCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_DESIGN_HOURS)
    constrCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_CONSTRUCTION_HOURS)
    debugCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_DEBUG_HOURS)
    installCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_INSTALLATION_HOURS)
    startCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_STARTUP_HOURS)
    matCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_MATERIAL_RATE)
    compListCol = headerFinder.FindHeaderColumn(wsTarget, HEADER_COMP_LIST)

    If compCol = 0 Then
        errorMessage = "Column '" & HEADER_COMPONENT & "' was not found in the target sheet."
        PopulateInsertedCustomRow = False
        Exit Function
    End If

    wsTarget.Cells(targetRow, compCol).Value = "[" & DOMAIN_ME & "] " & CStr(customData("name"))

    If compDesCol > 0 Then
        wsTarget.Cells(targetRow, compDesCol).Value = CLng(customData("desNum"))
    End If

    If compBuildCol > 0 Then
        wsTarget.Cells(targetRow, compBuildCol).Value = CLng(customData("buildNum"))
    End If

    If compListCol > 0 Then
        formulaService.WriteComponentListNumber wsTarget, targetRow, compListCol
    End If

    formulaService.WriteDesignFormula wsTarget, targetRow, designCol, sectDesCol, CDbl(customData("des"))
    formulaService.WriteBuildFormula wsTarget, targetRow, constrCol, sectBuildCol, CDbl(customData("constr"))
    formulaService.WriteBuildFormula wsTarget, targetRow, debugCol, sectBuildCol, CDbl(customData("debug"))
    formulaService.WriteBuildFormula wsTarget, targetRow, installCol, sectBuildCol, CDbl(customData("inst"))
    formulaService.WriteBuildFormula wsTarget, targetRow, startCol, sectBuildCol, CDbl(customData("start"))

    If matCol > 0 Then
        wsTarget.Cells(targetRow, matCol).Value = CLng(customData("mat"))
    End If

    PopulateInsertedCustomRow = True
End Function
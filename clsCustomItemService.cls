Option Explicit

' =========================================================
' clsCustomItemsService
'
' Purpose:
' Manages insertion of custom/manual components into an
' existing assembly inside the target worksheet.
' =========================================================

Public Function InsertCustomComponentIntoAssembly(ByVal wsTarget As Worksheet, _
                                                  ByVal sectionName As String, _
                                                  ByVal assemblyName As String, _
                                                  ByVal customData As Collection, _
                                                  ByRef errorMessage As String) As Boolean
    Dim wsLib As Worksheet
    Dim assemblyFinder As clsAssemblyFinder
    Dim sectionService As clsSectionService

    Dim insertRow As Long
    Dim templateRow As Long
    Dim lastCol As Long

    Dim src As Range
    Dim dst As Range
    Dim insertedRow As Range

    errorMessage = vbNullString

    Set assemblyFinder = New clsAssemblyFinder
    Set sectionService = New clsSectionService
    Set wsLib = ThisWorkbook.Worksheets(SHEET_ASSEMBLY_LIST)

    ' Find where the custom component should be inserted in the target sheet.
    insertRow = assemblyFinder.FindCustomComponentInsertRow(wsTarget, sectionName, assemblyName)
    If insertRow = 0 Then
        errorMessage = "Could not determine insert row for assembly '" & assemblyName & "' in section '" & sectionName & "'."
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    ' Find the template row used as the base for manual-entry components.
    templateRow = FindExactTextInColumn(wsLib, COL_LIBRARY_NAME, TEMPLATE_ME_MANUAL_ENTRY)
    If templateRow = 0 Then
        errorMessage = "Template row '" & TEMPLATE_ME_MANUAL_ENTRY & "' was not found in '" & SHEET_ASSEMBLY_LIST & "'."
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

    insertedRow.ClearFormats
    insertedRow.ClearOutline

    ' Copy the template row into the inserted target row.
    Set dst = wsTarget.Range( _
        wsTarget.Cells(insertRow, COL_COPY_START), _
        wsTarget.Cells(insertRow, lastCol))

    src.Copy Destination:=dst
    Application.CutCopyMode = False
    insertedRow.Interior.Pattern = xlNone

    ' Fill values and formulas into the inserted row.
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

' =========================================================
' Row population
' =========================================================

' Writes custom values and formulas into the inserted row.
Private Function PopulateInsertedCustomRow(ByVal wsTarget As Worksheet, _
                                           ByVal targetRow As Long, _
                                           ByVal customData As Collection, _
                                           ByRef errorMessage As String) As Boolean
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

    compCol = FindHeaderColumn(wsTarget, HEADER_COMPONENT)
    compDesCol = FindHeaderColumn(wsTarget, HEADER_COMP_DES)
    compBuildCol = FindHeaderColumn(wsTarget, HEADER_COMP_BUILD)
    sectDesCol = FindHeaderColumn(wsTarget, HEADER_SECT_DES)
    sectBuildCol = FindHeaderColumn(wsTarget, HEADER_SECT_BUILD)
    designCol = FindHeaderColumn(wsTarget, HEADER_DESIGN_HOURS)
    constrCol = FindHeaderColumn(wsTarget, HEADER_CONSTRUCTION_HOURS)
    debugCol = FindHeaderColumn(wsTarget, HEADER_DEBUG_HOURS)
    installCol = FindHeaderColumn(wsTarget, HEADER_INSTALLATION_HOURS)
    startCol = FindHeaderColumn(wsTarget, HEADER_STARTUP_HOURS)
    matCol = FindHeaderColumn(wsTarget, HEADER_MATERIAL_RATE)
    compListCol = FindHeaderColumn(wsTarget, HEADER_COMP_LIST)

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
        WriteComponentListNumber wsTarget, targetRow, compListCol
    End If

    WriteDesignFormula wsTarget, targetRow, designCol, sectDesCol, CDbl(customData("des"))
    WriteBuildFormula wsTarget, targetRow, constrCol, sectBuildCol, CDbl(customData("constr"))
    WriteBuildFormula wsTarget, targetRow, debugCol, sectBuildCol, CDbl(customData("debug"))
    WriteBuildFormula wsTarget, targetRow, installCol, sectBuildCol, CDbl(customData("inst"))
    WriteBuildFormula wsTarget, targetRow, startCol, sectBuildCol, CDbl(customData("start"))

    If matCol > 0 Then
        wsTarget.Cells(targetRow, matCol).Value = CLng(customData("mat"))
    End If

    PopulateInsertedCustomRow = True
End Function

' =========================================================
' Formula helpers
' =========================================================

' Writes the design-hours formula using Sect Des. as the multiplier source.
Private Sub WriteDesignFormula(ByVal wsTarget As Worksheet, _
                               ByVal targetRow As Long, _
                               ByVal targetCol As Long, _
                               ByVal sourceCol As Long, _
                               ByVal multiplier As Double)
    If targetCol > 0 And sourceCol > 0 Then
        wsTarget.Cells(targetRow, targetCol).Formula = "=" & _
            wsTarget.Cells(targetRow, sourceCol).Address(False, False) & "*" & multiplier
    End If
End Sub

' Writes a formula using Sect Build as the multiplier source.
Private Sub WriteBuildFormula(ByVal wsTarget As Worksheet, _
                              ByVal targetRow As Long, _
                              ByVal targetCol As Long, _
                              ByVal sourceCol As Long, _
                              ByVal multiplier As Double)
    If targetCol > 0 And sourceCol > 0 Then
        wsTarget.Cells(targetRow, targetCol).Formula = "=" & _
            wsTarget.Cells(targetRow, sourceCol).Address(False, False) & "*" & multiplier
    End If
End Sub

' Writes the next "Comp List" number based on the previous row when available.
Private Sub WriteComponentListNumber(ByVal wsTarget As Worksheet, _
                                     ByVal targetRow As Long, _
                                     ByVal compListCol As Long)
    If targetRow > 1 And IsNumeric(wsTarget.Cells(targetRow - 1, compListCol).Value) Then
        wsTarget.Cells(targetRow, compListCol).Value = CLng(wsTarget.Cells(targetRow - 1, compListCol).Value) + 1
    Else
        wsTarget.Cells(targetRow, compListCol).Value = 1
    End If
End Sub

' =========================================================
' Search helpers
' =========================================================

' Finds an exact text match in the requested worksheet column.
Private Function FindExactTextInColumn(ByVal ws As Worksheet, _
                                       ByVal colIndex As Long, _
                                       ByVal findText As String) As Long
    Dim foundCell As Range

    Set foundCell = ws.Columns(colIndex).Find( _
        What:=findText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        FindExactTextInColumn = 0
    Else
        FindExactTextInColumn = foundCell.Row
    End If
End Function

' Finds a header column by searching row 1 for an exact header label.
Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim foundCell As Range

    Set foundCell = ws.Rows(1).Find( _
        What:=headerText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        FindHeaderColumn = 0
    Else
        FindHeaderColumn = foundCell.Column
    End If
End Function
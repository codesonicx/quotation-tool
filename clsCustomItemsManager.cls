Option Explicit

' =========================================================
' clsCustomItemsManager
'
' Purpose:
' Manages insertion of custom/manual components into an
' existing assembly inside the target worksheet.
'
' Responsibilities:
' - find the correct insert row inside an assembly
' - copy the manual-entry template row from "Assembly List"
' - write custom values into the inserted row
' - write formulas into calculated columns when available
' =========================================================

' =========================================================
' Template and worksheet constants
' =========================================================
Private Const TEMPLATE_SHEET_NAME As String = "Assembly List"
Private Const TEMPLATE_NAME_COLUMN As Long = 2                 ' Column B
Private Const TEMPLATE_LABEL_MANUAL_ENTRY As String = "ME - Manual Entry"
Private Const TARGET_COPY_START_COLUMN As Long = 2             ' Column B

' =========================================================
' Public API
' =========================================================

Public Function InsertCustomComponentIntoAssembly(ByVal wsTarget As Worksheet, _
                                                  ByVal sectionName As String, _
                                                  ByVal assemblyName As String, _
                                                  ByVal customData As Collection, _
                                                  ByRef errorMessage As String) As Boolean
    Dim wsLib As Worksheet
    Dim navigator As clsSheetStructureNavigator

    Dim insertRow As Long
    Dim templateRow As Long
    Dim lastCol As Long

    Dim src As Range
    Dim dst As Range
    Dim insertedRow As Range

    errorMessage = vbNullString
    Set navigator = New clsSheetStructureNavigator
    Set wsLib = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME)

    ' 1) Find where the custom component should be inserted in the target sheet.
    insertRow = navigator.FindInsertRowForCustomComponent(wsTarget, sectionName, assemblyName)
    If insertRow = 0 Then
        errorMessage = "Could not determine insert row for assembly '" & assemblyName & "' in section '" & sectionName & "'."
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    ' 2) Find the template row used as the base for manual-entry components.
    templateRow = FindExactTextInColumn(wsLib, TEMPLATE_NAME_COLUMN, TEMPLATE_LABEL_MANUAL_ENTRY)
    If templateRow = 0 Then
        errorMessage = "Template row '" & TEMPLATE_LABEL_MANUAL_ENTRY & "' was not found in '" & TEMPLATE_SHEET_NAME & "'."
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    ' 3) Determine the copy range on the template row from column B to the last used column.
    lastCol = wsLib.Cells(templateRow, wsLib.Columns.count).End(xlToLeft).Column
    If lastCol < TARGET_COPY_START_COLUMN Then
        lastCol = TARGET_COPY_START_COLUMN
    End If

    Set src = wsLib.Range( _
        wsLib.Cells(templateRow, TARGET_COPY_START_COLUMN), _
        wsLib.Cells(templateRow, lastCol))

    ' 4) Insert one row in the target sheet and reset inherited formatting/grouping.
    wsTarget.Rows(insertRow).Insert Shift:=xlShiftDown
    Set insertedRow = wsTarget.Rows(insertRow)

    insertedRow.ClearFormats
    insertedRow.ClearOutline

    ' 5) Copy the template row into the inserted target row.
    Set dst = wsTarget.Range( _
        wsTarget.Cells(insertRow, TARGET_COPY_START_COLUMN), _
        wsTarget.Cells(insertRow, lastCol))

    src.Copy Destination:=dst
    Application.CutCopyMode = False
    insertedRow.Interior.Pattern = xlNone

    ' 6) Fill values and formulas into the inserted row.
    If Not PopulateInsertedCustomRow(wsTarget, insertRow, customData, errorMessage) Then
        InsertCustomComponentIntoAssembly = False
        Exit Function
    End If

    Dim sectionManager As clsSectionManager
    Set sectionManager = New clsSectionManager
    
    If Not sectionManager.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
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

    ' Locate important target columns by header name.
    compCol = FindHeaderColumn(wsTarget, "Component")
    compDesCol = FindHeaderColumn(wsTarget, "Comp Des.")
    compBuildCol = FindHeaderColumn(wsTarget, "Comp Build")
    sectDesCol = FindHeaderColumn(wsTarget, "Sect Des.")
    sectBuildCol = FindHeaderColumn(wsTarget, "Sect Build")
    designCol = FindHeaderColumn(wsTarget, "design [h]")
    constrCol = FindHeaderColumn(wsTarget, "construction [h]")
    debugCol = FindHeaderColumn(wsTarget, "debug/tryout [h]")
    installCol = FindHeaderColumn(wsTarget, "installation [h]")
    startCol = FindHeaderColumn(wsTarget, "start up @ client [h]")
    matCol = FindHeaderColumn(wsTarget, "Material Rate [$/Build]")
    compListCol = FindHeaderColumn(wsTarget, "Comp List")

    ' The component description column is required.
    If compCol = 0 Then
        errorMessage = "Column 'Component' was not found in the target sheet."
        PopulateInsertedCustomRow = False
        Exit Function
    End If

    ' Required base values.
    wsTarget.Cells(targetRow, compCol).value = "[ME] " & CStr(customData("name"))

    If compDesCol > 0 Then
        wsTarget.Cells(targetRow, compDesCol).value = CLng(customData("desNum"))
    End If

    If compBuildCol > 0 Then
        wsTarget.Cells(targetRow, compBuildCol).value = CLng(customData("buildNum"))
    End If

    ' Optional numbering in "Comp List".
    If compListCol > 0 Then
        WriteComponentListNumber wsTarget, targetRow, compListCol
    End If

    ' Formula-driven values.
    WriteDesignFormula wsTarget, targetRow, designCol, sectDesCol, CDbl(customData("des"))
    WriteBuildFormula wsTarget, targetRow, constrCol, sectBuildCol, CDbl(customData("constr"))
    WriteBuildFormula wsTarget, targetRow, debugCol, sectBuildCol, CDbl(customData("debug"))
    WriteBuildFormula wsTarget, targetRow, installCol, sectBuildCol, CDbl(customData("inst"))
    WriteBuildFormula wsTarget, targetRow, startCol, sectBuildCol, CDbl(customData("start"))

    If matCol > 0 Then
        wsTarget.Cells(targetRow, matCol).value = CLng(customData("mat"))
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
    If targetRow > 1 And IsNumeric(wsTarget.Cells(targetRow - 1, compListCol).value) Then
        wsTarget.Cells(targetRow, compListCol).value = CLng(wsTarget.Cells(targetRow - 1, compListCol).value) + 1
    Else
        wsTarget.Cells(targetRow, compListCol).value = 1
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
        FindExactTextInColumn = foundCell.row
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


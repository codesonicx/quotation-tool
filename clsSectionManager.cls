Option Explicit

' =========================================================
' clsSectionManager
'
' Purpose:
' Manages changes to section data in the worksheet.
'
' Current responsibilities:
' - add section
' - remove section
' - recalculate section totals
'
' Future candidates:
' - rename section
' =========================================================

' Returns the next row after the last used row.
' Returns row 1 when the sheet is empty.
Private Function GetNextRowAfterLastUsed(ByVal ws As Worksheet) As Long
    Dim lastRow As Long

    lastRow = modUtils.GetLastUsedRow(ws)

    If lastRow = 0 Then
        GetNextRowAfterLastUsed = 1
    Else
        GetNextRowAfterLastUsed = lastRow + 1
    End If
End Function

' =========================================================
' Public API
' =========================================================

' Adds a new section at the bottom of the used worksheet area.
'
' A section row is represented by:
' - section name in column A
' - full row colored with the section header color
Public Function AddSection(ByVal ws As Worksheet, _
                           ByVal sectionName As String, _
                           ByRef errorMessage As String) As Boolean
    Dim newRow As Long
    Dim existing As Range

    errorMessage = vbNullString
    sectionName = modUtils.NormalizeText(sectionName)

    If sectionName = "" Then
        errorMessage = "Section name cannot be empty."
        AddSection = False
        Exit Function
    End If

    Set existing = ws.Columns(COL_SECTION_NAME).Find(What:=sectionName, _
                                                     LookIn:=xlValues, _
                                                     LookAt:=xlWhole, _
                                                     MatchCase:=False)

    If Not existing Is Nothing Then
        errorMessage = "A section with this name already exists."
        AddSection = False
        Exit Function
    End If

    newRow = GetNextRowAfterLastUsed(ws)

    ' Safe guard:
    ' if the target row already contains data, insert a new row first.
    If Application.WorksheetFunction.CountA(ws.Rows(newRow)) > 0 Then
        ws.Rows(newRow).Insert Shift:=xlShiftDown
    End If

    ws.Cells(newRow, COL_SECTION_NAME).Value2 = sectionName
    ws.Rows(newRow).Interior.ColorIndex = COLOR_INDEX_SECTION_BLUE

    AddSection = True
End Function

' Removes the specified section and all rows that belong to it.
'
' The deleted range starts at the section header row and continues
' through the row immediately before the next section header.

Public Function RemoveSection(ByVal ws As Worksheet, _
                              ByVal sectionName As String, _
                              ByRef errorMessage As String) As Boolean
    Dim navigator As clsSheetStructureNavigator
    Dim sectionRow As Long
    Dim nextSectionRow As Long

    errorMessage = vbNullString
    sectionName = modUtils.NormalizeText(sectionName)

    If sectionName = "" Then
        errorMessage = "Section name cannot be empty."
        RemoveSection = False
        Exit Function
    End If

    Set navigator = New clsSheetStructureNavigator

    sectionRow = navigator.FindSectionRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        RemoveSection = False
        Exit Function
    End If

    nextSectionRow = navigator.FindNextSectionHeaderRow(ws, sectionRow)

    ' Delete from the section header row
    ' down to the row before the next section
    ws.Rows(sectionRow & ":" & (nextSectionRow - 1)).Delete Shift:=xlUp

    RemoveSection = True
End Function

' Rebuilds the section summary formulas from column M to AF
' by summing all assembly header rows inside the section.
Public Function RecalculateSectionTotals(ByVal ws As Worksheet, _
                                         ByVal sectionName As String, _
                                         ByRef errorMessage As String) As Boolean
    Dim navigator As clsSheetStructureNavigator
    Dim assemblyRows() As Long

    Dim sectionRow As Long
    Dim success As Boolean

    Dim colIndex As Long
    Dim i As Long
    Dim formulaText As String
    Dim columnLetter As String

    errorMessage = vbNullString
    Set navigator = New clsSheetStructureNavigator

    sectionRow = navigator.FindSectionRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        RecalculateSectionTotals = False
        Exit Function
    End If

    success = navigator.GetAssemblyHeaderRowsInSection(ws, sectionName, assemblyRows, errorMessage)
    If Not success Then
        RecalculateSectionTotals = False
        Exit Function
    End If

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

    RecalculateSectionTotals = True
End Function

' Removes the specified assembly and all child rows that belong to it.
' After deletion, the section totals are recalculated.
Public Function RemoveAssemblyFromSection(ByVal ws As Worksheet, _
                                          ByVal sectionName As String, _
                                          ByVal assemblyName As String, _
                                          ByRef errorMessage As String) As Boolean
    Dim navigator As clsSheetStructureNavigator
    Dim assemblyRow As Long
    Dim blockEndRow As Long

    errorMessage = vbNullString
    sectionName = modUtils.NormalizeText(sectionName)
    assemblyName = modUtils.NormalizeText(assemblyName)

    If sectionName = "" Then
        errorMessage = "Section name cannot be empty."
        RemoveAssemblyFromSection = False
        Exit Function
    End If

    If assemblyName = "" Then
        errorMessage = "Assembly name cannot be empty."
        RemoveAssemblyFromSection = False
        Exit Function
    End If

    Set navigator = New clsSheetStructureNavigator

    assemblyRow = navigator.FindAssemblyHeaderRowInSection(ws, sectionName, assemblyName)
    If assemblyRow = 0 Then
        errorMessage = "Assembly '" & assemblyName & "' was not found in section '" & sectionName & "'."
        RemoveAssemblyFromSection = False
        Exit Function
    End If

    blockEndRow = navigator.FindAssemblyBlockEndRowInSection(ws, sectionName, assemblyName)
    If blockEndRow = 0 Or blockEndRow < assemblyRow Then
        errorMessage = "Could not determine the full row range for assembly '" & assemblyName & "'."
        RemoveAssemblyFromSection = False
        Exit Function
    End If

    ' Delete the assembly header row and all child rows below it.
    ws.Rows(assemblyRow & ":" & blockEndRow).Delete Shift:=xlUp

    If Not RecalculateSectionTotals(ws, sectionName, errorMessage) Then
        RemoveAssemblyFromSection = False
        Exit Function
    End If

    RemoveAssemblyFromSection = True
End Function
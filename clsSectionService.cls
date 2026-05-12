Option Explicit

' =========================================================
' clsSectionService
'
' Purpose:
' Manages section-level worksheet changes.
'
' Responsibilities:
' - add sections
' - remove sections
' - refresh section totals
' - remove assemblies from sections
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

    Set existing = ws.Columns(COL_SECTION_NAME).Find( _
        What:=sectionName, _
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

Public Function RemoveSection(ByVal ws As Worksheet, _
                              ByVal sectionName As String, _
                              ByRef errorMessage As String) As Boolean
    Dim sectionFinder As clsSectionFinder
    Dim sectionRow As Long
    Dim nextSectionRow As Long

    errorMessage = vbNullString
    sectionName = modUtils.NormalizeText(sectionName)

    If sectionName = "" Then
        errorMessage = "Section name cannot be empty."
        RemoveSection = False
        Exit Function
    End If

    Set sectionFinder = New clsSectionFinder

    sectionRow = sectionFinder.FindSectionHeaderRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        RemoveSection = False
        Exit Function
    End If

    nextSectionRow = sectionFinder.FindNextSectionHeaderRowAfter(ws, sectionRow)

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
    Dim sectionFinder As clsSectionFinder
    Dim assemblyFinder As clsAssemblyFinder
    Dim formulaService As clsFormulaService

    Dim assemblyRows() As Long
    Dim sectionRow As Long

    errorMessage = vbNullString

    Set sectionFinder = New clsSectionFinder
    Set assemblyFinder = New clsAssemblyFinder
    Set formulaService = New clsFormulaService

    sectionRow = sectionFinder.FindSectionHeaderRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        RecalculateSectionTotals = False
        Exit Function
    End If

    If Not assemblyFinder.GetAssemblyHeaderRows(ws, sectionName, assemblyRows, errorMessage) Then
        RecalculateSectionTotals = False
        Exit Function
    End If

    formulaService.WriteSectionTotalFormulas ws, sectionRow, assemblyRows

    RecalculateSectionTotals = True
End Function

Public Function RemoveAssemblyFromSection(ByVal ws As Worksheet, _
                                          ByVal sectionName As String, _
                                          ByVal assemblyName As String, _
                                          ByRef errorMessage As String) As Boolean
    Dim assemblyFinder As clsAssemblyFinder
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

    Set assemblyFinder = New clsAssemblyFinder

    assemblyRow = assemblyFinder.FindAssemblyHeaderRow(ws, sectionName, assemblyName)
    If assemblyRow = 0 Then
        errorMessage = "Assembly '" & assemblyName & "' was not found in section '" & sectionName & "'."
        RemoveAssemblyFromSection = False
        Exit Function
    End If

    blockEndRow = assemblyFinder.FindAssemblyEndRow(ws, sectionName, assemblyName)
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
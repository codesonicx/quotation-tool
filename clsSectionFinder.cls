Option Explicit

' =========================================================
' clsSheetStructureNavigator
'
' Purpose:
' Provides read/query operations to interpret the worksheet
' structure, including:
' - section lookup
' - assembly lookup inside sections
' - row classification
' - insert row calculation for new items
'
' Worksheet structure rules assumed by this class:
' - Section header rows:
'     * full row colored blue
'     * section name stored in column A
' - Assembly header rows:
'     * full row colored gray
'     * assembly name stored in column B
' =========================================================

' Returns normalized text from the specified worksheet cell.
Private Function GetNormalizedCellText(ByVal ws As Worksheet, _
                                       ByVal rowNumber As Long, _
                                       ByVal columnNumber As Long) As String
    GetNormalizedCellText = modUtils.NormalizeText(CStr(ws.Cells(rowNumber, columnNumber).Value2))
End Function

' Returns the section name stored at the given row.
Private Function GetSectionNameAtRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    GetSectionNameAtRow = GetNormalizedCellText(ws, rowNumber, COL_SECTION_NAME)
End Function

' Returns the assembly name stored at the given row.
Private Function GetAssemblyNameAtRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    GetAssemblyNameAtRow = GetNormalizedCellText(ws, rowNumber, COL_ASSEMBLY_NAME)
End Function

' =========================================================
' Row classification
' =========================================================

' Returns True when the specified row is a valid section header row.
Public Function IsSectionHeaderRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    IsSectionHeaderRow = _
        (ws.Rows(rowNumber).Interior.ColorIndex = COLOR_INDEX_SECTION_BLUE) _
        And (GetSectionNameAtRow(ws, rowNumber) <> "")
End Function

' Returns True when the specified row is a valid assembly header row.
Public Function IsAssemblyHeaderRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    IsAssemblyHeaderRow = _
        (ws.Rows(rowNumber).Interior.ColorIndex = COLOR_INDEX_ASSEMBLY_GRAY) _
        And (GetAssemblyNameAtRow(ws, rowNumber) <> "")
End Function

' =========================================================
' Public API
' =========================================================

' Finds the next section header row below the given section header row.
' If there is no next section, returns the row after the last used row.
Public Function FindNextSectionHeaderRow(ByVal ws As Worksheet, ByVal sectionHeaderRow As Long) As Long
    Dim lastRow As Long
    Dim rowNumber As Long

    lastRow = modUtils.GetLastUsedRow(ws)

    If lastRow = 0 Then
        FindNextSectionHeaderRow = sectionHeaderRow + 1
        Exit Function
    End If

    For rowNumber = sectionHeaderRow + 1 To lastRow
        If IsSectionHeaderRow(ws, rowNumber) Then
            FindNextSectionHeaderRow = rowNumber
            Exit Function
        End If
    Next rowNumber

    FindNextSectionHeaderRow = lastRow + 1
End Function

' =========================================================
' Section queries
' =========================================================

' Returns all valid section names found in the worksheet.
Public Function GetSectionNames(ByVal ws As Worksheet, _
                                ByRef sections() As String, _
                                ByRef errorMessage As String) As Boolean
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim count As Long
    Dim sectionText As String

    errorMessage = vbNullString
    count = 0
    lastRow = modUtils.GetLastUsedRow(ws)

    If lastRow < FIRST_DATA_ROW Then
        errorMessage = "No sections found (blue rows) in column A."
        GetSectionNames = False
        Exit Function
    End If

    For rowNumber = FIRST_DATA_ROW To lastRow
        If IsSectionHeaderRow(ws, rowNumber) Then
            sectionText = GetSectionNameAtRow(ws, rowNumber)

            If sectionText <> "" Then
                count = count + 1
                ReDim Preserve sections(1 To count)
                sections(count) = sectionText
            End If
        End If
    Next rowNumber

    If count = 0 Then
        errorMessage = "No sections found (blue rows) in column A."
        GetSectionNames = False
        Exit Function
    End If

    GetSectionNames = True
End Function

' Finds the row of the requested section header.
' Returns 0 when the section does not exist or the matched row is not a valid section header.
Public Function FindSectionRow(ByVal ws As Worksheet, ByVal sectionName As String) As Long
    Dim matchCell As Range

    sectionName = modUtils.NormalizeText(sectionName)
    If sectionName = "" Then
        FindSectionRow = 0
        Exit Function
    End If

    Set matchCell = ws.Columns(COL_SECTION_NAME).Find(What:=sectionName, _
                                                         LookIn:=xlValues, _
                                                         LookAt:=xlWhole, _
                                                         MatchCase:=False)

    If matchCell Is Nothing Then
        FindSectionRow = 0
        Exit Function
    End If

    If IsSectionHeaderRow(ws, matchCell.row) Then
        FindSectionRow = matchCell.row
    Else
        FindSectionRow = 0
    End If
End Function

' =========================================================
' Assembly queries
' =========================================================

' Returns the assembly names found inside the requested section.
Public Function GetAssembliesInSection(ByVal ws As Worksheet, _
                                       ByVal sectionName As String, _
                                       ByRef assemblies() As String, _
                                       ByRef errorMessage As String) As Boolean
    Dim sectionRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long
    Dim count As Long
    Dim assemblyText As String

    errorMessage = vbNullString
    count = 0

    sectionRow = FindSectionRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        GetAssembliesInSection = False
        Exit Function
    End If

    nextSectionRow = FindNextSectionHeaderRow(ws, sectionRow)

    For rowNumber = sectionRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            assemblyText = GetAssemblyNameAtRow(ws, rowNumber)

            If assemblyText <> "" Then
                count = count + 1
                ReDim Preserve assemblies(1 To count)
                assemblies(count) = assemblyText
            End If
        End If
    Next rowNumber

    If count = 0 Then
        errorMessage = "No assemblies found inside section '" & sectionName & "'."
        GetAssembliesInSection = False
        Exit Function
    End If

    GetAssembliesInSection = True
End Function

' Finds the requested assembly header row, but only inside the specified section.
' Returns 0 when the section or assembly is not found.
Public Function FindAssemblyHeaderRowInSection(ByVal ws As Worksheet, _
                                               ByVal sectionName As String, _
                                               ByVal assemblyName As String) As Long
    Dim sectionRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long
    Dim currentAssemblyName As String

    sectionRow = FindSectionRow(ws, sectionName)
    If sectionRow = 0 Then
        FindAssemblyHeaderRowInSection = 0
        Exit Function
    End If

    nextSectionRow = FindNextSectionHeaderRow(ws, sectionRow)
    assemblyName = modUtils.NormalizeText(assemblyName)

    For rowNumber = sectionRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            currentAssemblyName = GetAssemblyNameAtRow(ws, rowNumber)

            If StrComp(currentAssemblyName, assemblyName, vbTextCompare) = 0 Then
                FindAssemblyHeaderRowInSection = rowNumber
                Exit Function
            End If
        End If
    Next rowNumber

    FindAssemblyHeaderRowInSection = 0
End Function

' =========================================================
' Insert row calculations
' =========================================================

' Returns the row where a new custom component should be inserted:
' - before the next assembly header inside the same section
' - or before the next section header if there is no next assembly
Public Function FindInsertRowForCustomComponent(ByVal ws As Worksheet, _
                                                ByVal sectionName As String, _
                                                ByVal assemblyName As String) As Long
    Dim sectionRow As Long
    Dim assemblyHeaderRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long

    sectionRow = FindSectionRow(ws, sectionName)
    If sectionRow = 0 Then
        FindInsertRowForCustomComponent = 0
        Exit Function
    End If

    assemblyHeaderRow = FindAssemblyHeaderRowInSection(ws, sectionName, assemblyName)
    If assemblyHeaderRow = 0 Then
        FindInsertRowForCustomComponent = 0
        Exit Function
    End If

    nextSectionRow = FindNextSectionHeaderRow(ws, sectionRow)

    For rowNumber = assemblyHeaderRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            FindInsertRowForCustomComponent = rowNumber
            Exit Function
        End If
    Next rowNumber

    FindInsertRowForCustomComponent = nextSectionRow
End Function

' Returns the row numbers of all assembly header rows inside the specified section.
Public Function GetAssemblyHeaderRowsInSection(ByVal ws As Worksheet, _
                                               ByVal sectionName As String, _
                                               ByRef assemblyRows() As Long, _
                                               ByRef errorMessage As String) As Boolean
    Dim sectionRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long
    Dim count As Long

    errorMessage = vbNullString
    count = 0

    sectionRow = FindSectionRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        GetAssemblyHeaderRowsInSection = False
        Exit Function
    End If

    nextSectionRow = FindNextSectionHeaderRow(ws, sectionRow)

    For rowNumber = sectionRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            count = count + 1
            ReDim Preserve assemblyRows(1 To count)
            assemblyRows(count) = rowNumber
        End If
    Next rowNumber

    If count = 0 Then
        errorMessage = "No assembly header rows were found inside section '" & sectionName & "'."
        GetAssemblyHeaderRowsInSection = False
        Exit Function
    End If

    GetAssemblyHeaderRowsInSection = True
End Function


' Returns the last row that belongs to the specified assembly block inside the section.
' The block ends at the row before the next assembly header,
' or at the row before the next section header if there is no next assembly.
Public Function FindAssemblyBlockEndRowInSection(ByVal ws As Worksheet, _
                                                 ByVal sectionName As String, _
                                                 ByVal assemblyName As String) As Long
    Dim assemblyHeaderRow As Long
    Dim sectionRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long

    sectionRow = FindSectionRow(ws, sectionName)
    If sectionRow = 0 Then
        FindAssemblyBlockEndRowInSection = 0
        Exit Function
    End If

    assemblyHeaderRow = FindAssemblyHeaderRowInSection(ws, sectionName, assemblyName)
    If assemblyHeaderRow = 0 Then
        FindAssemblyBlockEndRowInSection = 0
        Exit Function
    End If

    nextSectionRow = FindNextSectionHeaderRow(ws, sectionRow)

    For rowNumber = assemblyHeaderRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            FindAssemblyBlockEndRowInSection = rowNumber - 1
            Exit Function
        End If
    Next rowNumber

    FindAssemblyBlockEndRowInSection = nextSectionRow - 1
End Function

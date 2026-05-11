Option Explicit

' =========================================================
' clsAssemblyService
'
' Purpose:
' Coordinates assembly-related workflows, including:
' - reading available assembly items from the library sheet
' - inserting one or more assemblies into a target section
'
' This class relies on:
' - clsSheetStructureNavigator for locating positions in the target sheet
' =========================================================

' =========================================================
' Worksheet structure constants
' =========================================================
Private Const LIBRARY_FIRST_DATA_ROW As Long = 3
Private Const LIBRARY_NAME_COLUMN As Long = 2          ' Column B
Private Const LIBRARY_COLOR_COLUMN As Long = 1         ' Column A
Private Const TARGET_INSERT_COLUMN As Long = 2         ' Copy starts at column B

' =========================================================
' Public API
' =========================================================

' Returns the display names of available assemblies for the requested domain.
' The text is returned exactly as stored in the worksheet so it can later be
' used as the lookup value during insertion.
Public Function GetAssemblyDisplayList(ByVal wsLib As Worksheet, _
                                       ByVal domain As String, _
                                       ByRef items() As String, _
                                       ByRef errorMessage As String) As Boolean
    Dim lastRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim rowNumber As Long
    Dim count As Long
    Dim assemblyHeaderColor As Long
    Dim displayText As String

    errorMessage = vbNullString
    count = 0

    lastRow = modUtils.GetLastUsedRow(wsLib)
    If lastRow = 0 Then
        errorMessage = "'Assembly List' appears empty."
        GetAssemblyDisplayList = False
        Exit Function
    End If

    If Not GetDomainBounds(wsLib, domain, LIBRARY_FIRST_DATA_ROW, lastRow, startRow, endRow, errorMessage) Then
        GetAssemblyDisplayList = False
        Exit Function
    End If

    ' The library uses a green header row to identify each assembly block.
    assemblyHeaderColor = wsLib.Cells(LIBRARY_FIRST_DATA_ROW, LIBRARY_COLOR_COLUMN).Interior.Color

    For rowNumber = startRow To endRow
        If wsLib.Cells(rowNumber, LIBRARY_COLOR_COLUMN).Interior.Color = assemblyHeaderColor Then
            displayText = CStr(wsLib.Cells(rowNumber, LIBRARY_NAME_COLUMN).Value2)

            If displayText <> "" Then
                count = count + 1
                ReDim Preserve items(1 To count)
                items(count) = displayText
            End If
        End If
    Next rowNumber

    If count = 0 Then
        errorMessage = "No " & UCase$(domain) & " assembly items found in 'Assembly List'."
        GetAssemblyDisplayList = False
        Exit Function
    End If

    SortStringArray items
    GetAssemblyDisplayList = True
End Function

' Inserts the requested assembly items into the specified section of the target sheet.
Public Function InsertAssembliesIntoSection(ByVal wsTarget As Worksheet, _
                                            ByVal sectionName As String, _
                                            ByRef items() As String, _
                                            ByVal domain As String, _
                                            ByRef errorMessage As String) As Boolean
    Dim wsLib As Worksheet
    Dim navigator As clsSheetStructureNavigator

    Dim sectionRow As Long
    Dim insertRow As Long
    Dim itemIndex As Long
    Dim success As Boolean
    Dim oneError As String

    Dim lastRow As Long
    Dim domainStartRow As Long
    Dim domainEndRow As Long

    errorMessage = vbNullString
    Set navigator = New clsSheetStructureNavigator
    Set wsLib = ThisWorkbook.Worksheets("Assembly List")

    lastRow = modUtils.GetLastUsedRow(wsLib)
    If lastRow = 0 Then
        errorMessage = "'Assembly List' appears empty."
        InsertAssembliesIntoSection = False
        Exit Function
    End If

    If Not GetDomainBounds(wsLib, domain, LIBRARY_FIRST_DATA_ROW, lastRow, domainStartRow, domainEndRow, errorMessage) Then
        InsertAssembliesIntoSection = False
        Exit Function
    End If

    sectionRow = navigator.FindSectionRow(wsTarget, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found in the active sheet."
        InsertAssembliesIntoSection = False
        Exit Function
    End If

    For itemIndex = LBound(items) To UBound(items)
        ' Recalculate the section boundary before each insert,
        ' because previous inserts push the next section downward.
        insertRow = navigator.FindNextSectionHeaderRow(wsTarget, sectionRow)

        success = CopyAssemblyBlockIntoTarget( _
            wsLib, _
            wsTarget, _
            items(itemIndex), _
            insertRow, _
            domainStartRow, _
            domainEndRow, _
            oneError)

        If Not success Then
            errorMessage = oneError
            InsertAssembliesIntoSection = False
            Exit Function
        End If
    Next itemIndex

    Dim sectionManager As clsSectionManager

    Set sectionManager = New clsSectionManager
    
    If Not sectionManager.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertAssembliesIntoSection = False
        Exit Function
    End If
    
    InsertAssembliesIntoSection = True
End Function

' =========================================================
' Core copy workflow
' =========================================================

' Copies one assembly block from the library sheet into the target sheet.
Private Function CopyAssemblyBlockIntoTarget(ByVal wsLib As Worksheet, _
                                             ByVal wsTarget As Worksheet, _
                                             ByVal displayName As String, _
                                             ByVal insertRow As Long, _
                                             ByVal domainStartRow As Long, _
                                             ByVal domainEndRow As Long, _
                                             ByRef errorMessage As String) As Boolean
    Dim startCell As Range
    Dim sourceStartRow As Long
    Dim sourceEndRow As Long
    Dim lastCol As Long
    Dim rowsToInsert As Long
    Dim colsToCopy As Long

    Dim src As Range
    Dim dst As Range
    Dim insertedRows As Range

    Dim firstChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString

    If displayName = "" Then
        errorMessage = "Invalid item name."
        CopyAssemblyBlockIntoTarget = False
        Exit Function
    End If

    ' Find the assembly header row by its display name in column B,
    ' restricted to the requested domain window.
    Set startCell = FindInColumnWithinRows( _
        wsLib, _
        LIBRARY_NAME_COLUMN, _
        displayName, _
        domainStartRow, _
        domainEndRow)

    If startCell Is Nothing Then
        errorMessage = "Assembly '" & displayName & "' not found in '" & wsLib.name & "' within the selected domain."
        CopyAssemblyBlockIntoTarget = False
        Exit Function
    End If

    sourceStartRow = startCell.row

    ' Determine the last used column in the assembly header row.
    lastCol = wsLib.Cells(sourceStartRow, wsLib.Columns.count).End(xlToLeft).Column
    If lastCol < TARGET_INSERT_COLUMN Then
        lastCol = TARGET_INSERT_COLUMN
    End If

    ' Determine the end of the source block without crossing the domain boundary.
    sourceEndRow = FindEndRowTrimTrailingEmpties(wsLib, sourceStartRow, lastCol, domainEndRow)

    rowsToInsert = sourceEndRow - sourceStartRow + 1
    colsToCopy = lastCol - TARGET_INSERT_COLUMN + 1

    Set src = wsLib.Range( _
        wsLib.Cells(sourceStartRow, TARGET_INSERT_COLUMN), _
        wsLib.Cells(sourceEndRow, lastCol))

    ' Insert destination rows first to preserve the worksheet structure.
    wsTarget.Rows(insertRow & ":" & (insertRow + rowsToInsert - 1)).Insert Shift:=xlShiftDown

    Set insertedRows = wsTarget.Rows(insertRow & ":" & (insertRow + rowsToInsert - 1))

    ' Remove inherited formatting and outline/group settings.
    insertedRows.ClearFormats
    insertedRows.ClearOutline

    Set dst = wsTarget.Range( _
        wsTarget.Cells(insertRow, TARGET_INSERT_COLUMN), _
        wsTarget.Cells(insertRow + rowsToInsert - 1, TARGET_INSERT_COLUMN + colsToCopy - 1))

    src.Copy Destination:=dst
    Application.CutCopyMode = False

    ' Apply the expected format to the inserted assembly header row.
    FormatCoreItemRow wsTarget, insertRow

    ' Group only the child rows below the assembly header.
    firstChildRow = insertRow + 1
    lastChildRow = insertRow + rowsToInsert - 1

    If lastChildRow >= firstChildRow Then
        wsTarget.Rows(firstChildRow & ":" & lastChildRow).Group
    End If

    CopyAssemblyBlockIntoTarget = True
End Function

' =========================================================
' Domain bounds
' =========================================================

' Returns the row bounds for the requested domain inside the library sheet.
Private Function GetDomainBounds(ByVal ws As Worksheet, _
                                 ByVal domain As String, _
                                 ByVal firstDataRow As Long, _
                                 ByVal lastRow As Long, _
                                 ByRef startRow As Long, _
                                 ByRef endRow As Long, _
                                 ByRef errorMessage As String) As Boolean
    Dim meEndRow As Long
    Dim ceStartRow As Long

    errorMessage = vbNullString
    domain = UCase$(domain)

    meEndRow = FindMarkerRow(ws, "ME", "END", firstDataRow, lastRow)
    ceStartRow = FindMarkerRow(ws, "CE", "START", firstDataRow, lastRow)

    Select Case domain
        Case "ME"
            startRow = firstDataRow

            If meEndRow > 0 Then
                endRow = meEndRow - 1
            ElseIf ceStartRow > 0 Then
                endRow = ceStartRow - 1
            Else
                endRow = lastRow
            End If

        Case "CE"
            If ceStartRow > 0 Then
                startRow = ceStartRow + 1
            ElseIf meEndRow > 0 Then
                startRow = meEndRow + 1
            Else
                startRow = firstDataRow
            End If

            endRow = lastRow

        Case Else
            errorMessage = "Unknown domain: " & domain & ". Use 'ME' or 'CE'."
            GetDomainBounds = False
            Exit Function
    End Select

    If startRow < firstDataRow Then startRow = firstDataRow
    If endRow > lastRow Then endRow = lastRow

    If endRow < startRow Then
        errorMessage = "Invalid domain bounds for " & domain & ". Check marker rows in 'Assembly List'."
        GetDomainBounds = False
        Exit Function
    End If

    GetDomainBounds = True
End Function

' Finds a row containing both marker tokens within the specified search window.
Private Function FindMarkerRow(ByVal ws As Worksheet, _
                               ByVal token1 As String, _
                               ByVal token2 As String, _
                               ByVal firstRow As Long, _
                               ByVal lastRow As Long) As Long
    Dim rowNumber As Long
    Dim markerText As String

    For rowNumber = firstRow To lastRow
        markerText = CStr(ws.Cells(rowNumber, LIBRARY_NAME_COLUMN).Value2)
        If Len(markerText) = 0 Then
            markerText = CStr(ws.Cells(rowNumber, LIBRARY_COLOR_COLUMN).Value2)
        End If

        If Len(markerText) > 0 Then
            If InStr(1, markerText, token1, vbTextCompare) > 0 _
               And InStr(1, markerText, token2, vbTextCompare) > 0 Then
                FindMarkerRow = rowNumber
                Exit Function
            End If
        End If
    Next rowNumber
End Function

' =========================================================
' Source block boundary helpers
' =========================================================

' Finds the last meaningful row of the assembly block, trimming trailing empty rows.
Private Function FindEndRowTrimTrailingEmpties(ByVal wsLib As Worksheet, _
                                               ByVal startRow As Long, _
                                               ByVal lastCol As Long, _
                                               ByVal domainEndRow As Long) As Long
    Dim boundaryEnd As Long
    Dim rowNumber As Long

    boundaryEnd = FindBoundaryEndRowUntilNextHeader(wsLib, startRow, domainEndRow)

    For rowNumber = boundaryEnd To startRow Step -1
        If Not IsRowVisiblyEmpty(wsLib, rowNumber, TARGET_INSERT_COLUMN, lastCol) Then
            FindEndRowTrimTrailingEmpties = rowNumber
            Exit Function
        End If
    Next rowNumber

    FindEndRowTrimTrailingEmpties = startRow
End Function

' Finds the natural end of the current assembly block:
' the row before the next assembly header, limited by the domain end row.
Private Function FindBoundaryEndRowUntilNextHeader(ByVal wsLib As Worksheet, _
                                                   ByVal startRow As Long, _
                                                   ByVal domainEndRow As Long) As Long
    Dim rowNumber As Long
    Dim headerColor As Long

    headerColor = wsLib.Cells(startRow, LIBRARY_COLOR_COLUMN).Interior.Color

    For rowNumber = startRow + 1 To domainEndRow
        If wsLib.Cells(rowNumber, LIBRARY_COLOR_COLUMN).Interior.Color = headerColor Then
            FindBoundaryEndRowUntilNextHeader = rowNumber - 1
            Exit Function
        End If
    Next rowNumber

    FindBoundaryEndRowUntilNextHeader = domainEndRow
End Function

' Returns True when the specified row contains no visible values in the requested column range.
Private Function IsRowVisiblyEmpty(ByVal ws As Worksheet, _
                                   ByVal rowIndex As Long, _
                                   ByVal firstCol As Long, _
                                   ByVal lastCol As Long) As Boolean
    Dim colIndex As Long
    Dim cellValue As Variant

    For colIndex = firstCol To lastCol
        cellValue = ws.Cells(rowIndex, colIndex).Value2
        If Len(CStr(cellValue)) > 0 Then
            IsRowVisiblyEmpty = False
            Exit Function
        End If
    Next colIndex

    IsRowVisiblyEmpty = True
End Function

' =========================================================
' Search helpers
' =========================================================

' Finds an exact text match in one column within the specified row boundaries.
Private Function FindInColumnWithinRows(ByVal ws As Worksheet, _
                                        ByVal col As Long, _
                                        ByVal whatText As String, _
                                        ByVal startRow As Long, _
                                        ByVal endRow As Long) As Range
    Dim searchRange As Range
    Dim foundCell As Range

    Set searchRange = ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col))

    Set foundCell = searchRange.Find(What:=whatText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If Not foundCell Is Nothing Then
        Set FindInColumnWithinRows = foundCell
    End If
End Function

' =========================================================
' Utility helpers
' =========================================================

' Sorts the string array alphabetically using case-insensitive comparison.
Private Sub SortStringArray(ByRef arr() As String)
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

' Applies the standard format to the inserted assembly header row.
Private Sub FormatCoreItemRow(ByVal wsTarget As Worksheet, ByVal rowIndex As Long)
    With wsTarget.Rows(rowIndex).Interior
        .Pattern = xlSolid
        .Color = RGB(191, 191, 191)
    End With
End Sub


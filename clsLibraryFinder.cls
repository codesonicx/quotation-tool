Option Explicit

' =========================================================
' clsLibraryFinder
'
' Purpose:
' Reads/searches the Assembly List sheet.
'
' Responsibilities:
' - get available assembly names by domain
' - find ME / CE domain bounds
' - find source assembly block rows
'
' This class should not modify any worksheet.
' =========================================================

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
        errorMessage = "'" & SHEET_ASSEMBLY_LIST & "' appears empty."
        GetAssemblyDisplayList = False
        Exit Function
    End If

    If Not GetDomainBounds(wsLib, domain, FIRST_DATA_ROW, lastRow, startRow, endRow, errorMessage) Then
        GetAssemblyDisplayList = False
        Exit Function
    End If

    assemblyHeaderColor = wsLib.Cells(FIRST_DATA_ROW, COL_LIBRARY_COLOR).Interior.Color

    For rowNumber = startRow To endRow
        If wsLib.Cells(rowNumber, COL_LIBRARY_COLOR).Interior.Color = assemblyHeaderColor Then
            displayText = CStr(wsLib.Cells(rowNumber, COL_LIBRARY_NAME).Value2)

            If displayText <> "" Then
                count = count + 1
                ReDim Preserve items(1 To count)
                items(count) = displayText
            End If
        End If
    Next rowNumber

    If count = 0 Then
        errorMessage = "No " & UCase$(domain) & " assembly items found in '" & SHEET_ASSEMBLY_LIST & "'."
        GetAssemblyDisplayList = False
        Exit Function
    End If

    SortStringArray items
    GetAssemblyDisplayList = True
End Function

Public Function GetDomainBounds(ByVal ws As Worksheet, _
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

    meEndRow = FindMarkerRow(ws, DOMAIN_ME, "END", firstDataRow, lastRow)
    ceStartRow = FindMarkerRow(ws, DOMAIN_CE, "START", firstDataRow, lastRow)

    Select Case domain
        Case DOMAIN_ME
            startRow = firstDataRow

            If meEndRow > 0 Then
                endRow = meEndRow - 1
            ElseIf ceStartRow > 0 Then
                endRow = ceStartRow - 1
            Else
                endRow = lastRow
            End If

        Case DOMAIN_CE
            If ceStartRow > 0 Then
                startRow = ceStartRow + 1
            ElseIf meEndRow > 0 Then
                startRow = meEndRow + 1
            Else
                startRow = firstDataRow
            End If

            endRow = lastRow

        Case Else
            errorMessage = "Unknown domain: " & domain & ". Use '" & DOMAIN_ME & "' or '" & DOMAIN_CE & "'."
            GetDomainBounds = False
            Exit Function
    End Select

    If startRow < firstDataRow Then startRow = firstDataRow
    If endRow > lastRow Then endRow = lastRow

    If endRow < startRow Then
        errorMessage = "Invalid domain bounds for " & domain & ". Check marker rows in '" & SHEET_ASSEMBLY_LIST & "'."
        GetDomainBounds = False
        Exit Function
    End If

    GetDomainBounds = True
End Function

Public Function FindAssemblySourceBlock(ByVal wsLib As Worksheet, _
                                        ByVal displayName As String, _
                                        ByVal domainStartRow As Long, _
                                        ByVal domainEndRow As Long, _
                                        ByRef sourceStartRow As Long, _
                                        ByRef sourceEndRow As Long, _
                                        ByRef lastCol As Long, _
                                        ByRef errorMessage As String) As Boolean
    Dim startCell As Range

    errorMessage = vbNullString

    If displayName = "" Then
        errorMessage = "Invalid item name."
        FindAssemblySourceBlock = False
        Exit Function
    End If

    Set startCell = FindInColumnWithinRows( _
        wsLib, _
        COL_LIBRARY_NAME, _
        displayName, _
        domainStartRow, _
        domainEndRow)

    If startCell Is Nothing Then
        errorMessage = "Assembly '" & displayName & "' not found in '" & wsLib.Name & "' within the selected domain."
        FindAssemblySourceBlock = False
        Exit Function
    End If

    sourceStartRow = startCell.Row

    lastCol = wsLib.Cells(sourceStartRow, wsLib.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_COPY_START Then
        lastCol = COL_COPY_START
    End If

    sourceEndRow = FindEndRowTrimTrailingEmpties(wsLib, sourceStartRow, lastCol, domainEndRow)

    FindAssemblySourceBlock = True
End Function

Private Function FindMarkerRow(ByVal ws As Worksheet, _
                               ByVal token1 As String, _
                               ByVal token2 As String, _
                               ByVal firstRow As Long, _
                               ByVal lastRow As Long) As Long
    Dim rowNumber As Long
    Dim markerText As String

    For rowNumber = firstRow To lastRow
        markerText = CStr(ws.Cells(rowNumber, COL_LIBRARY_NAME).Value2)

        If Len(markerText) = 0 Then
            markerText = CStr(ws.Cells(rowNumber, COL_LIBRARY_COLOR).Value2)
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

Private Function FindEndRowTrimTrailingEmpties(ByVal wsLib As Worksheet, _
                                               ByVal startRow As Long, _
                                               ByVal lastCol As Long, _
                                               ByVal domainEndRow As Long) As Long
    Dim boundaryEnd As Long
    Dim rowNumber As Long

    boundaryEnd = FindBoundaryEndRowUntilNextHeader(wsLib, startRow, domainEndRow)

    For rowNumber = boundaryEnd To startRow Step -1
        If Not IsRowVisiblyEmpty(wsLib, rowNumber, COL_COPY_START, lastCol) Then
            FindEndRowTrimTrailingEmpties = rowNumber
            Exit Function
        End If
    Next rowNumber

    FindEndRowTrimTrailingEmpties = startRow
End Function

Private Function FindBoundaryEndRowUntilNextHeader(ByVal wsLib As Worksheet, _
                                                   ByVal startRow As Long, _
                                                   ByVal domainEndRow As Long) As Long
    Dim rowNumber As Long
    Dim headerColor As Long

    headerColor = wsLib.Cells(startRow, COL_LIBRARY_COLOR).Interior.Color

    For rowNumber = startRow + 1 To domainEndRow
        If wsLib.Cells(rowNumber, COL_LIBRARY_COLOR).Interior.Color = headerColor Then
            FindBoundaryEndRowUntilNextHeader = rowNumber - 1
            Exit Function
        End If
    Next rowNumber

    FindBoundaryEndRowUntilNextHeader = domainEndRow
End Function

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

Private Function FindInColumnWithinRows(ByVal ws As Worksheet, _
                                        ByVal col As Long, _
                                        ByVal whatText As String, _
                                        ByVal startRow As Long, _
                                        ByVal endRow As Long) As Range
    Dim searchRange As Range
    Dim foundCell As Range

    Set searchRange = ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col))

    Set foundCell = searchRange.Find( _
        What:=whatText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If Not foundCell Is Nothing Then
        Set FindInColumnWithinRows = foundCell
    End If
End Function

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
Option Explicit

' =========================================================
' clsSectionFinder
'
' Purpose:
' Finds and identifies section header rows in a target worksheet.
'
' Rules:
' - Section header rows are blue.
' - Section names are stored in column A.
' =========================================================

Private Function GetNormalizedCellText(ByVal ws As Worksheet, _
                                       ByVal rowNumber As Long, _
                                       ByVal columnNumber As Long) As String
    GetNormalizedCellText = modUtils.NormalizeText(CStr(ws.Cells(rowNumber, columnNumber).Value2))
End Function

Private Function GetSectionNameAtRow(ByVal ws As Worksheet, _
                                     ByVal rowNumber As Long) As String
    GetSectionNameAtRow = GetNormalizedCellText(ws, rowNumber, COL_SECTION_NAME)
End Function

Public Function IsSectionHeaderRow(ByVal ws As Worksheet, _
                                   ByVal rowNumber As Long) As Boolean
    IsSectionHeaderRow = _
        (ws.Rows(rowNumber).Interior.ColorIndex = COLOR_INDEX_SECTION_BLUE) _
        And (GetSectionNameAtRow(ws, rowNumber) <> "")
End Function

Public Function FindSectionHeaderRow(ByVal ws As Worksheet, _
                                     ByVal sectionName As String) As Long
    Dim matchCell As Range

    sectionName = modUtils.NormalizeText(sectionName)

    If sectionName = "" Then
        FindSectionHeaderRow = 0
        Exit Function
    End If

    Set matchCell = ws.Columns(COL_SECTION_NAME).Find( _
        What:=sectionName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If matchCell Is Nothing Then
        FindSectionHeaderRow = 0
        Exit Function
    End If

    If IsSectionHeaderRow(ws, matchCell.Row) Then
        FindSectionHeaderRow = matchCell.Row
    Else
        FindSectionHeaderRow = 0
    End If
End Function

Public Function FindNextSectionHeaderRowAfter(ByVal ws As Worksheet, _
                                              ByVal sectionHeaderRow As Long) As Long
    Dim lastRow As Long
    Dim rowNumber As Long

    lastRow = modUtils.GetLastUsedRow(ws)

    If lastRow = 0 Then
        FindNextSectionHeaderRowAfter = sectionHeaderRow + 1
        Exit Function
    End If

    For rowNumber = sectionHeaderRow + 1 To lastRow
        If IsSectionHeaderRow(ws, rowNumber) Then
            FindNextSectionHeaderRowAfter = rowNumber
            Exit Function
        End If
    Next rowNumber

    FindNextSectionHeaderRowAfter = lastRow + 1
End Function

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
        errorMessage = "No sections found."
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
        errorMessage = "No sections found."
        GetSectionNames = False
        Exit Function
    End If

    GetSectionNames = True
End Function
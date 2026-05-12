Option Explicit

' =========================================================
' clsAssemblyService
'
' Purpose:
' Inserts assembly blocks from the Assembly List sheet into
' a selected section of the target worksheet.
'
' Responsibilities:
' - coordinate assembly insertion workflow
' - insert copied rows into target worksheet
' - apply assembly row formatting
' - group child rows
' - refresh section totals
' =========================================================

Public Function GetAssemblyDisplayList(ByVal wsLib As Worksheet, _
                                       ByVal domain As String, _
                                       ByRef items() As String, _
                                       ByRef errorMessage As String) As Boolean
    Dim libraryFinder As clsLibraryFinder

    Set libraryFinder = New clsLibraryFinder

    GetAssemblyDisplayList = libraryFinder.GetAssemblyDisplayList( _
        wsLib, _
        domain, _
        items, _
        errorMessage)
End Function

Public Function InsertAssembliesIntoSection(ByVal wsTarget As Worksheet, _
                                            ByVal sectionName As String, _
                                            ByRef items() As String, _
                                            ByVal domain As String, _
                                            ByRef errorMessage As String) As Boolean
    Dim wsLib As Worksheet

    Dim sectionFinder As clsSectionFinder
    Dim sectionService As clsSectionService
    Dim libraryFinder As clsLibraryFinder

    Dim sectionRow As Long
    Dim insertRow As Long
    Dim itemIndex As Long
    Dim success As Boolean
    Dim oneError As String

    Dim lastRow As Long
    Dim domainStartRow As Long
    Dim domainEndRow As Long

    errorMessage = vbNullString

    Set sectionFinder = New clsSectionFinder
    Set sectionService = New clsSectionService
    Set libraryFinder = New clsLibraryFinder

    Set wsLib = ThisWorkbook.Worksheets(SHEET_ASSEMBLY_LIST)

    lastRow = modUtils.GetLastUsedRow(wsLib)
    If lastRow = 0 Then
        errorMessage = "'" & SHEET_ASSEMBLY_LIST & "' appears empty."
        InsertAssembliesIntoSection = False
        Exit Function
    End If

    If Not libraryFinder.GetDomainBounds(wsLib, domain, FIRST_DATA_ROW, lastRow, domainStartRow, domainEndRow, errorMessage) Then
        InsertAssembliesIntoSection = False
        Exit Function
    End If

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found in the active sheet."
        InsertAssembliesIntoSection = False
        Exit Function
    End If

    For itemIndex = LBound(items) To UBound(items)
        ' Recalculate the section boundary before each insert,
        ' because previous inserts push the next section downward.
        insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

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

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertAssembliesIntoSection = False
        Exit Function
    End If

    InsertAssembliesIntoSection = True
End Function

Private Function CopyAssemblyBlockIntoTarget(ByVal wsLib As Worksheet, _
                                             ByVal wsTarget As Worksheet, _
                                             ByVal displayName As String, _
                                             ByVal insertRow As Long, _
                                             ByVal domainStartRow As Long, _
                                             ByVal domainEndRow As Long, _
                                             ByRef errorMessage As String) As Boolean
    Dim libraryFinder As clsLibraryFinder

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
    Set libraryFinder = New clsLibraryFinder

    If Not libraryFinder.FindAssemblySourceBlock( _
        wsLib, _
        displayName, _
        domainStartRow, _
        domainEndRow, _
        sourceStartRow, _
        sourceEndRow, _
        lastCol, _
        errorMessage) Then

        CopyAssemblyBlockIntoTarget = False
        Exit Function
    End If

    rowsToInsert = sourceEndRow - sourceStartRow + 1
    colsToCopy = lastCol - COL_COPY_START + 1

    Set src = wsLib.Range( _
        wsLib.Cells(sourceStartRow, COL_COPY_START), _
        wsLib.Cells(sourceEndRow, lastCol))

    wsTarget.Rows(insertRow & ":" & (insertRow + rowsToInsert - 1)).Insert Shift:=xlShiftDown

    Set insertedRows = wsTarget.Rows(insertRow & ":" & (insertRow + rowsToInsert - 1))

    ' Remove inherited formatting and outline/group settings.
    insertedRows.ClearFormats
    insertedRows.ClearOutline

    Set dst = wsTarget.Range( _
        wsTarget.Cells(insertRow, COL_COPY_START), _
        wsTarget.Cells(insertRow + rowsToInsert - 1, COL_COPY_START + colsToCopy - 1))

    src.Copy Destination:=dst
    Application.CutCopyMode = False

    FormatCoreItemRow wsTarget, insertRow

    ' Group only the child rows below the assembly header.
    firstChildRow = insertRow + 1
    lastChildRow = insertRow + rowsToInsert - 1

    If lastChildRow >= firstChildRow Then
        wsTarget.Rows(firstChildRow & ":" & lastChildRow).Group
    End If

    CopyAssemblyBlockIntoTarget = True
End Function

Private Sub FormatCoreItemRow(ByVal wsTarget As Worksheet, ByVal rowIndex As Long)
    With wsTarget.Rows(rowIndex).Interior
        .Pattern = xlSolid
        .Color = RGB(COLOR_ASSEMBLY_GRAY_R, COLOR_ASSEMBLY_GRAY_G, COLOR_ASSEMBLY_GRAY_B)
    End With
End Sub
Option Explicit

' =========================================================
' clsAssemblyLibraryService
'
' Purpose:
' Saves an existing assembly block from the active estimate
' sheet into the Assembly List library.
'
' Direction:
' Target estimate sheet -> Assembly List
'
' Responsibilities:
' - find the source assembly block
' - determine the correct ME / CE library insertion point
' - generate a unique assembly name if needed
' - copy columns B:AG into Assembly List
' - recolor the assembly header row as a library row
' - write component ID lookup formulas in column A
' =========================================================

Public Function SaveAssemblyToLibrary(ByVal wsTarget As Worksheet, _
                                      ByVal sectionName As String, _
                                      ByVal assemblyName As String, _
                                      ByVal domain As String, _
                                      ByRef savedAssemblyName As String, _
                                      ByRef errorMessage As String) As Boolean
    Dim wsLib As Worksheet

    Dim assemblyFinder As clsAssemblyFinder
    Dim libraryFinder As clsLibraryFinder

    Dim sourceStartRow As Long
    Dim sourceEndRow As Long
    Dim rowsToCopy As Long

    Dim lastLibRow As Long
    Dim domainStartRow As Long
    Dim domainEndRow As Long
    Dim targetInsertRow As Long
    Dim targetEndRow As Long

    Dim src As Range
    Dim dst As Range

    errorMessage = vbNullString
    savedAssemblyName = vbNullString

    domain = UCase$(Trim$(domain))

    If domain <> DOMAIN_ME And domain <> DOMAIN_CE Then
        errorMessage = "Invalid library domain: [" & domain & "]."
        SaveAssemblyToLibrary = False
        Exit Function
    End If

    sectionName = modUtils.NormalizeText(sectionName)

    If sectionName = vbNullString Then
        errorMessage = "Section name cannot be empty."
        SaveAssemblyToLibrary = False
        Exit Function
    End If

    If assemblyName = vbNullString Then
        errorMessage = "Assembly name cannot be empty."
        SaveAssemblyToLibrary = False
        Exit Function
    End If

    Set wsLib = ThisWorkbook.Worksheets(SHEET_ASSEMBLY_LIST)
    Set assemblyFinder = New clsAssemblyFinder
    Set libraryFinder = New clsLibraryFinder

    sourceStartRow = assemblyFinder.FindAssemblyHeaderRow(wsTarget, sectionName, assemblyName)

    If sourceStartRow = 0 Then
        errorMessage = "Assembly '" & assemblyName & "' was not found in section '" & sectionName & "'."
        SaveAssemblyToLibrary = False
        Exit Function
    End If

    sourceEndRow = assemblyFinder.FindAssemblyEndRow(wsTarget, sectionName, assemblyName)

    If sourceEndRow < sourceStartRow Then
        errorMessage = "Could not determine the full assembly block for '" & assemblyName & "'."
        SaveAssemblyToLibrary = False
        Exit Function
    End If

    lastLibRow = modUtils.GetLastUsedRow(wsLib)

    If lastLibRow = 0 Then
        errorMessage = "'" & SHEET_ASSEMBLY_LIST & "' appears empty."
        SaveAssemblyToLibrary = False
        Exit Function
    End If

    If Not libraryFinder.GetDomainBounds( _
        wsLib, _
        domain, _
        FIRST_DATA_ROW, _
        lastLibRow, _
        domainStartRow, _
        domainEndRow, _
        errorMessage) Then

        SaveAssemblyToLibrary = False
        Exit Function
    End If

    savedAssemblyName = GetUniqueLibraryAssemblyName( _
        wsLib, _
        assemblyName, _
        domainStartRow, _
        domainEndRow)

    rowsToCopy = sourceEndRow - sourceStartRow + 1
    targetInsertRow = domainEndRow + 1
    targetEndRow = targetInsertRow + rowsToCopy - 1

    wsLib.Rows(targetInsertRow & ":" & targetEndRow).Insert Shift:=xlShiftDown

    Set src = wsTarget.Range( _
        wsTarget.Cells(sourceStartRow, COL_COPY_START), _
        wsTarget.Cells(sourceEndRow, COL_PERCENT))

    Set dst = wsLib.Range( _
        wsLib.Cells(targetInsertRow, COL_COPY_START), _
        wsLib.Cells(targetEndRow, COL_PERCENT))

    src.Copy Destination:=dst
    Application.CutCopyMode = False

    wsLib.Rows(targetInsertRow & ":" & targetEndRow).ClearOutline

    wsLib.Cells(targetInsertRow, COL_LIBRARY_COLOR).ClearContents
    wsLib.Cells(targetInsertRow, COL_LIBRARY_NAME).Value = savedAssemblyName

    ApplyLibraryAssemblyHeaderFormat wsLib, targetInsertRow
    WriteComponentIdFormulas wsLib, targetInsertRow + 1, targetEndRow, domain

    SaveAssemblyToLibrary = True
End Function

Private Sub ApplyLibraryAssemblyHeaderFormat(ByVal wsLib As Worksheet, _
                                             ByVal rowIndex As Long)
    Dim libraryHeaderColor As Long

    libraryHeaderColor = wsLib.Cells(FIRST_DATA_ROW, COL_LIBRARY_COLOR).Interior.Color

    wsLib.Range( _
        wsLib.Cells(rowIndex, COL_LIBRARY_COLOR), _
        wsLib.Cells(rowIndex, COL_PERCENT)).Interior.Color = libraryHeaderColor
End Sub

Private Sub WriteComponentIdFormulas(ByVal wsLib As Worksheet, _
                                     ByVal firstChildRow As Long, _
                                     ByVal lastChildRow As Long, _
                                     ByVal domain As String)
    Dim rowIndex As Long
    Dim tableName As String

    If lastChildRow < firstChildRow Then Exit Sub

    tableName = GetComponentListTableName(domain)

    For rowIndex = firstChildRow To lastChildRow
        wsLib.Cells(rowIndex, COL_LIBRARY_COLOR).Formula = _
            "=IF(ISBLANK($F" & rowIndex & "),"""",INDEX(" & tableName & "[ID],XMATCH($F" & rowIndex & "," & tableName & "[items],0)))"
    Next rowIndex
End Sub

Private Function GetComponentListTableName(ByVal domain As String) As String
    Select Case UCase$(Trim$(domain))
        Case DOMAIN_ME
            GetComponentListTableName = TABLE_ME_COMPONENT_LIST

        Case DOMAIN_CE
            GetComponentListTableName = TABLE_CE_COMPONENT_LIST
    End Select
End Function

Private Function GetUniqueLibraryAssemblyName(ByVal wsLib As Worksheet, _
                                              ByVal requestedName As String, _
                                              ByVal domainStartRow As Long, _
                                              ByVal domainEndRow As Long) As String
    Dim baseName As String
    Dim candidateName As String
    Dim suffixNumber As Long

    baseName = StripGeneratedSuffix(requestedName)
    candidateName = baseName

    If Not LibraryAssemblyNameExists(wsLib, candidateName, domainStartRow, domainEndRow) Then
        GetUniqueLibraryAssemblyName = candidateName
        Exit Function
    End If

    suffixNumber = 2

    Do
        candidateName = baseName & " - (" & suffixNumber & ")"

        If Not LibraryAssemblyNameExists(wsLib, candidateName, domainStartRow, domainEndRow) Then
            GetUniqueLibraryAssemblyName = candidateName
            Exit Function
        End If

        suffixNumber = suffixNumber + 1
    Loop
End Function

Private Function LibraryAssemblyNameExists(ByVal wsLib As Worksheet, _
                                           ByVal assemblyName As String, _
                                           ByVal domainStartRow As Long, _
                                           ByVal domainEndRow As Long) As Boolean
    Dim searchRange As Range
    Dim foundCell As Range

    Set searchRange = wsLib.Range( _
        wsLib.Cells(domainStartRow, COL_LIBRARY_NAME), _
        wsLib.Cells(domainEndRow, COL_LIBRARY_NAME))

    Set foundCell = searchRange.Find( _
        What:=assemblyName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    LibraryAssemblyNameExists = Not foundCell Is Nothing
End Function

Private Function StripGeneratedSuffix(ByVal assemblyName As String) As String
    Dim openPos As Long
    Dim closePos As Long
    Dim possibleNumber As String
    Dim prefixText As String

    assemblyName = Trim$(assemblyName)

    If Right$(assemblyName, 1) <> ")" Then
        StripGeneratedSuffix = assemblyName
        Exit Function
    End If

    openPos = InStrRev(assemblyName, " - (")
    closePos = Len(assemblyName)

    If openPos = 0 Then
        StripGeneratedSuffix = assemblyName
        Exit Function
    End If

    possibleNumber = Mid$(assemblyName, openPos + 4, closePos - (openPos + 4))

    If Not IsNumeric(possibleNumber) Then
        StripGeneratedSuffix = assemblyName
        Exit Function
    End If

    prefixText = Left$(assemblyName, openPos - 1)
    StripGeneratedSuffix = Trim$(prefixText)
End Function
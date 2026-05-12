Option Explicit

' =========================================================
' clsAssemblyFinder
'
' Purpose:
' Finds and identifies assembly header rows inside a section.
'
' Rules:
' - Assembly header rows are gray.
' - Assembly names are stored in column B.
' - Assemblies always belong to a section.
' =========================================================

Private Function GetNormalizedCellText(ByVal ws As Worksheet, _
                                       ByVal rowNumber As Long, _
                                       ByVal columnNumber As Long) As String
    GetNormalizedCellText = modUtils.NormalizeText(CStr(ws.Cells(rowNumber, columnNumber).Value2))
End Function

Private Function GetAssemblyNameAtRow(ByVal ws As Worksheet, _
                                      ByVal rowNumber As Long) As String
    GetAssemblyNameAtRow = GetNormalizedCellText(ws, rowNumber, COL_ASSEMBLY_NAME)
End Function

Public Function IsAssemblyHeaderRow(ByVal ws As Worksheet, _
                                    ByVal rowNumber As Long) As Boolean
    IsAssemblyHeaderRow = _
        (ws.Rows(rowNumber).Interior.ColorIndex = COLOR_INDEX_ASSEMBLY_GRAY) _
        And (GetAssemblyNameAtRow(ws, rowNumber) <> "")
End Function

Public Function GetAssembliesInSection(ByVal ws As Worksheet, _
                                       ByVal sectionName As String, _
                                       ByRef assemblies() As String, _
                                       ByRef errorMessage As String) As Boolean
    Dim sectionFinder As clsSectionFinder
    Dim sectionRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long
    Dim count As Long
    Dim assemblyText As String

    errorMessage = vbNullString
    count = 0

    Set sectionFinder = New clsSectionFinder

    sectionRow = sectionFinder.FindSectionHeaderRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        GetAssembliesInSection = False
        Exit Function
    End If

    nextSectionRow = sectionFinder.FindNextSectionHeaderRowAfter(ws, sectionRow)

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

Public Function FindAssemblyHeaderRow(ByVal ws As Worksheet, _
                                      ByVal sectionName As String, _
                                      ByVal assemblyName As String) As Long
    Dim sectionFinder As clsSectionFinder
    Dim sectionRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long
    Dim currentAssemblyName As String

    Set sectionFinder = New clsSectionFinder

    sectionRow = sectionFinder.FindSectionHeaderRow(ws, sectionName)
    If sectionRow = 0 Then
        FindAssemblyHeaderRow = 0
        Exit Function
    End If

    nextSectionRow = sectionFinder.FindNextSectionHeaderRowAfter(ws, sectionRow)
    assemblyName = modUtils.NormalizeText(assemblyName)

    For rowNumber = sectionRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            currentAssemblyName = GetAssemblyNameAtRow(ws, rowNumber)

            If StrComp(currentAssemblyName, assemblyName, vbTextCompare) = 0 Then
                FindAssemblyHeaderRow = rowNumber
                Exit Function
            End If
        End If
    Next rowNumber

    FindAssemblyHeaderRow = 0
End Function

Public Function FindCustomComponentInsertRow(ByVal ws As Worksheet, _
                                             ByVal sectionName As String, _
                                             ByVal assemblyName As String) As Long
    Dim sectionFinder As clsSectionFinder
    Dim sectionRow As Long
    Dim assemblyHeaderRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long

    Set sectionFinder = New clsSectionFinder

    sectionRow = sectionFinder.FindSectionHeaderRow(ws, sectionName)
    If sectionRow = 0 Then
        FindCustomComponentInsertRow = 0
        Exit Function
    End If

    assemblyHeaderRow = FindAssemblyHeaderRow(ws, sectionName, assemblyName)
    If assemblyHeaderRow = 0 Then
        FindCustomComponentInsertRow = 0
        Exit Function
    End If

    nextSectionRow = sectionFinder.FindNextSectionHeaderRowAfter(ws, sectionRow)

    For rowNumber = assemblyHeaderRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            FindCustomComponentInsertRow = rowNumber
            Exit Function
        End If
    Next rowNumber

    FindCustomComponentInsertRow = nextSectionRow
End Function

Public Function GetAssemblyHeaderRows(ByVal ws As Worksheet, _
                                      ByVal sectionName As String, _
                                      ByRef assemblyRows() As Long, _
                                      ByRef errorMessage As String) As Boolean
    Dim sectionFinder As clsSectionFinder
    Dim sectionRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long
    Dim count As Long

    errorMessage = vbNullString
    count = 0

    Set sectionFinder = New clsSectionFinder

    sectionRow = sectionFinder.FindSectionHeaderRow(ws, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        GetAssemblyHeaderRows = False
        Exit Function
    End If

    nextSectionRow = sectionFinder.FindNextSectionHeaderRowAfter(ws, sectionRow)

    For rowNumber = sectionRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            count = count + 1
            ReDim Preserve assemblyRows(1 To count)
            assemblyRows(count) = rowNumber
        End If
    Next rowNumber

    If count = 0 Then
        errorMessage = "No assembly header rows were found inside section '" & sectionName & "'."
        GetAssemblyHeaderRows = False
        Exit Function
    End If

    GetAssemblyHeaderRows = True
End Function

Public Function FindAssemblyEndRow(ByVal ws As Worksheet, _
                                   ByVal sectionName As String, _
                                   ByVal assemblyName As String) As Long
    Dim sectionFinder As clsSectionFinder
    Dim sectionRow As Long
    Dim assemblyHeaderRow As Long
    Dim nextSectionRow As Long
    Dim rowNumber As Long

    Set sectionFinder = New clsSectionFinder

    sectionRow = sectionFinder.FindSectionHeaderRow(ws, sectionName)
    If sectionRow = 0 Then
        FindAssemblyEndRow = 0
        Exit Function
    End If

    assemblyHeaderRow = FindAssemblyHeaderRow(ws, sectionName, assemblyName)
    If assemblyHeaderRow = 0 Then
        FindAssemblyEndRow = 0
        Exit Function
    End If

    nextSectionRow = sectionFinder.FindNextSectionHeaderRowAfter(ws, sectionRow)

    For rowNumber = assemblyHeaderRow + 1 To nextSectionRow - 1
        If IsAssemblyHeaderRow(ws, rowNumber) Then
            FindAssemblyEndRow = rowNumber - 1
            Exit Function
        End If
    Next rowNumber

    FindAssemblyEndRow = nextSectionRow - 1
End Function
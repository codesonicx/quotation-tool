Option Explicit

' =========================================================
' clsAssemblyEditorService
'
' Purpose:
' Edits existing assemblies by adding/removing child component rows.
'
' Responsibilities:
' - read current child components
' - add ME/CE component list items to an existing assembly
' - remove selected child component rows
' - renumber child rows
' - refresh assembly totals, grouping, section totals, and grand totals
' =========================================================

Public Function GetAssemblyChildren(ByVal wsTarget As Worksheet, _
                                    ByVal sectionName As String, _
                                    ByVal assemblyName As String, _
                                    ByRef childNames() As String, _
                                    ByRef childRows() As Long, _
                                    ByRef childCount As Long, _
                                    ByRef errorMessage As String) As Boolean
    Dim assemblyFinder As clsAssemblyFinder
    Dim assemblyRow As Long
    Dim assemblyEndRow As Long
    Dim rowIndex As Long
    Dim componentName As String

    errorMessage = vbNullString
    childCount = 0

    Set assemblyFinder = New clsAssemblyFinder

    assemblyRow = assemblyFinder.FindAssemblyHeaderRow(wsTarget, sectionName, assemblyName)
    If assemblyRow = 0 Then
        errorMessage = "Assembly '" & assemblyName & "' was not found in section '" & sectionName & "'."
        GetAssemblyChildren = False
        Exit Function
    End If

    assemblyEndRow = assemblyFinder.FindAssemblyEndRow(wsTarget, sectionName, assemblyName)

    If assemblyEndRow <= assemblyRow Then
        GetAssemblyChildren = True
        Exit Function
    End If

    For rowIndex = assemblyRow + 1 To assemblyEndRow
        componentName = Trim$(CStr(wsTarget.Cells(rowIndex, COL_COMPONENT).Value))

        If componentName <> vbNullString Then
            childCount = childCount + 1

            ReDim Preserve childNames(1 To childCount)
            ReDim Preserve childRows(1 To childCount)

            childNames(childCount) = componentName
            childRows(childCount) = rowIndex
        End If
    Next rowIndex

    GetAssemblyChildren = True
End Function

Public Function AddComponentsToAssembly(ByVal wsTarget As Worksheet, _
                                        ByVal sectionName As String, _
                                        ByVal assemblyName As String, _
                                        ByVal domain As String, _
                                        ByRef selectedComponents() As String, _
                                        ByRef errorMessage As String) As Boolean
    Dim assemblyFinder As clsAssemblyFinder
    Dim componentFinder As clsComponentListFinder
    Dim generatedAssemblyService As clsGeneratedAssemblyService

    Dim componentData As Collection
    Dim child As Collection

    Dim insertRow As Long
    Dim i As Long

    errorMessage = vbNullString
    domain = UCase$(Trim$(domain))

    Set assemblyFinder = New clsAssemblyFinder
    Set componentFinder = New clsComponentListFinder
    Set generatedAssemblyService = New clsGeneratedAssemblyService

    insertRow = assemblyFinder.FindCustomComponentInsertRow(wsTarget, sectionName, assemblyName)

    If insertRow = 0 Then
        errorMessage = "Could not determine insert row for assembly '" & assemblyName & "'."
        AddComponentsToAssembly = False
        Exit Function
    End If

    For i = LBound(selectedComponents) To UBound(selectedComponents)

        If Not componentFinder.GetComponentData(domain, selectedComponents(i), componentData, errorMessage) Then
            AddComponentsToAssembly = False
            Exit Function
        End If

        Set child = generatedAssemblyService.CreateComponentListChild(componentData)

        If Not generatedAssemblyService.InsertGeneratedChildRow(wsTarget, insertRow, child, errorMessage) Then
            AddComponentsToAssembly = False
            Exit Function
        End If

        insertRow = insertRow + 1
    Next i

    If Not RefreshAssemblyAfterEdit(wsTarget, sectionName, assemblyName, errorMessage) Then
        AddComponentsToAssembly = False
        Exit Function
    End If

    AddComponentsToAssembly = True
End Function

Public Function RemoveChildrenFromAssembly(ByVal wsTarget As Worksheet, _
                                           ByVal sectionName As String, _
                                           ByVal assemblyName As String, _
                                           ByRef selectedRows() As Long, _
                                           ByRef errorMessage As String) As Boolean
    Dim sortedRows() As Long
    Dim i As Long

    errorMessage = vbNullString

    sortedRows = SortRowsDescending(selectedRows)

    For i = LBound(sortedRows) To UBound(sortedRows)
        wsTarget.Rows(sortedRows(i)).Delete Shift:=xlUp
    Next i

    If Not RefreshAssemblyAfterEdit(wsTarget, sectionName, assemblyName, errorMessage) Then
        RemoveChildrenFromAssembly = False
        Exit Function
    End If

    RemoveChildrenFromAssembly = True
End Function

Private Function RefreshAssemblyAfterEdit(ByVal wsTarget As Worksheet, _
                                          ByVal sectionName As String, _
                                          ByVal assemblyName As String, _
                                          ByRef errorMessage As String) As Boolean
    Dim assemblyFinder As clsAssemblyFinder
    Dim sectionService As clsSectionService
    Dim formulaService As clsFormulaService
    Dim worksheetFormatter As clsWorksheetFormatter

    Dim assemblyRow As Long
    Dim assemblyEndRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString

    Set assemblyFinder = New clsAssemblyFinder
    Set sectionService = New clsSectionService
    Set formulaService = New clsFormulaService
    Set worksheetFormatter = New clsWorksheetFormatter

    assemblyRow = assemblyFinder.FindAssemblyHeaderRow(wsTarget, sectionName, assemblyName)

    If assemblyRow = 0 Then
        errorMessage = "Assembly '" & assemblyName & "' was not found after edit."
        RefreshAssemblyAfterEdit = False
        Exit Function
    End If

    assemblyEndRow = assemblyFinder.FindAssemblyEndRow(wsTarget, sectionName, assemblyName)

    firstChildRow = assemblyRow + 1
    lastChildRow = assemblyEndRow

    If lastChildRow >= firstChildRow Then
        RenumberAssemblyChildren wsTarget, firstChildRow, lastChildRow
    End If

    formulaService.WriteAssemblyTotalFormulas _
        wsTarget, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow

    worksheetFormatter.RefreshAssemblyChildGrouping _
        wsTarget, _
        firstChildRow, _
        lastChildRow

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        RefreshAssemblyAfterEdit = False
        Exit Function
    End If

    RefreshAssemblyAfterEdit = True
End Function

Private Sub RenumberAssemblyChildren(ByVal wsTarget As Worksheet, _
                                     ByVal firstChildRow As Long, _
                                     ByVal lastChildRow As Long)
    Dim rowIndex As Long
    Dim itemNumber As Long

    itemNumber = 1

    For rowIndex = firstChildRow To lastChildRow
        wsTarget.Cells(rowIndex, COL_COMP_LIST).Value = itemNumber
        itemNumber = itemNumber + 1
    Next rowIndex
End Sub

Private Function SortRowsDescending(ByRef sourceRows() As Long) As Long()
    Dim resultRows() As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long

    ReDim resultRows(LBound(sourceRows) To UBound(sourceRows))

    For i = LBound(sourceRows) To UBound(sourceRows)
        resultRows(i) = sourceRows(i)
    Next i

    For i = LBound(resultRows) To UBound(resultRows) - 1
        For j = i + 1 To UBound(resultRows)
            If resultRows(j) > resultRows(i) Then
                temp = resultRows(i)
                resultRows(i) = resultRows(j)
                resultRows(j) = temp
            End If
        Next j
    Next i

    SortRowsDescending = resultRows
End Function
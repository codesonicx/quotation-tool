Option Explicit

' =========================================================
' clsCustomAssemblyBuilderService
'
' Purpose:
' Creates a new generated assembly in a selected section
' using selected components from ME/CE Component List.
'
' Responsibilities:
' - read component data from component list sheet
' - build generated assembly children
' - insert assembly using clsGeneratedAssemblyService
' - refresh section totals and grand totals
' =========================================================

Public Function InsertCustomAssemblyWithComponents(ByVal wsTarget As Worksheet, _
                                                   ByVal sectionName As String, _
                                                   ByVal assemblyName As String, _
                                                   ByVal domain As String, _
                                                   ByRef selectedComponents() As String, _
                                                   ByRef errorMessage As String) As Boolean
    Dim sectionFinder As clsSectionFinder
    Dim sectionService As clsSectionService
    Dim componentFinder As clsComponentListFinder
    Dim generatedAssemblyService As clsGeneratedAssemblyService

    Dim sectionRow As Long
    Dim insertRow As Long

    Dim children As Collection
    Dim componentData As Collection

    Dim i As Long

    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString
    domain = UCase$(Trim$(domain))
    assemblyName = Trim$(assemblyName)

    Set sectionFinder = New clsSectionFinder
    Set sectionService = New clsSectionService
    Set componentFinder = New clsComponentListFinder
    Set generatedAssemblyService = New clsGeneratedAssemblyService

    If assemblyName = vbNullString Then
        errorMessage = "Assembly name cannot be empty."
        InsertCustomAssemblyWithComponents = False
        Exit Function
    End If

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, sectionName)

    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        InsertCustomAssemblyWithComponents = False
        Exit Function
    End If

    Set children = New Collection

    For i = LBound(selectedComponents) To UBound(selectedComponents)
        If Not componentFinder.GetComponentData(domain, selectedComponents(i), componentData, errorMessage) Then
            InsertCustomAssemblyWithComponents = False
            Exit Function
        End If

        children.Add generatedAssemblyService.CreateChild( _
            CStr(componentData("name")), _
            1, _
            1, _
            CDbl(componentData("des")), _
            CDbl(componentData("constr")), _
            CDbl(componentData("debug")), _
            CDbl(componentData("inst")), _
            CDbl(componentData("start")), _
            CDbl(componentData("mat")))
    Next i

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    If Not generatedAssemblyService.InsertGeneratedAssembly( _
        wsTarget, _
        insertRow, _
        assemblyName, _
        1, _
        1, _
        children, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow, _
        errorMessage) Then

        InsertCustomAssemblyWithComponents = False
        Exit Function
    End If

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertCustomAssemblyWithComponents = False
        Exit Function
    End If

    InsertCustomAssemblyWithComponents = True
End Function
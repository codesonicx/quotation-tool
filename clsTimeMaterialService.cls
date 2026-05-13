Option Explicit

' =========================================================
' clsTimeMaterialService
'
' Purpose:
' Defines and inserts the T&M section.
'
' This service only defines WHAT to insert.
' clsGeneratedAssemblyService handles HOW to insert it.
' =========================================================

Private Const SECTION_BASE_NAME As String = "T&M"

Private Const ASSEMBLY_TRAVEL As String = "Travel and Expense"
Private Const ASSEMBLY_LABOR As String = "Labor"

Private Const ITEM_PER_DIEM As String = "Per Diem"
Private Const ITEM_HOTEL As String = "Hotel"
Private Const ITEM_CAR_RENT As String = "Car Rent"
Private Const ITEM_FUEL As String = "Fuel"
Private Const ITEM_MILEAGE As String = "Mileage"
Private Const ITEM_AIR As String = "Air"

Private Const ITEM_STANDARD_HOURS As String = "Standard Hours"
Private Const ITEM_OVERTIME_HOURS As String = "Overtime Hours"

Public Function InsertTimeMaterialSection(ByVal wsTarget As Worksheet, _
                                          ByVal data As Collection, _
                                          ByRef errorMessage As String) As Boolean
    Dim sectionService As clsSectionService
    Dim sectionFinder As clsSectionFinder
    Dim generatedAssemblyService As clsGeneratedAssemblyService

    Dim sectionName As String
    Dim sectionRow As Long
    Dim insertRow As Long

    Dim travelChildren As Collection
    Dim laborChildren As Collection

    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString

    Set sectionService = New clsSectionService
    Set sectionFinder = New clsSectionFinder
    Set generatedAssemblyService = New clsGeneratedAssemblyService

    sectionName = sectionService.GetUniqueSectionName(wsTarget, SECTION_BASE_NAME)

    If Not sectionService.AddSection(wsTarget, sectionName, errorMessage) Then
        InsertTimeMaterialSection = False
        Exit Function
    End If

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Could not locate created section '" & sectionName & "'."
        InsertTimeMaterialSection = False
        Exit Function
    End If

    Set travelChildren = BuildTravelChildren(generatedAssemblyService, data)

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    If Not generatedAssemblyService.InsertGeneratedAssembly( _
        wsTarget, _
        insertRow, _
        ASSEMBLY_TRAVEL, _
        1, _
        1, _
        travelChildren, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow, _
        errorMessage) Then

        InsertTimeMaterialSection = False
        Exit Function
    End If

    Set laborChildren = BuildLaborChildren(generatedAssemblyService, data)

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    If Not generatedAssemblyService.InsertGeneratedAssembly( _
        wsTarget, _
        insertRow, _
        ASSEMBLY_LABOR, _
        1, _
        1, _
        laborChildren, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow, _
        errorMessage) Then

        InsertTimeMaterialSection = False
        Exit Function
    End If

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertTimeMaterialSection = False
        Exit Function
    End If

    InsertTimeMaterialSection = True
End Function

Private Function BuildTravelChildren(ByVal generatedAssemblyService As clsGeneratedAssemblyService, _
                                     ByVal data As Collection) As Collection
    Dim children As Collection

    Set children = New Collection

    children.Add generatedAssemblyService.CreateMaterialOnlyChild( _
        ITEM_PER_DIEM, _
        CDbl(data("perDiem")))

    children.Add generatedAssemblyService.CreateMaterialOnlyChild( _
        ITEM_HOTEL, _
        CDbl(data("hotelRate")))

    children.Add generatedAssemblyService.CreateMaterialOnlyChild( _
        ITEM_CAR_RENT, _
        CDbl(data("carRentRate")))

    children.Add generatedAssemblyService.CreateMaterialOnlyChild( _
        ITEM_FUEL, _
        CDbl(data("fuelRate")))

    children.Add generatedAssemblyService.CreateMaterialOnlyChild( _
        ITEM_MILEAGE, _
        CDbl(data("mileageRate")))

    children.Add generatedAssemblyService.CreateMaterialOnlyChild( _
        ITEM_AIR, _
        CDbl(data("flightRate")))

    Set BuildTravelChildren = children
End Function

Private Function BuildLaborChildren(ByVal generatedAssemblyService As clsGeneratedAssemblyService, _
                                    ByVal data As Collection) As Collection
    Dim children As Collection

    Set children = New Collection

    children.Add generatedAssemblyService.CreateStartupOnlyChild( _
        ITEM_STANDARD_HOURS, _
        CDbl(data("standardHours")), _
        CDbl(data("standardRate")))

    children.Add generatedAssemblyService.CreateStartupOnlyChild( _
        ITEM_OVERTIME_HOURS, _
        CDbl(data("overtimeHours")), _
        CDbl(data("overtimeRate")))

    Set BuildLaborChildren = children
End Function
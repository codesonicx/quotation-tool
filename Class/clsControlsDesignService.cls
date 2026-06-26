Option Explicit

' =========================================================
' clsControlsDesignService
'
' Purpose:
' Defines and inserts the fixed Controls Design Hours assembly.
'
' This service only defines WHAT to insert.
' clsGeneratedAssemblyService handles HOW to insert it.
' =========================================================

Private Const ASSEMBLY_NAME As String = "CONTROLS DESIGN HOURS"

Private Const ITEM_HARDWARE As String = "Hardware Hours"
Private Const ITEM_PLC As String = "PLC Hours"
Private Const ITEM_HMI As String = "HMI Hours"
Private Const ITEM_SCADA As String = "Scada Hours"

Public Function InsertControlsDesignHoursIntoSection(ByVal wsTarget As Worksheet, _
                                                     ByVal sectionName As String, _
                                                     ByVal hardwareHours As Double, _
                                                     ByVal plcHours As Double, _
                                                     ByVal hmiHours As Double, _
                                                     ByVal scadaHours As Double, _
                                                     ByRef errorMessage As String) As Boolean
    Dim sectionFinder As clsSectionFinder
    Dim sectionService As clsSectionService
    Dim generatedAssemblyService As clsGeneratedAssemblyService

    Dim sectionRow As Long
    Dim insertRow As Long

    Dim children As Collection

    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString

    Set sectionFinder = New clsSectionFinder
    Set sectionService = New clsSectionService
    Set generatedAssemblyService = New clsGeneratedAssemblyService

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, sectionName)
    If sectionRow = 0 Then
        errorMessage = "Section '" & sectionName & "' was not found."
        InsertControlsDesignHoursIntoSection = False
        Exit Function
    End If

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    Set children = BuildControlsDesignChildren( _
        generatedAssemblyService, _
        hardwareHours, _
        plcHours, _
        hmiHours, _
        scadaHours)

    If Not generatedAssemblyService.InsertGeneratedAssembly( _
        wsTarget, _
        insertRow, _
        ASSEMBLY_NAME, _
        1, _
        1, _
        children, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow, _
        errorMessage) Then

        InsertControlsDesignHoursIntoSection = False
        Exit Function
    End If

    If Not sectionService.RecalculateSectionTotals(wsTarget, sectionName, errorMessage) Then
        InsertControlsDesignHoursIntoSection = False
        Exit Function
    End If

    InsertControlsDesignHoursIntoSection = True
End Function

Private Function BuildControlsDesignChildren(ByVal generatedAssemblyService As clsGeneratedAssemblyService, _
                                             ByVal hardwareHours As Double, _
                                             ByVal plcHours As Double, _
                                             ByVal hmiHours As Double, _
                                             ByVal scadaHours As Double) As Collection
    Dim children As Collection

    Set children = New Collection

    children.Add generatedAssemblyService.CreateDesignOnlyChild(ITEM_HARDWARE, hardwareHours)
    children.Add generatedAssemblyService.CreateDesignOnlyChild(ITEM_PLC, plcHours)
    children.Add generatedAssemblyService.CreateDesignOnlyChild(ITEM_HMI, hmiHours)
    children.Add generatedAssemblyService.CreateDesignOnlyChild(ITEM_SCADA, scadaHours)

    Set BuildControlsDesignChildren = children
End Function
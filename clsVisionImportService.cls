Option Explicit

' =========================================================
' clsVisionImportService
'
' Purpose:
' Imports Vision data from another open workbook and creates
' a Vision section with a Vision assembly in the active estimate.
'
' Source workbook requirement:
' - Must contain worksheet: HW&Labor
' - Must contain label: Hardware with Markup
' - Must contain label: Vision Eng Rate
'
' Notes:
' - This first version preserves the main data extraction logic
'   from the legacy DanTransfer / SheetFill workflow.
' - Special legacy driver-rate replacements are not included yet.
'   We can add them after testing with real source workbooks.
' =========================================================

Private Const SOURCE_SHEET_HW_LABOR As String = "HW&Labor"
Private Const LABEL_HARDWARE_WITH_MARKUP As String = "Hardware with Markup"
Private Const LABEL_VISION_ENG_RATE As String = "Vision Eng Rate"

Private Const VISION_SECTION_BASE_NAME As String = "Vision"
Private Const VISION_ASSEMBLY_NAME As String = "Vision"

Public Function ImportVisionWorkbook(ByVal wsTarget As Worksheet, _
                                     ByVal sourceWorkbookName As String, _
                                     ByRef createdSectionName As String, _
                                     ByRef errorMessage As String) As Boolean
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet

    Dim sectionService As clsSectionService
    Dim sectionFinder As clsSectionFinder
    Dim generatedAssemblyService As clsGeneratedAssemblyService

    Dim visionData As Collection
    Dim titles As Collection
    Dim children As Collection

    Dim sectionRow As Long
    Dim insertRow As Long

    Dim assemblyRow As Long
    Dim firstChildRow As Long
    Dim lastChildRow As Long

    errorMessage = vbNullString
    createdSectionName = vbNullString

    If Trim$(sourceWorkbookName) = vbNullString Then
        errorMessage = "Source workbook name cannot be empty."
        ImportVisionWorkbook = False
        Exit Function
    End If

    If Not TryGetOpenWorkbook(sourceWorkbookName, sourceWorkbook) Then
        errorMessage = "Workbook '" & sourceWorkbookName & "' is not open."
        ImportVisionWorkbook = False
        Exit Function
    End If

    If Not TryGetWorksheet(sourceWorkbook, SOURCE_SHEET_HW_LABOR, sourceSheet) Then
        errorMessage = "Worksheet '" & SOURCE_SHEET_HW_LABOR & "' was not found in '" & sourceWorkbookName & "'."
        ImportVisionWorkbook = False
        Exit Function
    End If

    If Not ExtractVisionData(sourceSheet, visionData, titles, errorMessage) Then
        ImportVisionWorkbook = False
        Exit Function
    End If

    Set sectionService = New clsSectionService
    Set sectionFinder = New clsSectionFinder
    Set generatedAssemblyService = New clsGeneratedAssemblyService

    createdSectionName = sectionService.GetUniqueSectionName(wsTarget, VISION_SECTION_BASE_NAME)

    If Not sectionService.AddSection(wsTarget, createdSectionName, errorMessage) Then
        ImportVisionWorkbook = False
        Exit Function
    End If

    sectionRow = sectionFinder.FindSectionHeaderRow(wsTarget, createdSectionName)

    If sectionRow = 0 Then
        errorMessage = "Could not locate created Vision section '" & createdSectionName & "'."
        ImportVisionWorkbook = False
        Exit Function
    End If

    Set children = BuildVisionChildren(generatedAssemblyService, visionData, titles)

    insertRow = sectionFinder.FindNextSectionHeaderRowAfter(wsTarget, sectionRow)

    If Not generatedAssemblyService.InsertGeneratedAssembly( _
        wsTarget, _
        insertRow, _
        VISION_ASSEMBLY_NAME, _
        1, _
        1, _
        children, _
        assemblyRow, _
        firstChildRow, _
        lastChildRow, _
        errorMessage) Then

        ImportVisionWorkbook = False
        Exit Function
    End If

    If Not sectionService.RecalculateSectionTotals(wsTarget, createdSectionName, errorMessage) Then
        ImportVisionWorkbook = False
        Exit Function
    End If

    ImportVisionWorkbook = True
End Function

Private Function ExtractVisionData(ByVal sourceSheet As Worksheet, _
                                   ByRef visionData As Collection, _
                                   ByRef titles As Collection, _
                                   ByRef errorMessage As String) As Boolean
    Dim hardwareCell As Range
    Dim rateCell As Range
    Dim titleRange As Range
    Dim titleCell As Range

    Dim rate1 As Double
    Dim rate2 As Double
    Dim rate3 As Double
    Dim rate4 As Double

    errorMessage = vbNullString

    Set hardwareCell = FindRequiredCell(sourceSheet, LABEL_HARDWARE_WITH_MARKUP)

    If hardwareCell Is Nothing Then
        errorMessage = "Label '" & LABEL_HARDWARE_WITH_MARKUP & "' was not found in '" & sourceSheet.Name & "'."
        ExtractVisionData = False
        Exit Function
    End If

    Set rateCell = FindRequiredCell(sourceSheet, LABEL_VISION_ENG_RATE)

    If rateCell Is Nothing Then
        errorMessage = "Label '" & LABEL_VISION_ENG_RATE & "' was not found in '" & sourceSheet.Name & "'."
        ExtractVisionData = False
        Exit Function
    End If

    rate1 = ToNumber(rateCell.Offset(0, 2).Value)
    rate2 = ToNumber(rateCell.Offset(1, 2).Value)
    rate3 = ToNumber(rateCell.Offset(2, 2).Value)
    rate4 = ToNumber(rateCell.Offset(3, 2).Value)

    If rate1 = 0 Or rate2 = 0 Or rate3 = 0 Or rate4 = 0 Then
        errorMessage = "One or more Vision Eng Rate values are zero or invalid."
        ExtractVisionData = False
        Exit Function
    End If

    Set visionData = New Collection
    Set titles = New Collection

    visionData.Add ToNumber(hardwareCell.Offset(0, 3).Value), "Hardware"
    visionData.Add ToNumber(hardwareCell.Offset(1, 3).Value) / rate1, "VCHours"
    visionData.Add ToNumber(hardwareCell.Offset(2, 3).Value) / rate1, "VIHours"
    visionData.Add ToNumber(hardwareCell.Offset(3, 3).Value) / rate2, "CCHours"
    visionData.Add ToNumber(hardwareCell.Offset(4, 3).Value) / rate2, "CIHours"
    visionData.Add ToNumber(hardwareCell.Offset(5, 3).Value) / rate3, "PMHours"
    visionData.Add ToNumber(hardwareCell.Offset(6, 3).Value) / rate4, "SLHours"
    visionData.Add ToNumber(hardwareCell.Offset(7, 3).Value), "Travel"
    visionData.Add ToNumber(hardwareCell.Offset(8, 3).Value) / rate1, "FMEHours"
    visionData.Add ToNumber(hardwareCell.Offset(9, 3).Value), "Contingency"

    Set titleRange = sourceSheet.Range(hardwareCell, hardwareCell.End(xlDown))

    For Each titleCell In titleRange.Cells
        If Trim$(CStr(titleCell.Value)) <> vbNullString Then
            titles.Add CStr(titleCell.Value)
        End If
    Next titleCell

    ExtractVisionData = True
End Function

Private Function BuildVisionChildren(ByVal generatedAssemblyService As clsGeneratedAssemblyService, _
                                     ByVal visionData As Collection, _
                                     ByVal titles As Collection) As Collection
    Dim children As Collection

    Set children = New Collection

    ' Hardware is always added, matching the legacy workflow.
    children.Add generatedAssemblyService.CreateChild( _
        GetTitleOrDefault(titles, 1, "Hardware"), _
        1, _
        1, _
        0, _
        0, _
        0, _
        0, _
        0, _
        CDbl(visionData("Hardware")))

    If CDbl(visionData("VCHours")) > 0 Or CDbl(visionData("VIHours")) > 0 Then
        children.Add generatedAssemblyService.CreateChild( _
            "Vision Labor", _
            1, _
            1, _
            0, _
            CDbl(visionData("VCHours")), _
            0, _
            CDbl(visionData("VIHours")), _
            0, _
            0)
    End If

    If CDbl(visionData("CCHours")) > 0 Or CDbl(visionData("CIHours")) > 0 Then
        children.Add generatedAssemblyService.CreateChild( _
            "Controls Labor", _
            1, _
            1, _
            0, _
            CDbl(visionData("CCHours")), _
            0, _
            CDbl(visionData("CIHours")), _
            0, _
            0)
    End If

    If CDbl(visionData("PMHours")) > 0 Then
        children.Add generatedAssemblyService.CreateChild( _
            GetTitleOrDefault(titles, 6, "Project Management"), _
            1, _
            1, _
            0, _
            0, _
            0, _
            CDbl(visionData("PMHours")), _
            0, _
            0)
    End If

    If CDbl(visionData("SLHours")) > 0 Then
        children.Add generatedAssemblyService.CreateChild( _
            GetTitleOrDefault(titles, 7, "Site Lead"), _
            1, _
            1, _
            0, _
            CDbl(visionData("SLHours")), _
            0, _
            0, _
            0, _
            0)
    End If

    If CDbl(visionData("Travel")) > 0 Then
        children.Add generatedAssemblyService.CreateChild( _
            GetTitleOrDefault(titles, 8, "Travel"), _
            1, _
            1, _
            0, _
            0, _
            0, _
            0, _
            0, _
            CDbl(visionData("Travel")))
    End If

    If CDbl(visionData("FMEHours")) > 0 Then
        children.Add generatedAssemblyService.CreateChild( _
            GetTitleOrDefault(titles, 9, "FME"), _
            1, _
            1, _
            0, _
            0, _
            0, _
            CDbl(visionData("FMEHours")), _
            0, _
            0)
    End If

    If CDbl(visionData("Contingency")) > 0 Then
        children.Add generatedAssemblyService.CreateChild( _
            GetTitleOrDefault(titles, 10, "Contingency"), _
            1, _
            1, _
            0, _
            0, _
            0, _
            0, _
            0, _
            CDbl(visionData("Contingency")))
    End If

    Set BuildVisionChildren = children
End Function

Private Function FindRequiredCell(ByVal ws As Worksheet, _
                                  ByVal searchText As String) As Range
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find( _
        What:=searchText, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        MatchCase:=False)

    If Not foundCell Is Nothing Then
        Set FindRequiredCell = foundCell
    End If
End Function

Private Function TryGetOpenWorkbook(ByVal workbookName As String, _
                                    ByRef targetWorkbook As Workbook) As Boolean
    On Error GoTo CleanFail

    Set targetWorkbook = Application.Workbooks(workbookName)
    TryGetOpenWorkbook = True
    Exit Function

CleanFail:
    Set targetWorkbook = Nothing
    TryGetOpenWorkbook = False
End Function

Private Function TryGetWorksheet(ByVal sourceWorkbook As Workbook, _
                                 ByVal sheetName As String, _
                                 ByRef targetSheet As Worksheet) As Boolean
    On Error GoTo CleanFail

    Set targetSheet = sourceWorkbook.Worksheets(sheetName)
    TryGetWorksheet = True
    Exit Function

CleanFail:
    Set targetSheet = Nothing
    TryGetWorksheet = False
End Function

Private Function GetTitleOrDefault(ByVal titles As Collection, _
                                   ByVal index As Long, _
                                   ByVal defaultText As String) As String
    On Error GoTo UseDefault

    If index <= 0 Or index > titles.Count Then
        GetTitleOrDefault = defaultText
    ElseIf Trim$(CStr(titles(index))) = vbNullString Then
        GetTitleOrDefault = defaultText
    Else
        GetTitleOrDefault = CStr(titles(index))
    End If

    Exit Function

UseDefault:
    GetTitleOrDefault = defaultText
End Function

Private Function ToNumber(ByVal value As Variant) As Double
    If IsError(value) Then
        ToNumber = 0
    ElseIf Trim$(CStr(value)) = vbNullString Then
        ToNumber = 0
    Else
        ToNumber = CDbl(Val(value))
    End If
End Function
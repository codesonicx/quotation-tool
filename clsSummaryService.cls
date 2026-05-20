Option Explicit

' =========================================================
' clsSummaryService
'
' Purpose:
' Refreshes the Summary sheet by scanning all non-system
' estimate tabs and summarizing ME / CE hours and totals.
'
' Rules:
' - Ignore fixed workbook tabs.
' - Try to process every other sheet as an estimate tab.
' - A valid estimate tab must have the required estimate headers.
' - ME rows are rows where Component:
'       1) exists in ME Component List, or
'       2) contains [ME]
' - CE rows are rows where Component:
'       1) exists in CE Component List, or
'       2) contains [CE]
' - Unknown custom rows without [ME] or [CE] are ignored.
' =========================================================

Public Function RefreshSummary(ByVal wb As Workbook, _
                               ByRef errorMessage As String) As Boolean
    Dim summarySheet As Worksheet
    Dim summaryTable As ListObject

    Dim meComponents As Object
    Dim ceComponents As Object

    Dim ws As Worksheet
    Dim outputRow As Long
    Dim validSheetCount As Long

    Dim meDesign As Double
    Dim meConstruction As Double
    Dim meInstallation As Double

    Dim ceDesign As Double
    Dim ceConstruction As Double
    Dim ceInstallation As Double

    Dim budgetedTotal As Double
    Dim totalWithMarkup As Double

    errorMessage = vbNullString

    If wb Is Nothing Then
        errorMessage = "Workbook is not available."
        RefreshSummary = False
        Exit Function
    End If

    If Not TryGetWorksheet(wb, SHEET_SUMMARY, summarySheet) Then
        errorMessage = "Sheet '" & SHEET_SUMMARY & "' was not found."
        RefreshSummary = False
        Exit Function
    End If

    If Not TryGetSummaryTable(summarySheet, summaryTable, errorMessage) Then
        RefreshSummary = False
        Exit Function
    End If

    Set meComponents = CreateObject("Scripting.Dictionary")
    Set ceComponents = CreateObject("Scripting.Dictionary")

    meComponents.CompareMode = vbTextCompare
    ceComponents.CompareMode = vbTextCompare

    If Not LoadComponentDictionary(wb, SHEET_ME_COMPONENT_LIST, meComponents, errorMessage) Then
        RefreshSummary = False
        Exit Function
    End If

    If Not LoadComponentDictionary(wb, SHEET_CE_COMPONENT_LIST, ceComponents, errorMessage) Then
        RefreshSummary = False
        Exit Function
    End If

    ' Resize large enough first so old rows are cleared.
    ResizeSummaryTable summaryTable, wb.Worksheets.Count + 1
    ClearSummaryTable summaryTable

    outputRow = summaryTable.HeaderRowRange.Row + 1
    validSheetCount = 0

    For Each ws In wb.Worksheets
        If IsCandidateEstimateWorksheet(ws) Then

            CalculateWorksheetSummary _
                ws, _
                meComponents, _
                ceComponents, _
                meDesign, _
                meConstruction, _
                meInstallation, _
                ceDesign, _
                ceConstruction, _
                ceInstallation, _
                budgetedTotal, _
                totalWithMarkup

            If HasValidSummaryData( _
                meDesign, _
                meConstruction, _
                meInstallation, _
                ceDesign, _
                ceConstruction, _
                ceInstallation, _
                budgetedTotal, _
                totalWithMarkup) Then

                WriteSummaryRow _
                    summarySheet, _
                    outputRow, _
                    ws.Name, _
                    meDesign, _
                    meConstruction, _
                    meInstallation, _
                    ceDesign, _
                    ceConstruction, _
                    ceInstallation, _
                    budgetedTotal, _
                    totalWithMarkup

                outputRow = outputRow + 1
                validSheetCount = validSheetCount + 1
            End If
        End If
    Next ws

    WriteGrandTotalRow summarySheet, outputRow, summaryTable

    ' Keep only valid rows plus the Grand Total row.
    ResizeSummaryTable summaryTable, validSheetCount + 1
    ApplySummaryNumberFormats summaryTable

    RefreshSummary = True
End Function

' =========================================================
' Worksheet calculation
' =========================================================

Private Sub CalculateWorksheetSummary(ByVal ws As Worksheet, _
                                      ByVal meComponents As Object, _
                                      ByVal ceComponents As Object, _
                                      ByRef meDesign As Double, _
                                      ByRef meConstruction As Double, _
                                      ByRef meInstallation As Double, _
                                      ByRef ceDesign As Double, _
                                      ByRef ceConstruction As Double, _
                                      ByRef ceInstallation As Double, _
                                      ByRef budgetedTotal As Double, _
                                      ByRef totalWithMarkup As Double)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim componentName As String

    meDesign = 0
    meConstruction = 0
    meInstallation = 0

    ceDesign = 0
    ceConstruction = 0
    ceInstallation = 0

    budgetedTotal = 0
    totalWithMarkup = 0

    lastRow = modUtils.GetLastUsedRow(ws)

    For rowIndex = FIRST_DATA_ROW To lastRow
        componentName = Trim$(CStr(ws.Cells(rowIndex, COL_COMPONENT).Value))

        If componentName <> vbNullString Then
            If IsMEComponent(componentName, meComponents) Then
                meDesign = meDesign + ToNumber(ws.Cells(rowIndex, COL_DESIGN_HOURS).Value)
                meConstruction = meConstruction + ToNumber(ws.Cells(rowIndex, COL_CONSTRUCTION_HOURS).Value)
                meInstallation = meInstallation + ToNumber(ws.Cells(rowIndex, COL_INSTALLATION_HOURS).Value)

            ElseIf IsCEComponent(componentName, ceComponents) Then
                ceDesign = ceDesign + ToNumber(ws.Cells(rowIndex, COL_DESIGN_HOURS).Value)
                ceConstruction = ceConstruction + ToNumber(ws.Cells(rowIndex, COL_CONSTRUCTION_HOURS).Value)
                ceInstallation = ceInstallation + ToNumber(ws.Cells(rowIndex, COL_INSTALLATION_HOURS).Value)
            End If
        End If
    Next rowIndex

    budgetedTotal = ToNumber(ws.Cells(GRAND_TOTALS_ROW, COL_BUDGET_TOTAL).Value)
    totalWithMarkup = ToNumber(ws.Cells(GRAND_TOTALS_ROW, COL_MARKUP_TOTAL).Value)
End Sub

Private Function HasValidSummaryData(ByVal meDesign As Double, _
                                     ByVal meConstruction As Double, _
                                     ByVal meInstallation As Double, _
                                     ByVal ceDesign As Double, _
                                     ByVal ceConstruction As Double, _
                                     ByVal ceInstallation As Double, _
                                     ByVal budgetedTotal As Double, _
                                     ByVal totalWithMarkup As Double) As Boolean

    HasValidSummaryData = _
        (meDesign <> 0) _
        Or (meConstruction <> 0) _
        Or (meInstallation <> 0) _
        Or (ceDesign <> 0) _
        Or (ceConstruction <> 0) _
        Or (ceInstallation <> 0) _
        Or (budgetedTotal <> 0) _
        Or (totalWithMarkup <> 0)

End Function

' =========================================================
' Summary output
' =========================================================

Private Sub WriteSummaryRow(ByVal wsSummary As Worksheet, _
                            ByVal rowIndex As Long, _
                            ByVal tabName As String, _
                            ByVal meDesign As Double, _
                            ByVal meConstruction As Double, _
                            ByVal meInstallation As Double, _
                            ByVal ceDesign As Double, _
                            ByVal ceConstruction As Double, _
                            ByVal ceInstallation As Double, _
                            ByVal budgetedTotal As Double, _
                            ByVal totalWithMarkup As Double)

    wsSummary.Cells(rowIndex, 1).Value = tabName

    wsSummary.Cells(rowIndex, 2).Value = meDesign
    wsSummary.Cells(rowIndex, 3).Value = meConstruction
    wsSummary.Cells(rowIndex, 4).Value = meInstallation

    ' ME Total Hours = B + C + D from the same row.
    wsSummary.Cells(rowIndex, 5).FormulaR1C1 = "=SUM(RC[-3],RC[-2],RC[-1])"

    wsSummary.Cells(rowIndex, 6).Value = ceDesign
    wsSummary.Cells(rowIndex, 7).Value = ceConstruction
    wsSummary.Cells(rowIndex, 8).Value = ceInstallation

    ' CE Total Hours = F + G + H from the same row.
    wsSummary.Cells(rowIndex, 9).FormulaR1C1 = "=SUM(RC[-3],RC[-2],RC[-1])"

    wsSummary.Cells(rowIndex, 10).Value = budgetedTotal
    wsSummary.Cells(rowIndex, 11).Value = totalWithMarkup

End Sub

Private Sub WriteGrandTotalRow(ByVal wsSummary As Worksheet, _
                               ByVal rowIndex As Long, _
                               ByVal summaryTable As ListObject)
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim colIndex As Long

    firstDataRow = summaryTable.HeaderRowRange.Row + 1
    lastDataRow = rowIndex - 1

    wsSummary.Cells(rowIndex, 1).Value = "Grand Total"

    If lastDataRow < firstDataRow Then
        For colIndex = 2 To 11
            wsSummary.Cells(rowIndex, colIndex).Value = 0
        Next colIndex

        Exit Sub
    End If

    For colIndex = 2 To 11
        wsSummary.Cells(rowIndex, colIndex).Formula = "=SUM(" & _
            wsSummary.Cells(firstDataRow, colIndex).Address(False, False) & ":" & _
            wsSummary.Cells(lastDataRow, colIndex).Address(False, False) & ")"
    Next colIndex

End Sub

' =========================================================
' Estimate sheet detection
' =========================================================

Private Function IsCandidateEstimateWorksheet(ByVal ws As Worksheet) As Boolean

    If IsFixedWorkbookSheet(ws.Name) Then
        IsCandidateEstimateWorksheet = False
        Exit Function
    End If

    If Not WorksheetHasRequiredEstimateHeaders(ws) Then
        IsCandidateEstimateWorksheet = False
        Exit Function
    End If

    IsCandidateEstimateWorksheet = True

End Function

Private Function IsFixedWorkbookSheet(ByVal sheetName As String) As Boolean
    Select Case UCase$(Trim$(sheetName))
        Case UCase$(SHEET_COVER_PAGE), _
             UCase$(SHEET_SUMMARY), _
             UCase$(SHEET_SETUP), _
             UCase$(SHEET_UPDATING_SHEET), _
             UCase$(SHEET_ASSEMBLY_LIST), _
             UCase$(SHEET_ME_COMPONENT_LIST), _
             UCase$(SHEET_CE_COMPONENT_LIST), _
             UCase$(SHEET_DRIVERS), _
             UCase$(SHEET_TEMPLATE), _
             UCase$(SHEET_NOTES)

            IsFixedWorkbookSheet = True

        Case Else
            IsFixedWorkbookSheet = False
    End Select
End Function

Private Function WorksheetHasRequiredEstimateHeaders(ByVal ws As Worksheet) As Boolean
    WorksheetHasRequiredEstimateHeaders = _
        HeaderExists(ws, HEADER_COMPONENT) _
        And HeaderExists(ws, HEADER_DESIGN_HOURS) _
        And HeaderExists(ws, HEADER_CONSTRUCTION_HOURS) _
        And HeaderExists(ws, HEADER_INSTALLATION_HOURS) _
        And HeaderExists(ws, HEADER_BUDGETED_TOTAL) _
        And HeaderExists(ws, HEADER_TOTAL_WITH_MARKUP)
End Function

Private Function HeaderExists(ByVal ws As Worksheet, _
                              ByVal headerText As String) As Boolean
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find( _
        What:=headerText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    HeaderExists = Not foundCell Is Nothing
End Function

' =========================================================
' Component dictionary / classification
' =========================================================

Private Function LoadComponentDictionary(ByVal wb As Workbook, _
                                         ByVal sheetName As String, _
                                         ByVal targetDictionary As Object, _
                                         ByRef errorMessage As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim componentName As String
    Dim componentKey As String

    errorMessage = vbNullString

    If Not TryGetWorksheet(wb, sheetName, ws) Then
        errorMessage = "Sheet '" & sheetName & "' was not found."
        LoadComponentDictionary = False
        Exit Function
    End If

    lastRow = modUtils.GetLastUsedRow(ws)

    For rowIndex = COMPONENT_LIST_FIRST_DATA_ROW To lastRow
        componentName = CStr(ws.Cells(rowIndex, COL_COMPONENT_LIST_ITEM).Value)
        componentKey = MakeComponentKey(componentName)

        If componentKey <> vbNullString Then
            If Not targetDictionary.Exists(componentKey) Then
                targetDictionary.Add componentKey, True
            End If
        End If
    Next rowIndex

    LoadComponentDictionary = True
End Function

Private Function IsMEComponent(ByVal componentName As String, _
                               ByVal meComponents As Object) As Boolean
    If InStr(1, componentName, "[ME]", vbTextCompare) > 0 Then
        IsMEComponent = True
    Else
        IsMEComponent = meComponents.Exists(MakeComponentKey(componentName))
    End If
End Function

Private Function IsCEComponent(ByVal componentName As String, _
                               ByVal ceComponents As Object) As Boolean
    If InStr(1, componentName, "[CE]", vbTextCompare) > 0 Then
        IsCEComponent = True
    Else
        IsCEComponent = ceComponents.Exists(MakeComponentKey(componentName))
    End If
End Function

Private Function MakeComponentKey(ByVal componentName As String) As String
    MakeComponentKey = LCase$(Trim$(componentName))
End Function

' =========================================================
' Table helpers
' =========================================================

Private Sub ClearSummaryTable(ByVal summaryTable As ListObject)
    If Not summaryTable.DataBodyRange Is Nothing Then
        summaryTable.DataBodyRange.ClearContents
    End If
End Sub

Private Sub ResizeSummaryTable(ByVal summaryTable As ListObject, _
                               ByVal dataRowCount As Long)
    Dim topLeftCell As Range
    Dim bottomRightCell As Range
    Dim tableColumnCount As Long

    If dataRowCount < 1 Then dataRowCount = 1

    tableColumnCount = summaryTable.ListColumns.Count

    Set topLeftCell = summaryTable.HeaderRowRange.Cells(1, 1)
    Set bottomRightCell = topLeftCell.Offset(dataRowCount, tableColumnCount - 1)

    summaryTable.Resize topLeftCell.Worksheet.Range(topLeftCell, bottomRightCell)
End Sub

Private Sub ApplySummaryNumberFormats(ByVal summaryTable As ListObject)
    If summaryTable.DataBodyRange Is Nothing Then Exit Sub

    summaryTable.DataBodyRange.Columns(2).Resize(, 8).NumberFormat = FORMAT_DECIMAL
    summaryTable.DataBodyRange.Columns(10).Resize(, 2).NumberFormat = FORMAT_CURRENCY
End Sub

Private Function TryGetSummaryTable(ByVal summarySheet As Worksheet, _
                                    ByRef summaryTable As ListObject, _
                                    ByRef errorMessage As String) As Boolean
    On Error GoTo CleanFail

    Set summaryTable = summarySheet.ListObjects(TABLE_SUMMARY)
    TryGetSummaryTable = True
    Exit Function

CleanFail:
    errorMessage = "Table '" & TABLE_SUMMARY & "' was not found on sheet '" & SHEET_SUMMARY & "'."
    Set summaryTable = Nothing
    TryGetSummaryTable = False
End Function

Private Function TryGetWorksheet(ByVal wb As Workbook, _
                                 ByVal sheetName As String, _
                                 ByRef ws As Worksheet) As Boolean
    On Error GoTo CleanFail

    Set ws = wb.Worksheets(sheetName)
    TryGetWorksheet = True
    Exit Function

CleanFail:
    Set ws = Nothing
    TryGetWorksheet = False
End Function

Private Function ToNumber(ByVal value As Variant) As Double
    Dim cleanText As String

    If IsError(value) Then
        ToNumber = 0
    ElseIf IsNumeric(value) Then
        ToNumber = CDbl(value)
    Else
        cleanText = Trim$(CStr(value))
        cleanText = Replace(cleanText, "$", "")
        cleanText = Replace(cleanText, ",", "")

        If cleanText = vbNullString Then
            ToNumber = 0
        ElseIf IsNumeric(cleanText) Then
            ToNumber = CDbl(cleanText)
        Else
            ToNumber = 0
        End If
    End If
End Function
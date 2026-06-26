Option Explicit

' =========================================================
' clsComponentListFinder
'
' Purpose:
' Reads component names and component values from:
' - ME Component List
' - CE Component List
'
' Uses fixed columns because both sheets have the same structure.
' =========================================================

Public Function GetComponentNames(ByVal domain As String, _
                                  ByRef items() As String, _
                                  ByRef errorMessage As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim count As Long
    Dim itemName As String

    errorMessage = vbNullString
    domain = UCase$(Trim$(domain))

    Set ws = GetComponentListWorksheet(domain, errorMessage)

    If ws Is Nothing Then
        GetComponentNames = False
        Exit Function
    End If

    lastRow = modUtils.GetLastUsedRow(ws)

    If lastRow < COMPONENT_LIST_FIRST_DATA_ROW Then
        errorMessage = "No component items found in '" & ws.Name & "'."
        GetComponentNames = False
        Exit Function
    End If

    count = 0

    For rowIndex = COMPONENT_LIST_FIRST_DATA_ROW To lastRow
        itemName = Trim$(CStr(ws.Cells(rowIndex, COL_COMPONENT_LIST_ITEM).Value))

        If itemName <> vbNullString Then
            count = count + 1
            ReDim Preserve items(1 To count)
            items(count) = itemName
        End If
    Next rowIndex

    If count = 0 Then
        errorMessage = "No component items found in '" & ws.Name & "'."
        GetComponentNames = False
        Exit Function
    End If

    GetComponentNames = True
End Function

Public Function GetComponentData(ByVal domain As String, _
                                 ByVal componentName As String, _
                                 ByRef componentData As Collection, _
                                 ByRef errorMessage As String) As Boolean
    Dim ws As Worksheet
    Dim foundRow As Long

    errorMessage = vbNullString
    domain = UCase$(Trim$(domain))
    componentName = Trim$(componentName)

    Set ws = GetComponentListWorksheet(domain, errorMessage)

    If ws Is Nothing Then
        GetComponentData = False
        Exit Function
    End If

    foundRow = FindComponentRow(ws, componentName)

    If foundRow = 0 Then
        errorMessage = "Component '" & componentName & "' was not found in '" & ws.Name & "'."
        GetComponentData = False
        Exit Function
    End If

    Set componentData = New Collection

    componentData.Add CStr(ws.Cells(foundRow, COL_COMPONENT_LIST_ITEM).Value), "name"

    componentData.Add CDbl(Val(ws.Cells(foundRow, COL_COMPONENT_LIST_DESIGN).Value)), "des"
    componentData.Add CDbl(Val(ws.Cells(foundRow, COL_COMPONENT_LIST_CONSTRUCTION).Value)), "constr"
    componentData.Add CDbl(Val(ws.Cells(foundRow, COL_COMPONENT_LIST_DEBUG).Value)), "debug"
    componentData.Add CDbl(Val(ws.Cells(foundRow, COL_COMPONENT_LIST_INSTALLATION).Value)), "inst"
    componentData.Add CDbl(Val(ws.Cells(foundRow, COL_COMPONENT_LIST_STARTUP).Value)), "start"
    componentData.Add CDbl(Val(ws.Cells(foundRow, COL_COMPONENT_LIST_MATERIAL_RATE).Value)), "mat"

    GetComponentData = True
End Function

Private Function GetComponentListWorksheet(ByVal domain As String, _
                                           ByRef errorMessage As String) As Worksheet
    errorMessage = vbNullString

    Select Case UCase$(Trim$(domain))
        Case DOMAIN_ME
            Set GetComponentListWorksheet = ThisWorkbook.Worksheets(SHEET_ME_COMPONENT_LIST)

        Case DOMAIN_CE
            Set GetComponentListWorksheet = ThisWorkbook.Worksheets(SHEET_CE_COMPONENT_LIST)

        Case Else
            errorMessage = "Invalid component list domain: [" & domain & "]."
            Set GetComponentListWorksheet = Nothing
    End Select
End Function

Private Function FindComponentRow(ByVal ws As Worksheet, _
                                  ByVal componentName As String) As Long
    Dim foundCell As Range

    Set foundCell = ws.Columns(COL_COMPONENT_LIST_ITEM).Find( _
        What:=componentName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        FindComponentRow = 0
    Else
        FindComponentRow = foundCell.Row
    End If
End Function
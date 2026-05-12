Option Explicit

Public Function CreateNewSheet(ByVal sheetName As String, _
                               ByRef errorMessage As String) As Boolean

    Dim newSheet As Worksheet
    Dim previousSheet As Worksheet

    If Trim$(sheetName) = "" Then
        errorMessage = "Sheet name cannot be empty."
        CreateNewSheet = False
        Exit Function
    End If

    If SheetExists(sheetName) Then
        errorMessage = "Sheet already exists."
        CreateNewSheet = False
        Exit Function
    End If

    Set previousSheet = ActiveSheet

    ThisWorkbook.Sheets(SHEET_TEMPLATE).Copy _
        Before:=ThisWorkbook.Sheets(SHEET_ASSEMBLY_LIST)

    On Error GoTo Rollback

    Set newSheet = ActiveSheet

    newSheet.Name = sheetName
    newSheet.Tab.Color = RGB(0, 176, 240)

    CreateNewSheet = True
    Exit Function

Rollback:
    ' If rename fails, delete the copied sheet and return user to a safe place
    Application.DisplayAlerts = False
    newSheet.Delete
    Application.DisplayAlerts = True

    If Not previousSheet Is Nothing Then
        previousSheet.Activate
    End If

    errorMessage = "Invalid sheet name (too long, forbidden characters, or already used)."
    CreateNewSheet = False

End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws

    SheetExists = False

End Function
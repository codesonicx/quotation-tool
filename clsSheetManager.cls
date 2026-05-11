Option Explicit

Public Function CreateNewSheet(ByVal sheetName As String, ByRef errorMessage As String) As Boolean

    Dim NewSheet As Worksheet
    Dim prevSheet As Worksheet
    
    ' Validation
    If Trim$(sheetName) = "" Then
        errorMessage = "Sheet name cannot be empty."
        CreateNewSheet = False
        Exit Function
    End If
    
    ' Check if sheet already exists
    If SheetExists(sheetName) Then
        errorMessage = "Sheet already exists."
        CreateNewSheet = False
        Exit Function
    End If
    
    Set prevSheet = ActiveSheet
    
    ' Copy template
    ThisWorkbook.Sheets("Template").Copy _
        Before:=ThisWorkbook.Sheets("Assembly List")
        
    ' Try rename + formatting
    On Error GoTo Rollback
        
    Set NewSheet = ActiveSheet
    
    NewSheet.name = sheetName
    NewSheet.Tab.Color = RGB(0, 176, 240)
    
    CreateNewSheet = True
    Exit Function
    
Rollback:
    ' If rename fails, delete the copied sheet and return user to a safe place
    Application.DisplayAlerts = False
    NewSheet.Delete
    Application.DisplayAlerts = True

    If Not prevSheet Is Nothing Then prevSheet.Activate

    errorMessage = "Invalid sheet name (too long, forbidden characters, or already used)."
    CreateNewSheet = False

End Function


Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    
    SheetExists = False

End Function


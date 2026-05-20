Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 360
Private Const FORM_WIDTH As Long = 360
Private Const FORM_TITLE As String = "Vision Sheet Insert"

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    ConfigureForm
    LoadOpenWorkbooks
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub

' =========================================================
' Button events
' =========================================================

Private Sub btnSelect_Click()
    HandleSelectAction
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' =========================================================
' Main workflow
' =========================================================

Private Sub HandleSelectAction()
    Dim wsTarget As Worksheet
    Dim visionService As clsVisionImportService

    Dim selectedWorkbookName As String
    Dim createdSectionName As String

    Dim success As Boolean
    Dim errorMessage As String

    If Not IsWorkbookSelected Then Exit Sub

    selectedWorkbookName = GetSelectedWorkbookName()

    Set wsTarget = modUtils.GetTargetWorksheet()
    Set visionService = New clsVisionImportService

    success = visionService.ImportVisionWorkbook( _
        wsTarget, _
        selectedWorkbookName, _
        createdSectionName, _
        errorMessage)

    If success Then
        MsgBox "Vision section inserted successfully as '" & createdSectionName & "'.", vbInformation, FORM_TITLE
        Unload Me
    Else
        MsgBox errorMessage, vbExclamation, FORM_TITLE
    End If
End Sub

' =========================================================
' UI setup
' =========================================================

Private Sub ConfigureForm()
    Me.Height = FORM_HEIGHT
    Me.Width = FORM_WIDTH
End Sub

' =========================================================
' Data loading
' =========================================================

Private Sub LoadOpenWorkbooks()
    Dim wb As Workbook

    Me.lstWorkbooks.Clear

    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            Me.lstWorkbooks.AddItem wb.Name
        End If
    Next wb

    If Me.lstWorkbooks.ListCount = 0 Then
        MsgBox "No external workbooks are open. Please open the Vision source workbook first.", vbInformation, FORM_TITLE
    Else
        Me.lstWorkbooks.ListIndex = 0
    End If
End Sub

' =========================================================
' Validation helpers
' =========================================================

Private Function IsWorkbookSelected() As Boolean
    If Me.lstWorkbooks.ListIndex = -1 Then
        MsgBox "Please select a workbook.", vbExclamation, FORM_TITLE
        IsWorkbookSelected = False
    Else
        IsWorkbookSelected = True
    End If
End Function

Private Function GetSelectedWorkbookName() As String
    GetSelectedWorkbookName = CStr(Me.lstWorkbooks.List(Me.lstWorkbooks.ListIndex))
End Function
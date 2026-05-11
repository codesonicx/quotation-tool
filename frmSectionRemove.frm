Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 300
Private Const FORM_WIDTH As Long = 300
Private Const FORM_TITLE As String = "Remove Section"

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    ConfigureForm
    LoadSections
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub

' =========================================================
' Button events
' =========================================================

Private Sub btnRemove_Click()
    HandleRemoveAction
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' =========================================================
' Main workflow
' =========================================================

' Handles the full remove-section workflow.
Private Sub HandleRemoveAction()
    Dim ws As Worksheet
    Dim manager As clsSectionManager
    Dim selectedSection As String
    Dim errorMessage As String
    Dim success As Boolean

    If Not IsSectionSelected Then Exit Sub

    selectedSection = GetSelectedSectionName()

    If Not ConfirmSectionRemoval(selectedSection) Then Exit Sub

    Set ws = GetTargetWorksheet()
    Set manager = New clsSectionManager

    success = manager.RemoveSection(ws, selectedSection, errorMessage)

    If success Then
        MsgBox "Section removed successfully.", vbInformation, FORM_TITLE
        Unload Me
    Else
        MsgBox errorMessage, vbCritical, FORM_TITLE
    End If
End Sub

' =========================================================
' UI setup
' =========================================================

' Applies the initial form configuration.
Private Sub ConfigureForm()
    Me.Height = FORM_HEIGHT
    Me.Width = FORM_WIDTH
End Sub

' =========================================================
' Data loading
' =========================================================

' Loads all available sections from the active sheet.
Private Sub LoadSections()
    Dim ws As Worksheet
    Dim sectionFinder As clsSectionFinder
    Dim sections() As String
    Dim errorMessage As String
    Dim success As Boolean
    Dim i As Long

    lstSections.Clear

    Set ws = modUtils.GetTargetWorksheet()
    Set sectionFinder = New clsSectionFinder

    success = sectionFinder.GetSectionNames(ws, sections, errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    For i = LBound(sections) To UBound(sections)
        lstSections.AddItem sections(i)
    Next i
End Sub

' =========================================================
' Validation helpers
' =========================================================

' Validates that a section is selected in the list.
Private Function IsSectionSelected() As Boolean
    If lstSections.ListIndex < 0 Then
        MsgBox "Please select a section to remove.", vbExclamation, FORM_TITLE
        IsSectionSelected = False
    Else
        IsSectionSelected = True
    End If
End Function

' Returns the currently selected section name.
Private Function GetSelectedSectionName() As String
    GetSelectedSectionName = CStr(lstSections.list(lstSections.ListIndex))
End Function

' Asks the user to confirm the section removal.
Private Function ConfirmSectionRemoval(ByVal sectionName As String) As Boolean
    ConfirmSectionRemoval = _
        (MsgBox("Are you sure you want to remove section '" & sectionName & "'?", _
                vbQuestion + vbYesNo, _
                "Confirm Remove") = vbYes)
End Function


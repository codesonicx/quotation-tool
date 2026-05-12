Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 400
Private Const FORM_WIDTH As Long = 660
Private Const FORM_TITLE As String = "Remove Item from Assembly"

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
' List events
' =========================================================

Private Sub listSection_Change()
    LoadAssembliesForSelectedSection
End Sub

' =========================================================
' Main workflow
' =========================================================

' Handles the full remove-assembly workflow.
Private Sub HandleRemoveAction()
    Dim ws As Worksheet
    Dim sectionService As clsSectionService
    Dim sectionName As String
    Dim assemblyName As String
    Dim errorMessage As String
    Dim success As Boolean

    If Not IsSectionSelected Then Exit Sub
    If Not IsAssemblySelected Then Exit Sub

    sectionName = GetSelectedSectionName()
    assemblyName = GetSelectedAssemblyName()

    If Not ConfirmAssemblyRemoval(sectionName, assemblyName) Then Exit Sub

    Set ws = GetTargetWorksheet()
    Set sectionService = New clsSectionService

    success = sectionService.RemoveAssemblyFromSection(ws, sectionName, assemblyName, errorMessage)

    If success Then
        MsgBox "Assembly removed successfully.", vbInformation, FORM_TITLE
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

    listSection.Clear
    listItems.Clear

    Set ws = GetTargetWorksheet()
    Set sectionFinder = New clsSectionFinder

    success = sectionFinder.GetSectionNames(ws, sections, errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    For i = LBound(sections) To UBound(sections)
        listSection.AddItem sections(i)
    Next i
End Sub

' Loads all assembly headers for the currently selected section.
Private Sub LoadAssembliesForSelectedSection()
    Dim ws As Worksheet
    Dim sectionFinder As clsSectionFinder
    Dim assemblies() As String
    Dim errorMessage As String
    Dim success As Boolean
    Dim i As Long

    listItems.Clear

    If listSection.ListIndex < 0 Then Exit Sub

    Set ws = GetTargetWorksheet()
    Set sectionFinder = New clsSectionFinder

    success = sectionFinder.GetAssembliesInSection( _
        ws, _
        GetSelectedSectionName(), _
        assemblies, _
        errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    For i = LBound(assemblies) To UBound(assemblies)
        listItems.AddItem assemblies(i)
    Next i
End Sub

' =========================================================
' Validation helpers
' =========================================================

' Validates that a section is selected.
Private Function IsSectionSelected() As Boolean
    If listSection.ListIndex < 0 Then
        MsgBox "Please select a section first.", vbExclamation, FORM_TITLE
        IsSectionSelected = False
    Else
        IsSectionSelected = True
    End If
End Function

' Validates that an assembly is selected.
Private Function IsAssemblySelected() As Boolean
    If listItems.ListIndex < 0 Then
        MsgBox "Please select an assembly to remove.", vbExclamation, FORM_TITLE
        IsAssemblySelected = False
    Else
        IsAssemblySelected = True
    End If
End Function

' Returns the selected section name.
Private Function GetSelectedSectionName() As String
    GetSelectedSectionName = CStr(listSection.list(listSection.ListIndex))
End Function

' Returns the selected assembly name.
Private Function GetSelectedAssemblyName() As String
    GetSelectedAssemblyName = CStr(listItems.list(listItems.ListIndex))
End Function

' Confirms the remove action with the user.
Private Function ConfirmAssemblyRemoval(ByVal sectionName As String, _
                                        ByVal assemblyName As String) As Boolean
    ConfirmAssemblyRemoval = _
        (MsgBox("Are you sure you want to remove assembly '" & assemblyName & _
                "' from section '" & sectionName & "'?", _
                vbQuestion + vbYesNo, _
                "Confirm Remove") = vbYes)
End Function

' =========================================================
' Worksheet helpers
' =========================================================

' Returns the worksheet currently active in Excel.
Private Function GetTargetWorksheet() As Worksheet
    Set GetTargetWorksheet = ActiveSheet
End Function


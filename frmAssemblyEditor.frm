Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 520
Private Const FORM_WIDTH As Long = 700
Private Const FORM_TITLE As String = "Modify Item"

Private mDomain As String
Private mComponentFilter As clsListBoxFilter

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    ConfigureForm

    Set mComponentFilter = New clsListBoxFilter

    mDomain = DOMAIN_ME
    Me.optME.Value = True

    Me.lstCurrentComponents.MultiSelect = fmMultiSelectMulti
    Me.lstComponentLibrary.MultiSelect = fmMultiSelectMulti

    Me.lstCurrentComponents.ColumnCount = 2
    Me.lstCurrentComponents.ColumnWidths = "300 pt;0 pt"

    LoadSections
    LoadComponentLibraryForDomain
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub

' =========================================================
' Button events
' =========================================================

Private Sub btnAdd_Click()
    HandleAddAction
End Sub

Private Sub btnRemove_Click()
    HandleRemoveAction
End Sub

Private Sub btnDone_Click()
    Unload Me
End Sub

Private Sub btnClearSearch_Click()
    Me.txtComponentFilter.Value = vbNullString
    ApplyComponentFilter
End Sub

' =========================================================
' List events
' =========================================================

Private Sub lstSections_Click()
    LoadAssembliesForSelectedSection
End Sub

Private Sub lstAssemblies_Click()
    LoadCurrentComponentsForSelectedAssembly
End Sub

' =========================================================
' Option button events
' =========================================================

Private Sub optME_Click()
    mDomain = DOMAIN_ME
    LoadComponentLibraryForDomain
End Sub

Private Sub optCE_Click()
    mDomain = DOMAIN_CE
    LoadComponentLibraryForDomain
End Sub

' =========================================================
' Search events
' =========================================================

Private Sub txtComponentFilter_Change()
    ApplyComponentFilter
End Sub

' =========================================================
' Main workflows
' =========================================================

Private Sub HandleAddAction()
    Dim ws As Worksheet
    Dim editorService As clsAssemblyEditorService

    Dim sectionName As String
    Dim assemblyName As String
    Dim selectedComponents() As String

    Dim success As Boolean
    Dim errorMessage As String

    If Not IsSectionSelected Then Exit Sub
    If Not IsAssemblySelected Then Exit Sub

    If Not mComponentFilter.GetSelectedItems(Me.lstComponentLibrary, selectedComponents) Then
        MsgBox "Please select at least one component to add.", vbExclamation, FORM_TITLE
        Exit Sub
    End If

    Set ws = modUtils.GetTargetWorksheet()
    Set editorService = New clsAssemblyEditorService

    sectionName = GetSelectedSectionName()
    assemblyName = GetSelectedAssemblyName()

    success = editorService.AddComponentsToAssembly( _
        ws, _
        sectionName, _
        assemblyName, _
        mDomain, _
        selectedComponents, _
        errorMessage)

    If success Then
        mComponentFilter.ClearSelections Me.lstComponentLibrary
        LoadCurrentComponentsForSelectedAssembly
        MsgBox "Component(s) added successfully.", vbInformation, FORM_TITLE
    Else
        MsgBox errorMessage, vbExclamation, FORM_TITLE
    End If
End Sub

Private Sub HandleRemoveAction()
    Dim ws As Worksheet
    Dim editorService As clsAssemblyEditorService

    Dim sectionName As String
    Dim assemblyName As String
    Dim selectedRows() As Long

    Dim success As Boolean
    Dim errorMessage As String

    If Not IsSectionSelected Then Exit Sub
    If Not IsAssemblySelected Then Exit Sub

    If Not GetSelectedCurrentComponentRows(selectedRows) Then
        MsgBox "Please select at least one current component to remove.", vbExclamation, FORM_TITLE
        Exit Sub
    End If

    If MsgBox("Remove selected component(s)?", vbQuestion + vbYesNo, FORM_TITLE) <> vbYes Then
        Exit Sub
    End If

    Set ws = modUtils.GetTargetWorksheet()
    Set editorService = New clsAssemblyEditorService

    sectionName = GetSelectedSectionName()
    assemblyName = GetSelectedAssemblyName()

    success = editorService.RemoveChildrenFromAssembly( _
        ws, _
        sectionName, _
        assemblyName, _
        selectedRows, _
        errorMessage)

    If success Then
        LoadCurrentComponentsForSelectedAssembly
        MsgBox "Component(s) removed successfully.", vbInformation, FORM_TITLE
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

Private Sub LoadSections()
    Dim sectionFinder As clsSectionFinder
    Dim sections() As String
    Dim success As Boolean
    Dim errorMessage As String
    Dim i As Long

    Set sectionFinder = New clsSectionFinder

    Me.lstSections.Clear
    Me.lstAssemblies.Clear
    Me.lstCurrentComponents.Clear

    success = sectionFinder.GetSectionNames( _
        modUtils.GetTargetWorksheet(), _
        sections, _
        errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    For i = LBound(sections) To UBound(sections)
        Me.lstSections.AddItem sections(i)
    Next i
End Sub

Private Sub LoadAssembliesForSelectedSection()
    Dim assemblyFinder As clsAssemblyFinder
    Dim assemblies() As String
    Dim success As Boolean
    Dim errorMessage As String
    Dim i As Long

    Me.lstAssemblies.Clear
    Me.lstCurrentComponents.Clear

    If Me.lstSections.ListIndex = -1 Then Exit Sub

    Set assemblyFinder = New clsAssemblyFinder

    success = assemblyFinder.GetAssembliesInSection( _
        modUtils.GetTargetWorksheet(), _
        GetSelectedSectionName(), _
        assemblies, _
        errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    For i = LBound(assemblies) To UBound(assemblies)
        Me.lstAssemblies.AddItem assemblies(i)
    Next i
End Sub

Private Sub LoadCurrentComponentsForSelectedAssembly()
    Dim editorService As clsAssemblyEditorService
    Dim childNames() As String
    Dim childRows() As Long
    Dim childCount As Long
    Dim success As Boolean
    Dim errorMessage As String
    Dim i As Long
    Dim newIndex As Long

    Me.lstCurrentComponents.Clear

    If Me.lstSections.ListIndex = -1 Then Exit Sub
    If Me.lstAssemblies.ListIndex = -1 Then Exit Sub

    Set editorService = New clsAssemblyEditorService

    success = editorService.GetAssemblyChildren( _
        modUtils.GetTargetWorksheet(), _
        GetSelectedSectionName(), _
        GetSelectedAssemblyName(), _
        childNames, _
        childRows, _
        childCount, _
        errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    If childCount = 0 Then Exit Sub

    For i = 1 To childCount
        Me.lstCurrentComponents.AddItem childNames(i)
        newIndex = Me.lstCurrentComponents.ListCount - 1
        Me.lstCurrentComponents.List(newIndex, 1) = CStr(childRows(i))
    Next i
End Sub

Private Sub LoadComponentLibraryForDomain()
    Dim componentFinder As clsComponentListFinder
    Dim items() As String
    Dim success As Boolean
    Dim errorMessage As String

    Set componentFinder = New clsComponentListFinder

    If mComponentFilter Is Nothing Then
        Set mComponentFilter = New clsListBoxFilter
    End If

    Me.lstComponentLibrary.Clear
    Me.txtComponentFilter.Value = vbNullString
    mComponentFilter.Clear

    success = componentFinder.GetComponentNames( _
        mDomain, _
        items, _
        errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    mComponentFilter.LoadFromArray items
    ApplyComponentFilter
End Sub

Private Sub ApplyComponentFilter()
    If mComponentFilter Is Nothing Then
        Set mComponentFilter = New clsListBoxFilter
    End If

    mComponentFilter.ApplyFilter Me.lstComponentLibrary, Me.txtComponentFilter.Value
End Sub

' =========================================================
' Validation helpers
' =========================================================

Private Function IsSectionSelected() As Boolean
    If Me.lstSections.ListIndex = -1 Then
        MsgBox "Please select a section.", vbExclamation, FORM_TITLE
        IsSectionSelected = False
    Else
        IsSectionSelected = True
    End If
End Function

Private Function IsAssemblySelected() As Boolean
    If Me.lstAssemblies.ListIndex = -1 Then
        MsgBox "Please select an assembly.", vbExclamation, FORM_TITLE
        IsAssemblySelected = False
    Else
        IsAssemblySelected = True
    End If
End Function

' =========================================================
' Selection helpers
' =========================================================

Private Function GetSelectedSectionName() As String
    GetSelectedSectionName = CStr(Me.lstSections.List(Me.lstSections.ListIndex))
End Function

Private Function GetSelectedAssemblyName() As String
    GetSelectedAssemblyName = CStr(Me.lstAssemblies.List(Me.lstAssemblies.ListIndex))
End Function

Private Function GetSelectedCurrentComponentRows(ByRef selectedRows() As Long) As Boolean
    Dim i As Long
    Dim count As Long

    count = 0

    For i = 0 To Me.lstCurrentComponents.ListCount - 1
        If Me.lstCurrentComponents.Selected(i) Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        GetSelectedCurrentComponentRows = False
        Exit Function
    End If

    ReDim selectedRows(1 To count)

    count = 0

    For i = 0 To Me.lstCurrentComponents.ListCount - 1
        If Me.lstCurrentComponents.Selected(i) Then
            count = count + 1
            selectedRows(count) = CLng(Me.lstCurrentComponents.List(i, 1))
        End If
    Next i

    GetSelectedCurrentComponentRows = True
End Function
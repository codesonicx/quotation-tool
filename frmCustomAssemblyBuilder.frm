Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 420
Private Const FORM_WIDTH As Long = 750
Private Const FORM_TITLE As String = "Create Assembly"

Private mComponentFilter As clsListBoxFilter
Private mDomain As String

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    ConfigureForm

    Set mComponentFilter = New clsListBoxFilter

    mDomain = DOMAIN_ME
    Me.optME.value = True

    Me.lstComponents.MultiSelect = fmMultiSelectMulti

    LoadSections
    LoadComponentsForDomain
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub

' =========================================================
' Button events
' =========================================================

Private Sub btnCreate_Click()
    HandleCreateAction
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClearSearch_Click()
    Me.txtComponentFilter.value = vbNullString
    ApplyComponentFilter
End Sub

' =========================================================
' Option button events
' =========================================================

Private Sub optME_Click()
    mDomain = DOMAIN_ME
    LoadComponentsForDomain
End Sub

Private Sub optCE_Click()
    mDomain = DOMAIN_CE
    LoadComponentsForDomain
End Sub

' =========================================================
' Search events
' =========================================================

Private Sub txtComponentFilter_Change()
    ApplyComponentFilter
End Sub

' =========================================================
' Main workflow
' =========================================================

Private Sub HandleCreateAction()
    Dim ws As Worksheet
    Dim builderService As clsCustomAssemblyBuilderService

    Dim sectionName As String
    Dim assemblyName As String
    Dim selectedComponents() As String

    Dim success As Boolean
    Dim errorMessage As String

    If Not IsSectionSelected Then Exit Sub
    If Not IsAssemblyNameEntered Then Exit Sub

    If Not mComponentFilter.GetSelectedItems(Me.lstComponents, selectedComponents) Then
        MsgBox "Please select at least one component.", vbExclamation, FORM_TITLE
        Exit Sub
    End If

    Set ws = modUtils.GetTargetWorksheet()
    Set builderService = New clsCustomAssemblyBuilderService

    sectionName = GetSelectedSectionName()
    assemblyName = Trim$(Me.txtAssemblyName.value)

    success = builderService.InsertCustomAssemblyWithComponents( _
        ws, _
        sectionName, _
        assemblyName, _
        mDomain, _
        selectedComponents, _
        errorMessage)

    If success Then
        MsgBox "Assembly created successfully.", vbInformation, FORM_TITLE
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

Private Sub LoadSections()
    Dim sectionFinder As clsSectionFinder
    Dim sections() As String
    Dim success As Boolean
    Dim errorMessage As String
    Dim i As Long

    Set sectionFinder = New clsSectionFinder

    Me.lstSections.Clear

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

Private Sub LoadComponentsForDomain()
    Dim componentFinder As clsComponentListFinder
    Dim items() As String
    Dim success As Boolean
    Dim errorMessage As String

    Set componentFinder = New clsComponentListFinder

    If mComponentFilter Is Nothing Then
        Set mComponentFilter = New clsListBoxFilter
    End If

    Me.lstComponents.Clear
    Me.txtComponentFilter.value = vbNullString
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

    mComponentFilter.ApplyFilter Me.lstComponents, Me.txtComponentFilter.value
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

Private Function IsAssemblyNameEntered() As Boolean
    If Trim$(Me.txtAssemblyName.value) = vbNullString Then
        MsgBox "Please enter a new assembly name.", vbExclamation, FORM_TITLE
        IsAssemblyNameEntered = False
    Else
        IsAssemblyNameEntered = True
    End If
End Function

' =========================================================
' Selection helpers
' =========================================================

Private Function GetSelectedSectionName() As String
    GetSelectedSectionName = CStr(Me.lstSections.list(Me.lstSections.ListIndex))
End Function


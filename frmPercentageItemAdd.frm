Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 350
Private Const FORM_WIDTH As Long = 530
Private Const FORM_TITLE As String = "Add Percentage Item"

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

Private Sub btnCreate_Click()
    HandleCreateAction
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' =========================================================
' List events
' =========================================================

Private Sub lstSections_Click()
    LoadAssembliesForSelectedSection
End Sub

' =========================================================
' Main workflow
' =========================================================

Private Sub HandleCreateAction()
    Dim ws As Worksheet
    Dim percentService As clsPercentItemService

    Dim sectionName As String
    Dim assemblyName As String
    Dim componentName As String
    Dim percentAmount As Double

    Dim success As Boolean
    Dim errorMessage As String

    If Not IsSectionSelected Then Exit Sub
    If Not IsAssemblySelected Then Exit Sub
    If Not IsComponentNameEntered Then Exit Sub
    If Not IsPercentOptionSelected Then Exit Sub

    Set ws = modUtils.GetTargetWorksheet()

    sectionName = GetSelectedSectionName()
    assemblyName = GetSelectedAssemblyName()
    componentName = Trim$(Me.txtComponentName.Value)
    percentAmount = Val(Me.txtPercentAmount.Value)

    Set percentService = New clsPercentItemService

    success = percentService.InsertPercentItemIntoAssembly( _
        ws, _
        sectionName, _
        assemblyName, _
        componentName, _
        percentAmount, _
        Me.chkDesign.Value, _
        Me.chkConstruction.Value, _
        Me.chkDebug.Value, _
        Me.chkInstallation.Value, _
        errorMessage)

    If success Then
        MsgBox "Percentage item inserted successfully.", vbInformation, FORM_TITLE
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
    Me.lstAssemblies.Clear

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

Private Function IsComponentNameEntered() As Boolean
    If Trim$(Me.txtComponentName.Value) = vbNullString Then
        MsgBox "Please enter a component name.", vbExclamation, FORM_TITLE
        IsComponentNameEntered = False
    Else
        IsComponentNameEntered = True
    End If
End Function

Private Function IsPercentOptionSelected() As Boolean
    If Not Me.chkDesign.Value _
       And Not Me.chkConstruction.Value _
       And Not Me.chkDebug.Value _
       And Not Me.chkInstallation.Value Then

        MsgBox "Please select at least one percentage source.", vbExclamation, FORM_TITLE
        IsPercentOptionSelected = False
    Else
        IsPercentOptionSelected = True
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
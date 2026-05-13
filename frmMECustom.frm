Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 550
Private Const FORM_WIDTH As Long = 310

' Insert mode labels
Private Const MODE_ADD_COMPONENT As String = "Add component"
Private Const MODE_ADD_ASSEMBLY As String = "Add new assembly"

Private mDomain As String

Public Sub OpenForDomain(ByVal domain As String)
    mDomain = UCase$(Trim$(domain))

    If mDomain <> DOMAIN_ME And mDomain <> DOMAIN_CE Then
        MsgBox "Invalid custom item domain: [" & mDomain & "]", vbExclamation
        Exit Sub
    End If

    Me.Caption = mDomain & " Custom Item"
    Me.Show
End Sub

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    ConfigureForm
    LoadInsertModes
    LoadSections
    UpdateAssemblySelectorState
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

Private Sub btnCancel_Click()
    Unload Me
End Sub

' =========================================================
' ComboBox events
' =========================================================

Private Sub cmbInsertMode_Change()
    UpdateAssemblySelectorState
End Sub

Private Sub cmbSection_Change()
    RefreshAssemblyList
End Sub

' =========================================================
' Main workflow
' =========================================================

' Handles the Add button workflow based on the selected mode.
Private Sub HandleAddAction()
    Dim modeText As String

    modeText = Me.cmbInsertMode.Value

    If Not IsSectionSelected Then Exit Sub

    Select Case modeText
        Case MODE_ADD_COMPONENT
            InsertComponentWorkflow

        Case MODE_ADD_ASSEMBLY
            InsertCustomAssemblyWorkflow

        Case Else
            MsgBox "Unknown insert mode selected.", vbExclamation
    End Select
End Sub

' Handles the full process for inserting a custom component.
Private Sub InsertComponentWorkflow()
    Dim sectionName As String
    Dim assemblyName As String
    Dim customData As Collection

    Dim customItemService As clsCustomItemService
    Dim success As Boolean
    Dim errorMessage As String

    If Not IsAssemblySelected Then Exit Sub

    sectionName = Me.cmbSection.value
    assemblyName = Me.lstAssembly.value
    Set customData = BuildCustomData()

    Set customItemService = New clsCustomItemService

    success = customItemService.InsertCustomComponentIntoAssembly( _
        ActiveSheet, _
        sectionName, _
        assemblyName, _
        customData, _
        errorMessage)

    If success Then
        MsgBox "Custom component inserted successfully.", vbInformation
        Unload Me
    Else
        MsgBox errorMessage, vbExclamation
    End If
End Sub

' Handles the full process for inserting a new custom assembly.
Private Sub InsertCustomAssemblyWorkflow()
    Dim sectionName As String
    Dim customData As Collection

    Dim customItemService As clsCustomItemService
    Dim success As Boolean
    Dim errorMessage As String

    sectionName = Me.cmbSection.Value
    Set customData = BuildCustomData()

    Set customItemService = New clsCustomItemService

    success = customItemService.InsertCustomAssemblyIntoSection( _
        ActiveSheet, _
        sectionName, _
        customData, _
        errorMessage)

    If success Then
        MsgBox "Custom assembly inserted successfully.", vbInformation
        Unload Me
    Else
        MsgBox errorMessage, vbExclamation
    End If
End Sub

' =========================================================
' UI setup and state
' =========================================================

' Sets the initial size of the form.
Private Sub ConfigureForm()
    Me.Height = FORM_HEIGHT
    Me.Width = FORM_WIDTH
End Sub

' Loads the available insert modes into the mode ComboBox.
Private Sub LoadInsertModes()
    With Me.cmbInsertMode
        .Clear
        .AddItem MODE_ADD_COMPONENT
        .AddItem MODE_ADD_ASSEMBLY
        .ListIndex = 0
    End With
End Sub

' Enables or disables the assembly selector depending on the selected mode.
Private Sub UpdateAssemblySelectorState()
    Dim enableAssembly As Boolean

    enableAssembly = (Me.cmbInsertMode.value = MODE_ADD_COMPONENT)

    Me.lstAssembly.Enabled = enableAssembly

    If enableAssembly Then
        Me.lstAssembly.BackColor = RGB(255, 255, 255)
        Me.lstAssembly.ForeColor = RGB(0, 0, 0)

        LoadAssembliesForSelectedSection
    Else
        Me.lstAssembly.Clear
        Me.lstAssembly.AddItem "Not required for new assembly"

        Me.lstAssembly.BackColor = RGB(230, 230, 230)
        Me.lstAssembly.ForeColor = RGB(120, 120, 120)
    End If

    UpdateInputFieldsForMode
End Sub

' Refreshes the assembly list only when assembly selection is relevant.
Private Sub RefreshAssemblyList()
    If Me.cmbInsertMode.Value = MODE_ADD_COMPONENT Then
        LoadAssembliesForSelectedSection
    Else
        Me.lstAssembly.Clear
        Me.lstAssembly.AddItem "Not required for new assembly"
    End If
End Sub

' =========================================================
' Data loading
' =========================================================

' Loads all available section names from the active sheet.
Private Sub LoadSections()
    Dim sectionFinder As clsSectionFinder
    Dim success As Boolean
    Dim errorMessage As String
    Dim sections() As String
    Dim i As Long

    Set sectionFinder = New clsSectionFinder
    Me.cmbSection.Clear

    success = sectionFinder.GetSectionNames(ActiveSheet, sections, errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation
        Exit Sub
    End If

    For i = LBound(sections) To UBound(sections)
        Me.cmbSection.AddItem sections(i)
    Next i

    If Me.cmbSection.ListCount > 0 Then
        Me.cmbSection.ListIndex = 0
    End If
End Sub

' Loads assemblies for the currently selected section.
Private Sub LoadAssembliesForSelectedSection()
    Dim assemblyFinder As clsAssemblyFinder
    Dim success As Boolean
    Dim errorMessage As String
    Dim assemblies() As String
    Dim i As Long

    Me.lstAssembly.Clear

    If Me.cmbSection.ListIndex = -1 Then Exit Sub

    Set assemblyFinder = New clsAssemblyFinder
    success = assemblyFinder.GetAssembliesInSection( _
        ActiveSheet, _
        Me.cmbSection.value, _
        assemblies, _
        errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation
        Exit Sub
    End If

    For i = LBound(assemblies) To UBound(assemblies)
        Me.lstAssembly.AddItem assemblies(i)
    Next i
End Sub

' =========================================================
' Validation helpers
' =========================================================

' Validates that a section has been selected.
Private Function IsSectionSelected() As Boolean
    If Me.cmbSection.ListIndex = -1 Then
        MsgBox "Please select a section.", vbExclamation
        IsSectionSelected = False
    Else
        IsSectionSelected = True
    End If
End Function

' Validates that an assembly has been selected.
Private Function IsAssemblySelected() As Boolean
    If Me.lstAssembly.ListIndex = -1 Then
        MsgBox "Please select an assembly.", vbExclamation
        IsAssemblySelected = False
    Else
        IsAssemblySelected = True
    End If
End Function

' =========================================================
' Data collection
' =========================================================

' Collects user-entered values into a keyed collection.
' Keys must match the expectations of the business logic layer.
Private Function BuildCustomData() As Collection
    Dim holder As Collection
    Set holder = New Collection

    holder.Add mDomain, "domain"
    holder.Add Me.txtName.value, "name"
    holder.Add Val(Me.txtDesignHours.value), "des"
    holder.Add Val(Me.txtConstHours.value), "constr"
    holder.Add Val(Me.txtInstallHours.value), "inst"
    holder.Add Val(Me.txtMatRate.value), "mat"
    holder.Add Val(Me.txtDebugHours.value), "debug"
    holder.Add Val(Me.txtStartHours.value), "start"
    holder.Add Val(Me.txtDesNum.value), "desNum"
    holder.Add Val(Me.txtBuildNum.value), "buildNum"

    Set BuildCustomData = holder
End Function

Private Sub UpdateInputFieldsForMode()
    Dim isComponentMode As Boolean

    isComponentMode = (Me.cmbInsertMode.Value = MODE_ADD_COMPONENT)

    Me.txtDesignHours.Enabled = isComponentMode
    Me.txtConstHours.Enabled = isComponentMode
    Me.txtDebugHours.Enabled = isComponentMode
    Me.txtInstallHours.Enabled = isComponentMode
    Me.txtStartHours.Enabled = isComponentMode
    Me.txtMatRate.Enabled = isComponentMode

    Me.txtDesNum.Enabled = Not isComponentMode
    Me.txtBuildNum.Enabled = Not isComponentMode
End Sub
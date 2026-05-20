Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 330
Private Const FORM_WIDTH As Long = 165
Private Const FORM_TITLE As String = "Add Assembly to Library"

Private mDomain As String

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    ConfigureForm

    mDomain = DOMAIN_ME
    Me.optME.value = True

    LoadSections
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

Private Sub btnDone_Click()
    Unload Me
End Sub

' =========================================================
' Option button events
' =========================================================

Private Sub optME_Click()
    mDomain = DOMAIN_ME
End Sub

Private Sub optCE_Click()
    mDomain = DOMAIN_CE
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

Private Sub HandleAddAction()
    Dim ws As Worksheet
    Dim libraryService As clsAssemblyLibraryService

    Dim sectionName As String
    Dim assemblyName As String
    Dim savedAssemblyName As String

    Dim success As Boolean
    Dim errorMessage As String

    If Not IsSectionSelected Then Exit Sub
    If Not IsAssemblySelected Then Exit Sub

    Set ws = modUtils.GetTargetWorksheet()
    Set libraryService = New clsAssemblyLibraryService

    sectionName = GetSelectedSectionName()
    assemblyName = GetSelectedAssemblyName()

    success = libraryService.SaveAssemblyToLibrary( _
        ws, _
        sectionName, _
        assemblyName, _
        mDomain, _
        savedAssemblyName, _
        errorMessage)

    If success Then
        MsgBox "Assembly saved to " & mDomain & " library as '" & savedAssemblyName & "'.", vbInformation, FORM_TITLE
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

    Set assemblyFinder = New clsAssemblyFinder

    Me.lstAssemblies.Clear

    If Me.lstSections.ListIndex = -1 Then Exit Sub

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

' =========================================================
' Selection helpers
' =========================================================

Private Function GetSelectedSectionName() As String
    GetSelectedSectionName = CStr(Me.lstSections.list(Me.lstSections.ListIndex))
End Function

Private Function GetSelectedAssemblyName() As String
    GetSelectedAssemblyName = CStr(Me.lstAssemblies.list(Me.lstAssemblies.ListIndex))
End Function


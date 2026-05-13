Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 220
Private Const FORM_WIDTH As Long = 220
Private Const FORM_TITLE As String = "Control Design"

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
' Main workflow
' =========================================================

Private Sub HandleCreateAction()
    Dim ws As Worksheet
    Dim controlsService As clsControlsDesignService

    Dim selectedSection As String
    Dim errorMessage As String
    Dim success As Boolean

    If Not IsSectionSelected Then Exit Sub

    selectedSection = GetSelectedSectionName()

    Set ws = modUtils.GetTargetWorksheet()
    Set controlsService = New clsControlsDesignService

    success = controlsService.InsertControlsDesignHoursIntoSection( _
        ws, _
        selectedSection, _
        Val(Me.txtHardwareHours.Value), _
        Val(Me.txtPLCHours.Value), _
        Val(Me.txtHMIHours.Value), _
        Val(Me.txtScadaHours.Value), _
        errorMessage)

    If success Then
        MsgBox "Controls design hours inserted successfully.", vbInformation, FORM_TITLE
        Unload Me
    Else
        MsgBox errorMessage, vbCritical, FORM_TITLE
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
    Dim ws As Worksheet
    Dim sectionFinder As clsSectionFinder
    Dim sections() As String
    Dim errorMessage As String
    Dim success As Boolean
    Dim i As Long

    Me.cmbSection.Clear

    Set ws = modUtils.GetTargetWorksheet()
    Set sectionFinder = New clsSectionFinder

    success = sectionFinder.GetSectionNames(ws, sections, errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation, FORM_TITLE
        Exit Sub
    End If

    For i = LBound(sections) To UBound(sections)
        Me.cmbSection.AddItem sections(i)
    Next i

    If Me.cmbSection.ListCount > 0 Then
        Me.cmbSection.ListIndex = 0
    End If
End Sub

' =========================================================
' Validation helpers
' =========================================================

Private Function IsSectionSelected() As Boolean
    If Me.cmbSection.ListIndex < 0 Then
        MsgBox "Please select a section.", vbExclamation, FORM_TITLE
        IsSectionSelected = False
    Else
        IsSectionSelected = True
    End If
End Function

Private Function GetSelectedSectionName() As String
    GetSelectedSectionName = CStr(Me.cmbSection.Value)
End Function
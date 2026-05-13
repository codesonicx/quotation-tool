Option Explicit

Private mDomain As String
Private mAllItems() As String
Private mItemCount As Long

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    Me.Height = 500
    Me.Width = 600

    Me.ItemList.MultiSelect = fmMultiSelectMulti
    Me.txtItemFilter.Value = vbNullString
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub

' =========================================================
' Public entry point
' =========================================================

Public Sub OpenForDomain(ByVal domain As String)
    mDomain = UCase$(Trim$(domain))

    If mDomain <> DOMAIN_ME And mDomain <> DOMAIN_CE Then
        MsgBox "Domain is not set. Open this form using OpenForDomain(""ME"") or OpenForDomain(""CE"")." & vbCrLf & _
               "Current value: [" & mDomain & "]", vbExclamation
        Exit Sub
    End If

    LoadLists
    Me.Show
End Sub

' =========================================================
' Search / filter events
' =========================================================

Private Sub txtItemFilter_Change()
    ApplyItemFilter
End Sub

Private Sub btnClearSearch_Click()
    Me.txtItemFilter.Value = vbNullString
    ApplyItemFilter
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
' Main workflow
' =========================================================

Private Sub HandleAddAction()
    Dim sectionName As String
    Dim selectedItems() As String
    Dim i As Long
    Dim count As Long

    Dim assemblyService As clsAssemblyService
    Dim success As Boolean
    Dim errorMessage As String

    If Me.SectionList.ListIndex = -1 Then
        MsgBox "Please select a section first.", vbExclamation
        Exit Sub
    End If

    sectionName = Me.SectionList.Value

    count = 0

    For i = 0 To Me.ItemList.ListCount - 1
        If Me.ItemList.Selected(i) Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        MsgBox "Please select at least one item to add.", vbExclamation
        Exit Sub
    End If

    ReDim selectedItems(1 To count)

    count = 0

    For i = 0 To Me.ItemList.ListCount - 1
        If Me.ItemList.Selected(i) Then
            count = count + 1
            selectedItems(count) = Me.ItemList.List(i)
        End If
    Next i

    Set assemblyService = New clsAssemblyService

    success = assemblyService.InsertAssembliesIntoSection( _
        ActiveSheet, _
        sectionName, _
        selectedItems, _
        mDomain, _
        errorMessage)

    If success Then
        MsgBox "Assemblies inserted into '" & sectionName & "'.", vbInformation
        Unload Me
    Else
        MsgBox errorMessage, vbExclamation
    End If
End Sub

' =========================================================
' Data loading
' =========================================================

Private Sub LoadLists()
    LoadSections
    LoadAssemblyItems
End Sub

Private Sub LoadSections()
    Dim sectionFinder As clsSectionFinder
    Dim success As Boolean
    Dim errorMessage As String
    Dim sections() As String
    Dim i As Long

    Set sectionFinder = New clsSectionFinder

    Me.SectionList.Clear

    success = sectionFinder.GetSectionNames(ActiveSheet, sections, errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation
        Exit Sub
    End If

    For i = LBound(sections) To UBound(sections)
        Me.SectionList.AddItem sections(i)
    Next i
End Sub

Private Sub LoadAssemblyItems()
    Dim assemblyService As clsAssemblyService
    Dim success As Boolean
    Dim errorMessage As String
    Dim items() As String
    Dim i As Long

    Set assemblyService = New clsAssemblyService

    Me.ItemList.Clear
    mItemCount = 0
    Erase mAllItems

    success = assemblyService.GetAssemblyDisplayList( _
        ThisWorkbook.Worksheets(SHEET_ASSEMBLY_LIST), _
        mDomain, _
        items, _
        errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation
        Exit Sub
    End If

    mItemCount = UBound(items) - LBound(items) + 1
    ReDim mAllItems(1 To mItemCount)

    For i = LBound(items) To UBound(items)
        mAllItems(i - LBound(items) + 1) = items(i)
    Next i

    ApplyItemFilter
End Sub

' =========================================================
' Filtering
' =========================================================

Private Sub ApplyItemFilter()
    Dim filterText As String
    Dim itemText As String
    Dim i As Long

    Me.ItemList.Clear

    If mItemCount = 0 Then Exit Sub

    filterText = LCase$(Trim$(Me.txtItemFilter.Value))

    For i = 1 To mItemCount
        itemText = CStr(mAllItems(i))

        If filterText = vbNullString Then
            Me.ItemList.AddItem itemText
        ElseIf InStr(1, LCase$(itemText), filterText, vbTextCompare) > 0 Then
            Me.ItemList.AddItem itemText
        End If
    Next i
End Sub
Option Explicit

Private mDomain As String

Private Sub UserForm_Initialize()
    Me.Height = 500
    Me.Width = 600
    Me.ItemList.MultiSelect = fmMultiSelectMulti

End Sub


Public Sub OpenForDomain(ByVal domain As String)
    mDomain = domain
    
    ' Safety: if the form was shown directly (not via OpenForDomain),
    If mDomain <> "ME" And mDomain <> "CE" Then
        MsgBox "Domain is not set. Open this form using OpenForDomain(""ME"") or OpenForDomain(""CE"")." & vbCrLf & _
               "Current value: [" & mDomain & "]", vbExclamation
        Exit Sub
    End If

    LoadLists
    Me.Show
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub


Private Sub LoadLists()
    Dim sectionFinder As New clsSectionFinder
    Dim assemblyService As New clsAssemblyService
    Dim success As Boolean
    Dim errorMessage As String
    Dim items() As String
    Dim sections() As String
    Dim i As Long

    ' Load available sections from the active target sheet.
    Me.SectionList.Clear
    success = sectionFinder.GetSectionNames(ActiveSheet, sections, errorMessage)

    If success Then
        For i = LBound(sections) To UBound(sections)
            Me.SectionList.AddItem sections(i)
        Next i
    Else
        MsgBox errorMessage, vbExclamation
    End If

    ' Load available assembly items from the library sheet,
    ' filtered by the selected domain (ME / CE).
    Me.ItemList.Clear
    success = assemblyService.GetAssemblyDisplayList(ThisWorkbook.Worksheets("Assembly List"), mDomain, items, errorMessage)

    If success Then
        For i = LBound(items) To UBound(items)
            Me.ItemList.AddItem items(i)
        Next i
    Else
        MsgBox errorMessage, vbExclamation
    End If
End Sub

Private Sub btnAdd_Click()
    Dim sectionName As String
    Dim selectedItems() As String
    Dim i As Long, count As Long

    Dim assemblyService As New clsAssemblyService
    Dim success As Boolean
    Dim errorMessage As String

    If Me.SectionList.ListIndex = -1 Then
        MsgBox "Please select a section first.", vbExclamation
        Exit Sub
    End If
    sectionName = Me.SectionList.value

    count = 0
    For i = 0 To Me.ItemList.ListCount - 1
        If Me.ItemList.Selected(i) Then count = count + 1
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
            selectedItems(count) = Me.ItemList.list(i)
        End If
    Next i

    success = assemblyService.InsertAssembliesIntoSection(ActiveSheet, sectionName, selectedItems, mDomain, errorMessage)

    If success Then
        MsgBox "Assemblies inserted into '" & sectionName & "'.", vbInformation
        Unload Me
    Else
        MsgBox errorMessage, vbExclamation
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


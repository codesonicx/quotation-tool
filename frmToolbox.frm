Private Sub UserForm_Initialize()
    Me.Height = 300
    Me.Width = 520
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub


Private Sub btnNewSheet_Click()

    Dim sheetName As String
    Dim manager As New clsSheetManager
    Dim success As Boolean
    Dim errorMessage As String

    sheetName = InputBox("Name of new sheet:", "Inserting New Sheet")
    
    success = manager.CreateNewSheet(sheetName, errorMessage)
    
    If Not success Then
        MsgBox errorMessage
    End If

End Sub

Private Sub btnNewSection_Click()

    Dim sectionName As String
    Dim manager As New clsSectionManager
    Dim success As Boolean
    Dim errorMessage As String

    sectionName = InputBox("Section Name:", "Add New Section")

    success = manager.AddSection(ActiveSheet, sectionName, errorMessage)

    If Not success Then
        MsgBox errorMessage, vbExclamation
    End If

End Sub

Private Sub btnMEAssembly_Click()
    frmAssemblyPicker.OpenForDomain "ME"
End Sub


Private Sub btnCEAssembly_Click()
    frmAssemblyPicker.OpenForDomain "CE"
End Sub

Private Sub btnMECustom_Click()
    frmMECustom.Show
End Sub

Private Sub btnRemoveSection_Click()
    frmSectionRemove.Show
End Sub

Private Sub btnRemoveAssembly_Click()
    frmItemRemove.Show
End Sub

Option Explicit

' =========================================================
' clsListBoxFilter
'
' Purpose:
' Stores a master list of text items, applies a text filter
' to a target ListBox, and preserves selected items across
' filter changes.
'
' Notes:
' - Filtering is case-insensitive.
' - Selections are preserved even when filtered items are hidden.
' - This class assumes item display text is unique.
' =========================================================

Private mItems() As String
Private mItemCount As Long
Private mSelectedItems As Object

Private Sub Class_Initialize()
    Set mSelectedItems = CreateObject("Scripting.Dictionary")
    mSelectedItems.CompareMode = vbTextCompare
End Sub

Public Sub LoadFromArray(ByRef sourceItems() As String)
    Dim i As Long
    Dim targetIndex As Long

    Clear

    mItemCount = UBound(sourceItems) - LBound(sourceItems) + 1
    ReDim mItems(1 To mItemCount)

    targetIndex = 0

    For i = LBound(sourceItems) To UBound(sourceItems)
        targetIndex = targetIndex + 1
        mItems(targetIndex) = CStr(sourceItems(i))
    Next i
End Sub

Public Sub LoadFromListBox(ByVal sourceListBox As MSForms.ListBox)
    Dim i As Long

    Clear

    If sourceListBox.ListCount = 0 Then Exit Sub

    mItemCount = sourceListBox.ListCount
    ReDim mItems(1 To mItemCount)

    For i = 0 To sourceListBox.ListCount - 1
        mItems(i + 1) = CStr(sourceListBox.List(i))
    Next i
End Sub

Public Sub ApplyFilter(ByVal targetListBox As MSForms.ListBox, _
                       ByVal filterText As String)
    Dim i As Long
    Dim currentItem As String
    Dim normalizedFilter As String

    SyncSelectionsFromListBox targetListBox

    targetListBox.Clear

    If mItemCount = 0 Then Exit Sub

    normalizedFilter = LCase$(Trim$(filterText))

    For i = 1 To mItemCount
        currentItem = CStr(mItems(i))

        If normalizedFilter = vbNullString Then
            AddVisibleItem targetListBox, currentItem
        ElseIf ItemMatchesFilter(currentItem, normalizedFilter) Then
            AddVisibleItem targetListBox, currentItem
        End If
    Next i
End Sub

Public Function GetSelectedItems(ByVal currentListBox As MSForms.ListBox, _
                                 ByRef selectedItems() As String) As Boolean
    Dim i As Long
    Dim selectedCount As Long

    SyncSelectionsFromListBox currentListBox

    selectedCount = 0

    For i = 1 To mItemCount
        If mSelectedItems.Exists(MakeKey(mItems(i))) Then
            selectedCount = selectedCount + 1
        End If
    Next i

    If selectedCount = 0 Then
        GetSelectedItems = False
        Exit Function
    End If

    ReDim selectedItems(1 To selectedCount)

    selectedCount = 0

    For i = 1 To mItemCount
        If mSelectedItems.Exists(MakeKey(mItems(i))) Then
            selectedCount = selectedCount + 1
            selectedItems(selectedCount) = mItems(i)
        End If
    Next i

    GetSelectedItems = True
End Function

Public Function SelectedCount(ByVal currentListBox As MSForms.ListBox) As Long
    SyncSelectionsFromListBox currentListBox
    SelectedCount = mSelectedItems.Count
End Function

Public Sub ClearSelections(ByVal targetListBox As MSForms.ListBox)
    Dim i As Long

    mSelectedItems.RemoveAll

    For i = 0 To targetListBox.ListCount - 1
        targetListBox.Selected(i) = False
    Next i
End Sub

Public Sub Clear()
    mItemCount = 0
    Erase mItems

    If Not mSelectedItems Is Nothing Then
        mSelectedItems.RemoveAll
    End If
End Sub

Public Function HasItems() As Boolean
    HasItems = (mItemCount > 0)
End Function

Public Function ItemCount() As Long
    ItemCount = mItemCount
End Function

Private Sub AddVisibleItem(ByVal targetListBox As MSForms.ListBox, _
                           ByVal itemText As String)
    Dim newIndex As Long

    targetListBox.AddItem itemText
    newIndex = targetListBox.ListCount - 1

    If mSelectedItems.Exists(MakeKey(itemText)) Then
        targetListBox.Selected(newIndex) = True
    End If
End Sub

Private Sub SyncSelectionsFromListBox(ByVal sourceListBox As MSForms.ListBox)
    Dim i As Long
    Dim itemText As String
    Dim itemKey As String

    If sourceListBox Is Nothing Then Exit Sub

    For i = 0 To sourceListBox.ListCount - 1
        itemText = CStr(sourceListBox.List(i))
        itemKey = MakeKey(itemText)

        If sourceListBox.Selected(i) Then
            If Not mSelectedItems.Exists(itemKey) Then
                mSelectedItems.Add itemKey, itemText
            End If
        Else
            If mSelectedItems.Exists(itemKey) Then
                mSelectedItems.Remove itemKey
            End If
        End If
    Next i
End Sub

Private Function ItemMatchesFilter(ByVal itemText As String, _
                                   ByVal normalizedFilter As String) As Boolean
    ItemMatchesFilter = _
        (InStr(1, LCase$(itemText), normalizedFilter, vbTextCompare) > 0)
End Function

Private Function MakeKey(ByVal itemText As String) As String
    MakeKey = LCase$(Trim$(itemText))
End Function
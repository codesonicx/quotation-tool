Option Explicit

' =========================================================
' clsListBoxFilter
'
' Purpose:
' Stores a master list of text items and applies a text filter
' to a target ListBox.
'
' Notes:
' - Filtering is case-insensitive.
' - Current selection is cleared when the list is rebuilt.
' - This class can be reused by multiple forms.
' =========================================================

Private mItems() As String
Private mItemCount As Long

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

    targetListBox.Clear

    If mItemCount = 0 Then Exit Sub

    normalizedFilter = LCase$(Trim$(filterText))

    For i = 1 To mItemCount
        currentItem = CStr(mItems(i))

        If normalizedFilter = vbNullString Then
            targetListBox.AddItem currentItem
        ElseIf ItemMatchesFilter(currentItem, normalizedFilter) Then
            targetListBox.AddItem currentItem
        End If
    Next i
End Sub

Public Sub Clear()
    mItemCount = 0
    Erase mItems
End Sub

Public Function HasItems() As Boolean
    HasItems = (mItemCount > 0)
End Function

Public Function ItemCount() As Long
    ItemCount = mItemCount
End Function

Private Function ItemMatchesFilter(ByVal itemText As String, _
                                   ByVal normalizedFilter As String) As Boolean
    ItemMatchesFilter = _
        (InStr(1, LCase$(itemText), normalizedFilter, vbTextCompare) > 0)
End Function
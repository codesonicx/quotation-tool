Option Explicit

' =========================================================
' clsHeaderFinder
'
' Purpose:
' Finds worksheet columns by header labels.
' =========================================================

Public Function FindHeaderColumn(ByVal ws As Worksheet, _
                                 ByVal headerText As String) As Long
    Dim foundCell As Range

    Set foundCell = ws.Rows(1).Find( _
        What:=headerText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        FindHeaderColumn = 0
    Else
        FindHeaderColumn = foundCell.Column
    End If
End Function
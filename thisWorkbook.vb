Option Explicit

Dim ciCollection As New Collection
Dim agentCollection As New Collection

Public Sub ShowAssignmentForm()
    initializeCICollection
    initializeAgentCollection

    initializeCIBox

    AssignmentForm.Show

    Set ciCollection = Nothing
    Set agentCollection = Nothing

End Sub

Private Sub initializeCICollection()
    Dim currentCellLine As Integer
    For currentCellLine = 2 To getLastCIRow()
        Dim ciName As String
        ciName = getCiNameAtLine(currentCellLine)
        ciCollection.Add (ciName)
    Next currentCellLine
End Sub

Private Sub initializeAgentCollection()
End Sub

Private Sub initializeCIBox()
    Dim Item As Variant
    For Each Item In ciCollection
        AssignmentForm.CI.AddItem (Item)
    Next Item
End Sub

Private Function getCiNameAtLine(line As Integer) As String
    getCiNameAtLine = Sheets("Matrix").Cells(line, 1)
End Function

Private Sub addCiNameToCIBox(name As String)
    If name <> "" Then
        AssignmentForm.CI.AddItem (name)
    End If
End Sub

Private Function getLastCIRow() As Integer
    getLastCIRow = Sheets("Matrix").Range("A" & Rows.Count).End(xlUp).Row
End Function

Private Function getLastAgentColumn() As Integer
    getLastAgentColumn = Sheets("Matrix").Cells(1, Columns.Count).End(xlToLeft).column
End Function


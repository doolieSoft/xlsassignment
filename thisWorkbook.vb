Option Explicit

Public Sub ShowAssignmentForm()
    Dim i As Integer
    Dim ciName As String

    For i = 1 To getLastCIRow()
        ciName = Sheets("Matrix").Cells(i, 1)
        addCiNameToCIBox (ciName)
    Next i
    AssignmentForm.Show
End Sub

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

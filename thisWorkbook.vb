Option Explicit

Dim ciCollection As New Collection
Dim agentCollection As New Collection
Dim MatrixSheetName As String
Dim TotalSheetName As String

Public Sub ShowAssignmentForm()
    MatrixSheetName = "Matrix"
    TotalSheetName = "Total"
    
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
    Dim currentCellColumn As Integer
    
    For currentCellColumn = 2 To getLastAgentColumn()
        Dim agentName As String
        agentName = getAgentNameAtColumn(currentCellColumn)
        agentCollection.Add (agentName)
    Next currentCellColumn
    
End Sub

Private Sub initializeCIBox()
    Dim Item As Variant
    
    For Each Item In ciCollection
        AssignmentForm.CI.AddItem (Item)
    Next Item
End Sub

Public Function getLastCIRow() As Integer
    getLastCIRow = Sheets(MatrixSheetName).Range("A" & Rows.Count).End(xlUp).Row
End Function

Private Function getCiNameAtLine(line As Integer) As String
    getCiNameAtLine = Sheets(MatrixSheetName).Cells(line, 1)
End Function

Private Function getAgentNameAtColumn(column As Integer) As String
    getAgentNameAtColumn = Sheets(MatrixSheetName).Cells(1, column)
End Function

Private Function getLastAgentColumn() As Integer
    getLastAgentColumn = Sheets(MatrixSheetName).Cells(1, Columns.Count).End(xlToLeft).column
End Function


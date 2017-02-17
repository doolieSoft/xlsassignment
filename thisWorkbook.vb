Option Explicit

Dim ciCollection As Collection
Dim agentCollection As Collection
Public MatrixSheetName As String
Public TotalSheetName As String

Public Sub ShowAssignmentForm()
    MatrixSheetName = "Matrix Ratio"
    TotalSheetName = "Total"
    Set ciCollection = New Collection
    Set agentCollection = New Collection
    
    initializeCICollection
    initializeAgentCollection
    initializeCIBox

    AssignmentForm.Show

    Set ciCollection = Nothing
    Set agentCollection = Nothing
    Set AssignmentForm = Nothing
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

Public Function getColumnByAgentName(agentNameToFind As String) As Integer
    Dim currentAgent As String
    Dim currentCellColumn As Integer
    currentCellColumn = 2
    
    currentAgent = Sheets(MatrixSheetName).Cells(1, currentCellColumn)
    While currentAgent <> agentNameToFind And currentCellColumn <= getLastAgentColumn()
        currentCellColumn = currentCellColumn + 1
        currentAgent = Sheets(MatrixSheetName).Cells(1, currentCellColumn)
    Wend
    If currentAgent = agentNameToFind Then
        getColumnByAgentName = currentCellColumn
    End If
End Function

Private Function getLastAgentColumn() As Integer
    getLastAgentColumn = Sheets(MatrixSheetName).Cells(1, Columns.Count).End(xlToLeft).column
End Function

Public Function getCILine(ciNameToFind As String) As Integer
    Dim currentCI As String
    Dim currentCellLine As Integer
    currentCellLine = 2
    
    currentCI = Sheets(MatrixSheetName).Cells(currentCellLine, 1)
    While currentCI <> ciNameToFind And currentCellLine <= getLastCIRow()
        currentCellLine = currentCellLine + 1
        currentCI = Sheets(MatrixSheetName).Cells(currentCellLine, 1)
    Wend
    If currentCI = ciNameToFind Then
        getCILine = currentCellLine
    End If
End Function

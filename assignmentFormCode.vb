Private Sub CI_Change()
    Dim ciName As String
    Dim ciRange As Range
    Dim ciLine As Integer
    
    ciName = CI.Value
    Set ciRange = Worksheets("Matrix").Range("A1:A10000").Find(ciName, lookat:=xlPart)
    ciLine = ciRange.Row
    Debug.Print (ciLine)
    
    ListAgent.Clear
    
    Dim nbAgent As Integer
    Dim agentName As String
    
    nbAgent = getLastAgentColumn
    For i = 1 To nbAgent
        If Sheets("Matrix").Cells(ciLine, i) > 0 Then
            addAgentToListAgent (i)
        End If
    Next i
    
    ListAgent.AddItem
End Sub

Private Function getLastAgentColumn() As Integer
    getLastAgentColumn = Sheets("Matrix").Cells(1, Columns.Count).End(xlToLeft).column
End Function

Private Sub addAgentToListAgent(agentColumn As Integer)
    If Cells(1, agentColumn) <> "" Then
        ListAgent.AddItem (Cells(1, agentColumn))
    End If
End Sub

Private Sub AssignButton_Click()
    Debug.Print ListAgent.Value
    For r = 0 To ListAgent.ListCount - 1
    If ListAgent.Selected(r) Then
        Sheets("Total").Cells(10, 10) = "coucou"
        Debug.Print "You selected row #" & r + 1
        Debug.Print ListAgent.List(r, 1)
    End If
Next
End Sub

Private Sub CI_Change()
    Dim ciName As String
    Dim ciRange As Range
    Dim ciLine As Integer
    Dim nbAgent As Integer
    Dim agentName As String
    Dim agentColumn As Integer

    ciName = CI.Value
    Set ciRange = Worksheets("Matrix").Range("A1:A10000").Find(ciName, lookat:=xlPart)
    ciLine = ciRange.Row
    ListAgent.Clear
    
    nbAgent = getLastAgentColumn
    For agentColumn = 2 To nbAgent
        agentName = Sheets("Matrix").Cells(1, agentColumn)
        If IsNumeric(Sheets("Matrix").Cells(ciLine, agentColumn)) Then
                
            If Sheets("Matrix").Cells(ciLine, agentColumn) > 0 Then
                Dim idLastAgentAdded As Integer
                idLastAgentAdded = addAgentToListAgent(agentName)
                
                Dim totalTicketsForCISelected As Integer
                Dim ratio As Double
                Dim grandTotalForAgent As Integer
                
                totalTicketsForCISelected = Sheets("Total").Cells(ciLine, agentColumn)
                grandTotalForAgent = getGrandTotalForAgent(agentColumn)
                ratio = 0
    
                If grandTotalForAgent > 0 Then
                    ratio = Round(totalTicketsForCISelected / grandTotalForAgent * 100, 2)
                End If
                updateRatioInListAgent idLastAgentAdded, ratio
                updateTotalInListAgent idLastAgentAdded, totalTicketsForCISelected
            End If
        End If
    Next agentColumn
End Sub

Private Function getLastAgentColumn() As Integer
    getLastAgentColumn = Sheets("Matrix").Cells(1, Columns.Count).End(xlToLeft).column
End Function

Private Function addAgentToListAgent(agentName As String) As Integer
    If agentName <> "" Then
        ListAgent.AddItem (agentName)
    End If
    addAgentToListAgent = ListAgent.ListCount
End Function

Private Sub updateTotalInListAgent(id As Integer, totalTicket As Integer)
    ListAgent.List(ListAgent.ListCount - 1, 2) = totalTicket
End Sub

Private Sub updateRatioInListAgent(idLastAgentAdded As Integer, ratio As Double)
    ListAgent.List(ListAgent.ListCount - 1, 1) = ratio
End Sub

Private Function getGrandTotalForAgent(agentColumn As Integer) As Integer
    getGrandTotalForAgent = Application.WorksheetFunction.Sum(Range(Sheets("Total").Cells(2, agentColumn), Sheets("Total").Cells(ThisWorkbook.getLastCIRow, agentColumn)))
End Function

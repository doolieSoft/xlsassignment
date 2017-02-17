Private Sub AssignButton_Click()
    Dim ciLine As Integer
    Dim agentColumn As Integer
    Dim ratio As Double
    Dim nbTicketsForCISelected As Integer
    Dim grandTotalForAgent As Integer
    
    For r = 0 To ListAgent.ListCount - 1
        If ListAgent.Selected(r) Then
            ciLine = ThisWorkbook.getCILine(CI.Value)
            agentColumn = ThisWorkbook.getColumnByAgentName(ListAgent.List(r))
            addOneTicketToAgent ciLine, agentColumn
            nbTicketsForCISelected = Sheets(ThisWorkbook.TotalSheetName).Cells(ciLine, agentColumn)
            grandTotalForAgent = getGrandTotalForAgent(agentColumn)
            ratio = getRatio(nbTicketsForCISelected, grandTotalForAgent)
            updateRatioInListAgent r, ratio
            updateTotalInListAgent r, nbTicketsForCISelected
            updateGrandTotalInListAgent r, grandTotalForAgent
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
    Dim nbTicketsForCISelected As Integer
    Dim ratio As Double
    Dim theoriticalRatio As Double
    Dim grandTotalForAgent As Integer
    Dim idLastAgentAdded As Integer
    
    ciName = CI.Value
    ciLine = ThisWorkbook.getCILine(CI.Value)
    ListAgent.Clear
    
    nbAgent = getLastAgentColumn
    For agentColumn = 2 To nbAgent
        agentName = Sheets(ThisWorkbook.MatrixSheetName).Cells(1, agentColumn)
        If IsNumeric(Sheets(ThisWorkbook.MatrixSheetName).Cells(ciLine, agentColumn)) Then
            If Sheets(ThisWorkbook.MatrixSheetName).Cells(ciLine, agentColumn) > 0 Then
                idLastAgentAdded = addAgentToListAgent(agentName)
                nbTicketsForCISelected = Sheets(ThisWorkbook.TotalSheetName).Cells(ciLine, agentColumn)
                grandTotalForAgent = getGrandTotalForAgent(agentColumn)
                ratio = getRatio(nbTicketsForCISelected, grandTotalForAgent)
                theoriticalRatio = getTheoriticalRatio(ciLine, agentColumn)
                updateRatioInListAgent idLastAgentAdded - 1, ratio
                updateTheoriticalRatioInListAgent idLastAgentAdded - 1, theoriticalRatio
                updateTotalInListAgent idLastAgentAdded - 1, nbTicketsForCISelected
                updateGrandTotalInListAgent idLastAgentAdded - 1, grandTotalForAgent
            End If
        End If
    Next agentColumn
End Sub

Private Function getLastAgentColumn() As Integer
    getLastAgentColumn = Sheets(ThisWorkbook.MatrixSheetName).Cells(1, Columns.Count).End(xlToLeft).column
End Function

Private Function addAgentToListAgent(agentName As String) As Integer
    If agentName <> "" Then
        ListAgent.AddItem (agentName)
    End If
    addAgentToListAgent = ListAgent.ListCount
End Function

Private Sub addOneTicketToAgent(ciLine As Integer, agentColumn As Integer)
    Sheets(ThisWorkbook.TotalSheetName).Cells(ciLine, agentColumn) = Sheets(ThisWorkbook.TotalSheetName).Cells(ciLine, agentColumn) + 1
End Sub

Private Sub updateTheoriticalRatioInListAgent(ByVal idLastAgentAdded As Integer, theoriticalRatio As Double)
    ListAgent.List(idLastAgentAdded, 1) = theoriticalRatio
End Sub

Private Sub updateRatioInListAgent(ByVal idLastAgentAdded As Integer, ratio As Double)
    ListAgent.List(idLastAgentAdded, 2) = ratio
End Sub

Private Sub updateTotalInListAgent(ByVal id As Integer, totalTicket As Integer)
    ListAgent.List(id, 3) = totalTicket
End Sub

Private Sub updateGrandTotalInListAgent(ByVal id As Integer, grandTotalForAgent As Integer)
    ListAgent.List(id, 4) = grandTotalForAgent
End Sub

Private Function getGrandTotalForAgent(agentColumn As Integer) As Integer
    getGrandTotalForAgent = Application.WorksheetFunction.Sum(Range(Sheets(ThisWorkbook.TotalSheetName).Cells(2, agentColumn), Sheets(ThisWorkbook.TotalSheetName).Cells(ThisWorkbook.getLastCIRow, agentColumn)))
End Function

Private Function getRatio(nbTicketsForCISelected As Integer, grandTotalForAgent As Integer) As Double
    Dim ratio As Double
    ratio = 0

    If grandTotalForAgent > 0 Then
        ratio = Round(nbTicketsForCISelected / grandTotalForAgent * 100, 2)
    End If
    getRatio = ratio
End Function

Private Function getTheoriticalRatio(ciLine As Integer, agentColumn As Integer) As Double
    Dim ratio As Double
    ratio = Sheets(ThisWorkbook.MatrixSheetName).Cells(ciLine, agentColumn) * 100
    getTheoriticalRatio = ratio
End Function


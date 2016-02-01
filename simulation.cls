Option Explicit
Dim dealerValue2 As Integer
Dim simIndex As Integer
Dim simIndexRange As Range
Dim dealerHand2  As Range
Dim status As Range
Dim rand As Integer
Dim card As String
Dim str As String
Sub simulation()
    Dim constantTrial As Integer
    Dim j As Integer
    
    constantTrial = 25000
    Set simIndexRange = Sheet2.Range("A4")
    Set dealerHand2 = Sheet2.Range("B4")
    Set status = Sheet2.Range("C4")
    str = "K"
    simIndex = 1
    dealerHand2 = str
    
    For j = 0 To constantTrial
    dealerValue2 = valueCalc(dealerHand2.Text)
        Do Until dealerValue2 > 16
            Call hitRandom
            dealerValue2 = valueCalc(dealerHand2.Text)
        Loop
       
        simIndexRange = simIndex
        simIndex = simIndex + 1
        
        If dealerValue2 > 21 Then
            status = "BUSTED"
        Else
            status = dealerValue2
        End If
        Set simIndexRange = simIndexRange.Offset(1, 0)
        Set dealerHand2 = dealerHand2.Offset(1, 0)
        Set status = status.Offset(1, 0)
        dealerHand2 = str
        Next j
End Sub
Sub hitRandom()
    rand = WorksheetFunction.RandBetween(1, 13)
    
    If (rand) < 11 And (rand) > 1 Then
            card = rand
        ElseIf (rand) = 11 Then
            card = "J"
        ElseIf (rand) = 12 Then
            card = "Q"
        ElseIf (rand) = 0 Then
            card = "K"
        ElseIf (rand) = 1 Then
            card = "A"
    End If
    
    dealerHand2 = dealerHand2.Text + "," + card
    
End Sub


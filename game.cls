Option Explicit
Dim deck(311) As String
Dim Hand As String
Dim deckIndex As Integer
Dim initialDeckIndex As Integer
Dim dealer2Card As String
Dim dealerValue As Integer
Dim playerValue As Integer
Dim bet As Integer
Dim money As Double
Dim handIndex As Integer
Dim dealerHand  As Range
Dim playerHand As Range
Dim winCheckRange As Range
Dim moneyRange As Range
Dim regularBet As Range
Dim handNumber As Range
Dim dealerBust(1 To 13, 1 To 6) As Double
Dim numberOfCards(12) As Integer
Dim playerOdd(12) As Double
Dim numberofSim As Integer
Sub simulation()
    Dim i As Integer
    numberofSim = InputBox("Enter the number of hand you want to simulate")
    For i = 0 To numberofSim - 1
        Call dealClicked
        Do Until Not winCheckRange = ""
            If Range("K2").Value > Range("K3").Value Then
                Call standClicked
            Else
                Call hitClicked
            End If
        Loop
    Next i
End Sub
Sub martingaleSimulation()
    Dim i As Integer
    Dim initialBet As Integer
    initialBet = bet
    numberofSim = InputBox("Enter the number of hand you want to simulate")
    For i = 0 To numberofSim - 1
        Call dealClicked
        Do Until Not winCheckRange = ""
            If Range("K2").Value > Range("K3").Value Then
                Call standClicked
            Else
                Call hitClicked
            End If
        Loop
        If winCheckRange.Text = "Lost" Then
            changeBet (bet * 2)
        ElseIf winCheckRange.Text = "Win" Or _
            winCheckRange.Text = "BLACKJACK" Then
            changeBet (initialBet)
        End If
    Next i
End Sub
Sub heuristic()
    Dim i As Integer
    numberofSim = InputBox("Enter the number of hand you want to simulate")
    For i = 0 To numberofSim - 1
        Call dealClicked
        Do Until Not winCheckRange = ""
            If playerValue > 13 Then
                Call standClicked
                Exit Do
            End If
            Call hitClicked
        Loop
    Next i
End Sub
Sub martingaleHeuristic()
    Dim i As Integer
    Dim initialBet As Integer
    initialBet = bet
    numberofSim = InputBox("Enter the number of hand you want to simulate")
    For i = 0 To numberofSim - 1
        Call dealClicked
        Do Until Not winCheckRange = ""
            If playerValue > 13 Then
                Call standClicked
                Exit Do
            End If
            Call hitClicked
        Loop
        If winCheckRange.Text = "Lost" Then
            changeBet (bet * 2)
        ElseIf winCheckRange.Text = "Win" Or _
            winCheckRange.Text = "BLACKJACK" Then
            changeBet (initialBet)
        End If
    Next i
End Sub
Sub cardCounter()
    Dim i As Integer
    For i = initialDeckIndex To deckIndex - 1
        If deck(i) = "A" Then
            numberOfCards(0) = numberOfCards(0) - 1
        ElseIf deck(i) = "2" Then
            numberOfCards(1) = numberOfCards(1) - 1
        ElseIf deck(i) = "3" Then
            numberOfCards(2) = numberOfCards(2) - 1
        ElseIf deck(i) = "4" Then
            numberOfCards(3) = numberOfCards(3) - 1
        ElseIf deck(i) = "5" Then
            numberOfCards(4) = numberOfCards(4) - 1
        ElseIf deck(i) = "6" Then
            numberOfCards(5) = numberOfCards(5) - 1
        ElseIf deck(i) = "7" Then
            numberOfCards(6) = numberOfCards(6) - 1
        ElseIf deck(i) = "8" Then
            numberOfCards(7) = numberOfCards(7) - 1
        ElseIf deck(i) = "9" Then
            numberOfCards(8) = numberOfCards(8) - 1
        ElseIf deck(i) = "10" Then
            numberOfCards(9) = numberOfCards(9) - 1
        ElseIf deck(i) = "J" Then
            numberOfCards(10) = numberOfCards(10) - 1
        ElseIf deck(i) = "Q" Then
            numberOfCards(11) = numberOfCards(11) - 1
        ElseIf deck(i) = "K" Then
            numberOfCards(12) = numberOfCards(12) - 1
        End If
    Next i
    Call hitOddGenerator
End Sub
Sub gimmePossibilities()
    Sheet1.Range("K2").Value = calculateStand(playerValue, dealerValue)
    Sheet1.Range("K3").Value = basicHitOddCalculator(playerValue)
End Sub
Sub hitOddGenerator()
    Dim i As Integer
    For i = 0 To 12
        playerOdd(i) = numberOfCards(i) / (312 - initialDeckIndex)
    Next i
End Sub
Function basicHitOddCalculator(playerValue) As Double
    Dim i As Integer
    If playerValue = 21 Then
        basicHitOddCalculator = 0
    ElseIf playerValue > 11 Then
        For i = playerValue To 20
            basicHitOddCalculator = basicHitOddCalculator + playerOdd(20 - i)
        Next i
    ElseIf playerValue <= 11 Then
        basicHitOddCalculator = 1
    End If
    
End Function
'Function basicHitOrNot(playerValue, dealerValue)
'
'    If playerValue > 11 Then
'        Dim temp As Double
'        temp = calculateStand(playerValue, dealerValue)
'        If basicHitOddCalculator(playerValue) >= calculateStand(playerValue, dealerValue) Then
'
'
'        End If
'
'    End If
'End Function

Function calculateStand(p, d) As Double
    If p < 17 Then
        calculateStand = dealerBust(d, 6)
    ElseIf p = 17 Then
        calculateStand = dealerBust(d, 6) + dealerBust(d, 1)
    ElseIf p = 18 Then
        calculateStand = dealerBust(d, 6) + dealerBust(d, 1) + dealerBust(d, 2)
    ElseIf p = 19 Then
        calculateStand = dealerBust(d, 6) + dealerBust(d, 1) + _
            dealerBust(d, 2) + dealerBust(d, 3)
    ElseIf p = 20 Then
        calculateStand = dealerBust(d, 6) + dealerBust(d, 1) + _
            dealerBust(d, 2) + dealerBust(d, 3) + dealerBust(d, 4)
    ElseIf p = 21 Then
        calculateStand = dealerBust(d, 6) + dealerBust(d, 1) + _
            dealerBust(d, 2) + dealerBust(d, 3) + dealerBust(d, 4) + dealerBust(d, 5)
    Else
        calculateStand = 0
    End If
    
End Function

Sub newGame()
    Call cleanTable
    Set dealerHand = Sheet1.Range("C6")
    Set playerHand = Sheet1.Range("D6")
    Set winCheckRange = Sheet1.Range("E6")
    Set moneyRange = Sheet1.Range("F6")
    Set regularBet = Sheet1.Range("B6")
    Set handNumber = Sheet1.Range("A6")
    money = InputBox("What is the initial money? ")
    bet = InputBox("What is the regular bet? ")
    deckIndex = 0
    handIndex = 0
    Call initDeck
    Call initCardNumbers
    Call initPossibilities
    Call cardCounter
    initialDeckIndex = deckIndex
    Call hitOddGenerator
    
    
End Sub
Sub initCardNumbers()
    Dim j As Integer
    For j = 0 To 12
        numberOfCards(j) = 24
    Next j
End Sub
Sub cleanTable()
    Application.ScreenUpdating = False
    
    Sheet1.Range("A7").Select
    Range(Selection.End(xlDown), Selection.End(xlToRight)).Select
    With Selection
        .ClearContents
        .Interior.ColorIndex = 0
        
    End With
    
    Range("K2:K3").ClearContents
    Range("A1").Select
    
End Sub
Sub initDeck()
    Dim i As Integer
    
    For i = 0 To 311
        If (i Mod 13) < 11 And (i Mod 13) > 1 Then
            deck(i) = i Mod 13
        ElseIf (i Mod 13) = 11 Then
            deck(i) = "J"
        ElseIf (i Mod 13) = 12 Then
            deck(i) = "Q"
        ElseIf (i Mod 13) = 0 Then
            deck(i) = "K"
        ElseIf (i Mod 13) = 1 Then
            deck(i) = "A"
        End If
    Next i
    initialDeckIndex = 0
    deckIndex = 0
    Call shuffler
    
End Sub
Sub shuffler()
    Dim randomArray(311) As Double
    Dim i As Integer
    Dim j As Integer
    Dim rand As Double
    Dim temp As Double
    Dim temp2 As String
    
    For i = 0 To 311
        rand = Rnd() * 312
        randomArray(i) = rand
    Next i
    
    For i = 0 To 311
        For j = 0 To 311
            If randomArray(i) > randomArray(j) Then
                temp = randomArray(i)
                randomArray(i) = randomArray(j)
                randomArray(j) = temp
                temp2 = deck(i)
                deck(i) = deck(j)
                deck(j) = temp2
            End If
        Next j
    Next i
    
End Sub
Public Function hit() As String
    If deckIndex = 0 Then
        Call initDeck
    End If
    
    hit = deck(deckIndex)
    deckIndex = deckIndex + 1
    
    If deckIndex = 207 Then '207 random joker card olucak
        'MsgBox "New suffeled deck has been created"
        Call initDeck
        Call initCardNumbers
        Call initPossibilities
        Call cardCounter
        initialDeckIndex = deckIndex
        Call hitOddGenerator
    End If
    
    Range("K2:K3").ClearContents
        
End Function
Sub dealClicked()
    Set dealerHand = dealerHand.Offset(1, 0)
    Set playerHand = playerHand.Offset(1, 0)
    Set winCheckRange = winCheckRange.Offset(1, 0)
    Set regularBet = regularBet.Offset(1, 0)
    regularBet = bet
    Set handNumber = handNumber.Offset(1, 0)
    Set moneyRange = moneyRange.Offset(1, 0)
    handIndex = handIndex + 1
    winCheckRange.Interior.ColorIndex = 0
    winCheckRange = ""
    handNumber = handIndex
    dealerHand = hit() & ","
    dealerValue = valueCalc(dealerHand.Text)
    dealer2Card = hit()
    dealerHand = dealerHand & "?"
    playerHand = hit() & "," & hit()
    playerValue = valueCalc(playerHand.Text)
    If playerValue = 21 Then
        winCheckRange = "BLACKJACK"
        winCheckRange.Interior.ColorIndex = 30
        winCheckRange.Font.ColorIndex = 2
        money = money + bet * 1.5
        moneyRange = money
    End If
    Range("K2:K3").ClearContents
    Call cardCounter
    initialDeckIndex = deckIndex
    Call hitOddGenerator
    Call gimmePossibilities
End Sub

Sub hitClicked()
    If Not winCheckRange = "" Then
        MsgBox "It's finished, you cannot hit anymore!"
        Call dealClicked
    End If
    playerHand = playerHand & "," & hit()
    playerValue = valueCalc(playerHand.Text)
    
    If playerValue > 21 Then
      Call displayWinner
    End If
    
    Range("K2:K3").ClearContents
    Call cardCounter
    initialDeckIndex = deckIndex
    Call hitOddGenerator
    Call gimmePossibilities
End Sub
Sub standClicked()
    playerValue = valueCalc(playerHand.Text)
    Call playDealer
    
    Range("K2:K3").ClearContents
    
End Sub
Sub changeBet(newBet As Integer)
    bet = newBet
End Sub
Sub playDealer()
    If Left(dealerHand.Text, 2) = "10" Then
        dealerHand = Left(dealerHand.Text, 3) & dealer2Card
    Else
        dealerHand = Left(dealerHand.Text, 2) & dealer2Card
    End If

    dealerValue = valueCalc(dealerHand)
    
    If dealerValue < 22 Then
        Do Until dealerValue > 16
        dealerHand = dealerHand & "," & hit()
        dealerValue = valueCalc(dealerHand)
        Loop
    End If
    Call displayWinner
    
End Sub
Sub displayWinner()
       
    If dealerValue > playerValue And dealerValue < 22 Then
        winCheckRange = "Lost"
        winCheckRange.Interior.ColorIndex = 3
        money = money - bet
        moneyRange = money
    ElseIf dealerValue < playerValue And playerValue < 22 Then
        winCheckRange = "Win"
        winCheckRange.Interior.ColorIndex = 4
        money = money + bet
        moneyRange = money
     ElseIf playerValue > 21 Then
        winCheckRange = "Lost"
        winCheckRange.Interior.ColorIndex = 3
        money = money - bet
        moneyRange = money
    ElseIf dealerValue > 21 Then
        winCheckRange = "Win"
        winCheckRange.Interior.ColorIndex = 4
        money = money + bet
        moneyRange = money
    ElseIf dealerValue = playerValue Then
        winCheckRange = "Draw"
        winCheckRange.Interior.ColorIndex = 6
        moneyRange = money
    End If
    
    
End Sub

Public Function valueCalc(Hand) As Integer
    
    Dim i As Integer
    Dim handArray() As String
    ReDim handArray(Len(Hand) - Len(WorksheetFunction.Substitute(Hand, ",", ""))) As String
    handArray = Split(Hand, ",")
    valueCalc = 0
    
    For i = 0 To Len(Hand) - Len(WorksheetFunction.Substitute(Hand, ",", ""))
        If handArray(i) = "1" Then
            valueCalc = valueCalc + 1
        ElseIf handArray(i) = "2" Then
            valueCalc = valueCalc + 2
        ElseIf handArray(i) = "3" Then
            valueCalc = valueCalc + 3
        ElseIf handArray(i) = "4" Then
            valueCalc = valueCalc + 4
        ElseIf handArray(i) = "5" Then
            valueCalc = valueCalc + 5
        ElseIf handArray(i) = "6" Then
            valueCalc = valueCalc + 6
        ElseIf handArray(i) = "7" Then
            valueCalc = valueCalc + 7
        ElseIf handArray(i) = "8" Then
            valueCalc = valueCalc + 8
        ElseIf handArray(i) = "9" Then
            valueCalc = valueCalc + 9
        ElseIf handArray(i) = "10" Then
            valueCalc = valueCalc + 10
        ElseIf handArray(i) = "J" Then
            valueCalc = valueCalc + 10
        ElseIf handArray(i) = "Q" Then
            valueCalc = valueCalc + 10
        ElseIf handArray(i) = "K" Then
            valueCalc = valueCalc + 10
        ElseIf handArray(i) = "A" Then
        End If
    Next i
    
    For i = 0 To Len(Hand) - Len(WorksheetFunction.Substitute(Hand, ",", ""))
        If handArray(i) = "A" Then
            valueCalc = valueCalcAce(Hand)
            Exit For
        End If
    Next i
    
End Function

Function valueCalcAce(Hand)
    Dim i As Integer
    Dim handArray() As String
    ReDim handArray(Len(Hand) - Len(WorksheetFunction.Substitute(Hand, ",", ""))) As String
    handArray = Split(Hand, ",")
    valueCalcAce = 0
    
    For i = 0 To Len(Hand) - Len(WorksheetFunction.Substitute(Hand, ",", ""))
        If handArray(i) = "1" Then
            valueCalcAce = valueCalcAce + 1
        ElseIf handArray(i) = "2" Then
            valueCalcAce = valueCalcAce + 2
        ElseIf handArray(i) = "3" Then
            valueCalcAce = valueCalcAce + 3
        ElseIf handArray(i) = "4" Then
            valueCalcAce = valueCalcAce + 4
        ElseIf handArray(i) = "5" Then
            valueCalcAce = valueCalcAce + 5
        ElseIf handArray(i) = "6" Then
            valueCalcAce = valueCalcAce + 6
        ElseIf handArray(i) = "7" Then
            valueCalcAce = valueCalcAce + 7
        ElseIf handArray(i) = "8" Then
            valueCalcAce = valueCalcAce + 8
        ElseIf handArray(i) = "9" Then
            valueCalcAce = valueCalcAce + 9
        ElseIf handArray(i) = "10" Then
            valueCalcAce = valueCalcAce + 10
        ElseIf handArray(i) = "J" Then
            valueCalcAce = valueCalcAce + 10
        ElseIf handArray(i) = "Q" Then
            valueCalcAce = valueCalcAce + 10
        ElseIf handArray(i) = "K" Then
            valueCalcAce = valueCalcAce + 10
        ElseIf handArray(i) = "A" Then
            If valueCalcAce > 10 Then
                valueCalcAce = valueCalcAce + 1
            Else
                valueCalcAce = valueCalcAce + 11
            End If
        End If
    Next i
    
    If valueCalcAce > 21 Then

        Hand = Reverse(Hand)
        ReDim handArray(Len(Hand) - Len(WorksheetFunction.Substitute(Hand, ",", ""))) As String
        handArray = Split(Hand, ",")
        valueCalcAce = 0
    
        For i = 0 To Len(Hand) - Len(WorksheetFunction.Substitute(Hand, ",", ""))
            If handArray(i) = "1" Then
                valueCalcAce = valueCalcAce + 1
            ElseIf handArray(i) = "2" Then
                valueCalcAce = valueCalcAce + 2
            ElseIf handArray(i) = "3" Then
                valueCalcAce = valueCalcAce + 3
            ElseIf handArray(i) = "4" Then
                valueCalcAce = valueCalcAce + 4
            ElseIf handArray(i) = "5" Then
                valueCalcAce = valueCalcAce + 5
            ElseIf handArray(i) = "6" Then
                valueCalcAce = valueCalcAce + 6
            ElseIf handArray(i) = "7" Then
                valueCalcAce = valueCalcAce + 7
            ElseIf handArray(i) = "8" Then
                valueCalcAce = valueCalcAce + 8
            ElseIf handArray(i) = "9" Then
                valueCalcAce = valueCalcAce + 9
            ElseIf handArray(i) = "10" Then
                valueCalcAce = valueCalcAce + 10
            ElseIf handArray(i) = "J" Then
                valueCalcAce = valueCalcAce + 10
            ElseIf handArray(i) = "Q" Then
                valueCalcAce = valueCalcAce + 10
            ElseIf handArray(i) = "K" Then
                valueCalcAce = valueCalcAce + 10
            ElseIf handArray(i) = "A" Then
                If valueCalcAce > 10 Then
                    valueCalcAce = valueCalcAce + 1
                Else
                    valueCalcAce = valueCalcAce + 11
                End If
            End If
        Next i
    End If
End Function
Function Reverse(Text) As String
    Dim i As Integer
    Dim StrNew As String
    Dim strOld As String
    strOld = Trim(Text)
    For i = 1 To Len(strOld)
      StrNew = Mid(strOld, i, 1) & StrNew
    Next i
    Reverse = StrNew
End Function
Sub initPossibilities()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To 13
        For j = 1 To 6
            dealerBust(i, j) = Sheet2.Range("F4").Offset(i - 1, j - 1)
        Next j
    Next i
End Sub



Attribute VB_Name = "MetagameStuff"
Option Explicit

Public Function ClearMetaData()
    Dim shtMeta As Worksheet
    Set shtMeta = Sheets("Meta")
    
    ' Clear current meta counts
    Range(shtMeta.Cells(rMetaCounts + 1, cMetaCounts), _
              shtMeta.Cells(rMetaCounts + nDecksPerClass, cMetaCounts + 2 * nClasses - 1)).Value2 = 0
End Function
    
Public Function LoadMetaFromLogs()
    Dim shtLog, shtMeta As Worksheet
    Set shtMeta = Sheets("Meta")
    Set shtLog = Sheets("Log")
    
    Dim classWins(nClasses, nClasses) As Integer
    Dim classLosses(nClasses, nClasses) As Integer
    
    ' Count how many rows of data there are
    Dim maxI, i, initValue As Integer
    maxI = rLogs
    While shtLog.Cells(maxI, cLogs_Date).Value2 <> ""
            maxI = maxI + 1
    Wend
    maxI = maxI - 1
    
    ' Don't read more games than the configured max # of games; This must be checked in the loop because of filters
    Dim configMax As Integer
    configMax = CInt(shtMeta.Cells(rMetaCfg_MaxGames, cMetaCfg).Value2) + rLogs - 1
    
    Dim myDeck, oppDeck As Deck
    Dim gamesCounted As Integer
    gamesCounted = 0
    For i = maxI To rLogs Step -1
        If Not ValidateRowForMeta(i) Then GoTo NextI

        Set myDeck = Factory.CreateDeck(shtLog.Cells(i, cLogs_MyDeck).Value2)
        Set oppDeck = Factory.CreateDeck(shtLog.Cells(i, cLogs_OppDeck).Value2)
        If Not myDeck.IsValid Or Not oppDeck.IsValid Then
            Range(shtLog.Cells(i, cLogs_Date), shtLog.Cells(i, cLogs_Notes)).Interior.Color = vbRed
            GoTo NextI
        End If

        ' Increment counts for opposing deck (don't count the player's deck for determining meta)
        initValue = CInt(shtMeta.Cells(rMetaCounts + oppDeck.rowOffset, cMetaCounts + 2 * oppDeck.colOffset).Value2)
        shtMeta.Cells(rMetaCounts + oppDeck.rowOffset, cMetaCounts + 2 * oppDeck.colOffset).Value2 = initValue + 1
        Range(shtLog.Cells(i, cLogs_Date), shtLog.Cells(i, cLogs_Notes)).Interior.Color = RGB(200, 200, 200) ' Light gray
        gamesCounted = gamesCounted + 1
        
        ' Consider win/loss for class vs class calculation
        Dim wonGame As Boolean
        wonGame = shtLog.Cells(i, cLogs_Won).Value2 = 1 Or _
                            StartsWith(LCase(shtLog.Cells(i, cLogs_Won).Value2), "t") Or _
                            StartsWith(LCase(shtLog.Cells(i, cLogs_Won).Value2), "y")
        If wonGame Then
            classWins(1 + myDeck.colOffset, 1 + oppDeck.colOffset) = classWins(1 + myDeck.colOffset, 1 + oppDeck.colOffset) + 1
            classLosses(1 + oppDeck.colOffset, 1 + myDeck.colOffset) = classLosses(1 + oppDeck.colOffset, 1 + myDeck.colOffset) + 1
        Else
            classLosses(1 + myDeck.colOffset, 1 + oppDeck.colOffset) = classLosses(1 + myDeck.colOffset, 1 + oppDeck.colOffset) + 1
            classWins(1 + oppDeck.colOffset, 1 + myDeck.colOffset) = classWins(1 + oppDeck.colOffset, 1 + myDeck.colOffset) + 1
        End If
        
        If gamesCounted = configMax Then Exit For
NextI:
    Next i
    
    ' Call this here because it's quickest to sum these as we are going through the logs anyway
    UpdateClassMatchups classWins, classLosses
    
End Function

Public Function ValidateRowForMeta(ByVal Row As Integer)
    Dim shtMeta As Worksheet
    Dim shtLog As Worksheet
    Set shtMeta = Sheets("Meta")
    Set shtLog = Sheets("Log")

    ValidateRowForMeta = True
    
    ' Check for date out of bounds, or a note of "repeat" to signify a requeue that shouldn't affect the meta calculations
    If CDate(shtLog.Cells(Row, cLogs_Date).Value2) < CDate(shtMeta.Cells(rMetaCfg_MinDate, cMetaCfg).Value2) Or _
        CDate(shtLog.Cells(Row, cLogs_Date).Value2) > CDate(shtMeta.Cells(rMetaCfg_MaxDate, cMetaCfg).Value2) Or _
        LCase(shtLog.Cells(Row, cLogs_Notes).Value2) = "repeat" Then
            ValidateRowForMeta = False
    Else
        ' If rank data is missing, do not exclude the game
        If shtLog.Cells(Row, cLogs_MyRank).Value2 <> "" Then
            If CInt(shtLog.Cells(Row, cLogs_MyRank).Value2) > CInt(shtMeta.Cells(rMetaCfg_MyMinRank, cMetaCfg).Value2) Or _
                CInt(shtLog.Cells(Row, cLogs_MyRank).Value2) < CInt(shtMeta.Cells(rMetaCfg_MyMaxRank, cMetaCfg).Value2) Then _
                    ValidateRowForMeta = False
        End If
        
        ' Check that row is still valid first to skip the computation when possible
        If ValidateRowForMeta And shtLog.Cells(Row, cLogs_OppRank).Value2 <> "" Then
            If CInt(shtLog.Cells(Row, cLogs_OppRank).Value2) > CInt(shtMeta.Cells(rMetaCfg_OppMinRank, cMetaCfg).Value2) Or _
                CInt(shtLog.Cells(Row, cLogs_OppRank).Value2) < CInt(shtMeta.Cells(rMetaCfg_OppMaxRank, cMetaCfg).Value2) Then _
                    ValidateRowForMeta = False
        End If
    End If
End Function

Public Function UpdateClassMatchups(ByRef classWins() As Integer, ByRef classLosses() As Integer)
    Dim shtMeta As Worksheet
    Set shtMeta = Sheets("Meta")
    Dim i, j As Integer
    Dim perc As String
    
    For i = 1 To nClasses
        For j = 1 To nClasses
            If i <> j Then
                If (classWins(i, j) + classLosses(i, j)) = 0 Then
                    perc = ""
                Else
                    perc = CDbl(classWins(i, j)) / CDbl(classWins(i, j) + classLosses(i, j))
                End If
                shtMeta.Cells(rMetaClassMatchups + i - 1, cMetaClassMatchups + j - 1) = perc
            End If
        Next
    Next
End Function

Public Function ClearClassMatchups()
    Dim shtMeta As Worksheet
    Set shtMeta = Sheets("Meta")
    Dim i, j As Integer
    For i = 1 To nClasses
        For j = 1 To nClasses
            If i = j Then
                shtMeta.Cells(rMetaClassMatchups + i - 1, cMetaClassMatchups + j - 1).Value2 = "---"
            Else
                shtMeta.Cells(rMetaClassMatchups + i - 1, cMetaClassMatchups + j - 1).Value2 = ""
            End If
        Next
    Next
End Function

Public Function UpdateMostPlayedClasses()
    Dim shtMeta
    Set shtMeta = Sheets("Meta")
    
    ' No need to clear data here because it will all be overwritten every time (there are always 9 classes!)
    
    Dim curMax, nextMax, curVal As Double
    Dim curRow, i As Integer
    curMax = 1#
    curRow = rMetaMPC
    While curRow <= rMetaMPC + nClasses - 1
        nextMax = 0#
        For i = 1 To nClasses
            curVal = shtMeta.Cells(rMetaPerc_Class, cMetaPerc + 2 * (i - 1))
            If curVal = curMax Then
                shtMeta.Cells(curRow, cMetaMPC_Name).Value2 = shtMeta.Cells(rMetaPerc_ClassName, cMetaPerc + 2 * (i - 1)).Value2
                shtMeta.Cells(curRow, cMetaMPC_Value).Value2 = curVal
                curRow = curRow + 1
            ElseIf curVal > nextMax And curVal < curMax Then
                nextMax = curVal
            End If
        Next
        curMax = nextMax
    Wend
End Function

Public Function UpdateMostPlayedDecks()
    Dim shtMeta
    Set shtMeta = Sheets("Meta")
    
    ' Clear current data
    shtMeta.Range("U3:U27").Value = ""
    shtMeta.Range("V3:V27").Value = ""
    
    ' Use currentMax to find the top decks in order
    Dim i, maxI, rowOffset, colOffset As Integer
    Dim currentMax, nextMax, thisVal As Double
    currentMax = 1#
    i = rMMeta
    
    ' Determine number of decks that will be listed
    maxI = i - 1 + Application.WorksheetFunction.CountIf(Range( _
        shtMeta.Cells(rMetaPerc_Deck, cMetaPerc), _
        shtMeta.Cells(rMetaPerc_Deck + nDecksPerClass - 1, cMetaPerc + 2 * nClasses - 1)), ">0")
    If maxI > rMMeta + nMetaDecks - 1 Then maxI = rMMeta + nMetaDecks - 1

    While i <= maxI
        nextMax = 0#
        For rowOffset = 0 To nDecksPerClass - 1
            For colOffset = 0 To nClasses - 1
                If shtMeta.Cells(rMetaPerc_Deck + rowOffset, cMetaPerc + 2 * colOffset).Value2 > 0 And i <= maxI Then
                    ' Look for values equal to the current maximum
                    thisVal = CDbl(shtMeta.Cells(rMetaPerc_Deck + rowOffset, cMetaPerc + 2 * colOffset).Value2)
                    If thisVal = currentMax Then
                        ' Found it, write data and increment i
                        shtMeta.Cells(i, cMMeta_Name).Value2 = _
                            shtMeta.Cells(rMeta_DeckName + 1 + rowOffset, cMeta_DeckName + 2 * colOffset).Value2 _
                            & " " & shtMeta.Cells(rMeta_DeckName, cMeta_DeckName + 2 * colOffset).Value2
                        shtMeta.Cells(i, cMMeta_Value).Value2 = thisVal
                        i = i + 1
                    ElseIf thisVal > nextMax And thisVal < currentMax Then
                        ' Determine the maximum for the next iteration
                        nextMax = thisVal
                    End If
                End If
            Next
        Next
        currentMax = nextMax
    Wend
End Function

Public Function UpdateBestMetaDecks()
    Dim decks As New Collection
    Dim rates As New Collection
    Dim shtMeta, ws As Worksheet
    Set shtMeta = Sheets("Meta")

    ' Clear old data
    Range(shtMeta.Cells(rBestDecks, cBestDecks_Name), _
                shtMeta.Cells(rBestDecks + nMetaDecks - 2, cBestDecks_Value)).Value2 = ""
    
    For Each ws In Worksheets
        If IsDeckSheet(ws.Name) Then
            If ws.Cells(rDeckSheet_WinRate, cDeckSheet_WinRate).Value2 <> "NA" And _
                ws.Cells(rDeckSheet_GameCount, cDeckSheet_GameCount).Value2 >= shtMeta.Cells(rBestDecksCfg_MinGames, cBestDecksCfg_MinGames).Value2 Then
                decks.Add (ws.Name)
                rates.Add (ws.Cells(rDeckSheet_WinRate, cDeckSheet_WinRate))
            End If
        End If
    Next

    Dim curMax, nextMax, curVal As Double
    Dim numSlots, curDeckOffset, i As Integer
    numSlots = nMetaDecks
    If decks.count < numSlots Then numSlots = decks.count
    curDeckOffset = 0
    curMax = 1#
    While curDeckOffset < numSlots
        nextMax = 0#
        For i = 1 To decks.count
            curVal = rates(i)
            If curVal = curMax Then
                shtMeta.Cells(rBestDecks + curDeckOffset, cBestDecks_Name).Value2 = decks(i)
                shtMeta.Cells(rBestDecks + curDeckOffset, cBestDecks_Value).Value2 = curVal
                curDeckOffset = curDeckOffset + 1
            ElseIf curVal > nextMax And curVal < curMax Then
                nextMax = curVal
            End If
        Next
        curMax = nextMax
    Wend
End Function

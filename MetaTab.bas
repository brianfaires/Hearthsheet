Attribute VB_Name = "MetaTab"
Option Explicit

' Global settings to persist across numerous calls to ValidateRowForMeta()
Dim metaMaxGames As Integer, metaMyMinRank As Integer, metaMyMaxRank As Integer, metaOppMinRank As Integer, metaOppMaxRank As Integer
Dim metaMinDate As Date, metaMaxDate

Public Function LoadMetaSettings()
    With Meta
        metaMinDate = CDate(.Range("MinDate").Value2)
        metaMaxDate = CDate(.Range("MaxDate").Value2)
        metaMaxGames = .Range("MaxGames").Value2
        metaMyMinRank = .Range("MyMinRank").Value2
        metaMyMaxRank = .Range("MyMaxRank").Value2
        metaOppMinRank = .Range("OppMinRank").Value2
        metaOppMaxRank = .Range("OppMaxRank").Value2
    End With
End Function

' Clears table prior to writing fresh data
Public Function ClearMetaData()
    Meta.Range("WeightTable").Value2 = 0
End Function

' Read logs, skipping games as necessary due to filters. Write data to Meta sheet.
Public Function LoadMetaFromLogs()
    Dim curRow As Range
    Dim curDeck As Deck
    Dim i As Integer, gamesCounted As Integer
    gamesCounted = 0
    
    LoadMetaSettings
    
    
    With Meta.Range("WeightTable")
        ' Clear old color coding
        ' TODO: this will overwrite past red cells, and only the ones encountered will be marked.
        Log.Range("AllLogs").Interior.ColorIndex = xlNone
        
        For Each curRow In Log.Range("AllLogs").Rows
            If curRow.Cells(LogCfgCol_Date) = "" Then
                Exit For
            ElseIf ValidateRowForMeta(curRow) Then
                Set curDeck = Factory.CreateDeck(curRow.Cells(LogCfgCol_OppDeck).Value2)
                If curDeck.IsValid Then
                    ' Increment the meta count
                    .Cells(curDeck.rowOffset, 2 * curDeck.colOffset - 1).Value2 = _
                    .Cells(curDeck.rowOffset, 2 * curDeck.colOffset - 1).Value2 + 1
                    curRow.Interior.Color = RGB(200, 200, 200) ' Light gray
                    
                    gamesCounted = gamesCounted + 1
                    If gamesCounted >= metaMaxGames Then Exit For
                Else
                    curRow.Interior.Color = vbRed
                End If
            End If
        Next curRow
    End With
End Function

' Check that the current game should be included in this meta calculation
Public Function ValidateRowForMeta(curRow As Range)
    ValidateRowForMeta = True
    
    Dim curDate As Date
    Dim curStr As String
    Dim curRank As Integer
    
    ' Check for date out of bounds, or a note of "repeat" to signify a requeue that shouldn't affect the meta calculations
    curDate = CDate(curRow.Cells(LogCfgCol_Date).Value2)
    If curDate < metaMinDate Or curDate > metaMaxDate Or LCase(curRow.Cells(LogCfgCol_Notes).Value2) = "repeat" Then
            ValidateRowForMeta = False
    Else
        ' If rank data is missing, do not exclude the game
        curStr = curRow.Cells(LogCfgCol_MyRank).Value2
        If curStr <> "" Then
            curRank = CInt(curStr)
            If curRank > metaMyMinRank Or curRank < metaMyMaxRank Then ValidateRowForMeta = False
        End If
        
        ' Check that row is still valid first to skip computations when possible
        If ValidateRowForMeta Then
            curStr = curRow.Cells(LogCfgCol_OppRank).Value2
            If curStr <> "" Then
                curRank = CInt(curStr)
                If curRank > metaOppMinRank Or curRank < metaOppMaxRank Then ValidateRowForMeta = False
            End If
        End If
    End If
End Function

' Populates the most played classes table on the Meta sheet
Public Function UpdateMostPlayedClasses()
    Dim classWinRates As Range, MPClasses_Names As Range, MPClasses_Values As Range
    With Meta
        Set classWinRates = .Range("ClassPerc")
        Set MPClasses_Names = .Range("MPClasses_Names")
        Set MPClasses_Values = .Range("MPClasses_Values")
    End With
    
    ' No need to clear data here because it will all be overwritten every time (there are always 9 classes)
    
    ' Iterate up to 9 times, taking the highest remaining class win rate(s) each time
    Dim curCell As Range
    Dim i As Integer
    Dim curMax As Double, nextMax As Double, curVal As Double
    curMax = 1#
    i = 1
    
    While i <= nClasses
        nextMax = 0#
        For Each curCell In classWinRates
            If curCell.Value2 <> "" Then
                curVal = CDbl(curCell.Value2)
                If curVal = curMax Then
                    MPClasses_Names(i).Value2 = curCell.Offset(-1 * nDecksPerClass - 1).Value2 ' Get the class name from the head of this table
                    MPClasses_Values(i).Value2 = curVal
                    i = i + 1
                ElseIf curVal > nextMax And curVal < curMax Then
                    nextMax = curVal
                End If
            End If
        Next curCell
        curMax = nextMax
    Wend
End Function

' Update the (master copy of) most played decks in this meta
Public Function UpdateMostPlayedDecks()
    Dim classNamesRow As Integer
    Dim MPD_Names As Range, MPD_Values As Range, metaCounts
    Set MPD_Names = Meta.Range("MPDecks_Names")
    Set MPD_Values = Meta.Range("MPDecks_Values")
    Set metaCounts = Meta.Range("PercTable")
    
    classNamesRow = Meta.Range("ClassNames").Row
    
    ' Clear current data
    MPD_Names.Value2 = ""
    MPD_Values.Value2 = ""
    
    ' Determine number of decks that will be listed
    Dim curCell As Range
    Dim numDecks As Integer
    numDecks = WorksheetFunction.CountIf(metaCounts, ">0")
    If numDecks > nMetaDecks Then numDecks = nMetaDecks
    
    ' Iterate up to numDecks times, each time adding the deck(s) with the highest win rates
    Dim i As Integer
    Dim curMax As Double, nextMax As Double, curVal As Double
    curMax = 1#
    i = 1
    
    While i <= numDecks
        nextMax = 0#
        For Each curCell In metaCounts
            If i <= numDecks And curCell.Value2 > 0 Then
                'Look for values equal to the current maximum
                curVal = CDbl(curCell.Value2)
                If curVal = curMax Then
                    ' Found it, write data and increment i
                    
                    'Get deckName from indexes of curCell
                    Dim curDeckName As String, curClassName As String
                    curDeckName = curCell.Offset(-19).Value2 ' 19 rows between deck names and deck percents; TODO: Unhack this
                    curClassName = Meta.Cells(classNamesRow, curCell.Column)
                    
                    MPD_Names.Cells(i).Value2 = curDeckName & " " & curClassName
                    MPD_Values.Cells(i).Value2 = curVal
                    i = i + 1
                ElseIf curVal > nextMax And curVal < curMax Then
                        ' Determine the maximum for the next iteration
                        nextMax = curVal
                    End If
            End If
        Next curCell

        curMax = nextMax
    Wend
End Function

'Determine each deck's expected win rate in this meta, and write to the table on the Meta sheet
' TODO: This only includes decks with current deck sheets. Could manually recalculate the meta win rates instead.
Public Function UpdateBestMetaDecks()
    Dim decks As New Collection
    Dim values As New Collection
    Dim ws As Worksheet
    Dim bestDecksNames As Range, bestDecksValues As Range
    Dim curWinRate As Double
    
    Set bestDecksNames = Meta.Range("BestDecks_Names")
    Set bestDecksValues = Meta.Range("BestDecks_Values")
    
    ' Clear old data
    bestDecksNames.Value2 = ""
    bestDecksValues.Value2 = ""
    
    For Each ws In Worksheets
        If IsDeckSheet(ws.Name) Then
            ' TODO: Check for curWinRate = "NA" ?
            curWinRate = ws.Range("MetaWinRate").Value2
            If ws.Range("TotalGames").Value2 >= Meta.Range("MinGames").Value2 Then
                decks.Add (ws.Name)
                values.Add (curWinRate)
            End If
        End If
    Next

    ' Now iterate through list, taking highest value(s) each time and adding to the Best Decks list on Meta sheet
    Dim curMax, nextMax, curVal As Double
    Dim numSlots, curDeckOffset, i As Integer
    numSlots = nMetaDecks
    If decks.count < numSlots Then numSlots = decks.count
    curDeckOffset = 1
    curMax = 1#
    While curDeckOffset <= numSlots
        nextMax = 0#
        For i = 1 To decks.count
            curVal = values(i)
            If curVal = curMax Then
                bestDecksNames.Cells(curDeckOffset).Value2 = decks(i)
                bestDecksValues.Cells(curDeckOffset).Value2 = values(i)
                curDeckOffset = curDeckOffset + 1
            ElseIf curVal > nextMax And curVal < curMax Then
                nextMax = curVal
            End If
        Next
        curMax = nextMax
    Wend
End Function

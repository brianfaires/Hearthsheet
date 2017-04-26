Attribute VB_Name = "LogsTab"
' LogStuff Module: Handles Log processing and deck sheet generation
Option Explicit

' Settings as globals to persist across function calls that process each row
Dim minDate As Date
Dim myMinRank As Integer, myMaxRank As Integer, oppMinRank As Integer, oppMaxRank As Integer

Public Function ReadSettings()
    With Log
        minDate = CDate(.Range("StartDate").Value2)
        myMinRank = CInt(.Range("MyMinRank").Value2)
        myMaxRank = CInt(.Range("MyMaxRank").Value2)
        oppMinRank = CInt(.Range("OppMinRank").Value2)
        oppMaxRank = CInt(.Range("OppMaxRank").Value2)
    End With
End Function

Public Function ClearDeckSheetGames()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsDeckSheet(ws.Name) Then
            ws.Range("WL_Table").Value2 = 0
        End If
    Next
End Function

' For easy un/protecting of all deck sheets
Public Function SetProtectionAllDeckSheets(protection As Boolean)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsDeckSheet(ws.Name) Then
            If protection Then
                ws.Protect
            Else
                ws.Unprotect
            End If
        End If
    Next
End Function

' Creates deck sheets for each deck, some will be removed in CleanUpDeckSheets()
Public Function CreateMissingDeckSheets()
    ' Allow edits
    Template.Visible = True
    Template.Unprotect
    
    Dim curClass As String, curDeck As String, fullName As String
    Dim classRange As Range, deckRange As Range
    Set classRange = Log.Range("Classes")
    Set deckRange = Log.Range("Decks")
    
    Dim i As Integer, j As Integer
    For i = 1 To nClasses
        curClass = classRange.Cells(1, i).Value2
        
        For j = 1 To nDecksPerClass
            curDeck = deckRange.Cells(j, i).Value2
            
            If curDeck <> "" Then
                fullName = curDeck & " " & curClass
                
                If Not SheetExists(fullName) Then
                    'Create a new Deck sheet, move it to the end
                    Template.Copy After:=Sheets(Sheets.count)
                    ActiveSheet.Name = fullName
                End If
            End If
        Next j
    Next i
    
    ' Disable edits
    Template.Protect
    Template.Visible = False
End Function

' Reads each log, creates a game object, and writes the data to the deck sheets
Public Function ProcessLogs()
    Dim timeRange As Range, myDeckRange As Range, myRankRange As Range, oppDeckRange As Range
    Dim oppRankRange As Range, winRange As Range, notesRange As Range, allLogsRange As Range

    With Log
        Set timeRange = .Range("Timestamp")
        Set myDeckRange = .Range("myDeck")
        Set myRankRange = .Range("myRank")
        Set oppDeckRange = .Range("oppDeck")
        Set oppRankRange = .Range("oppRank")
        Set winRange = .Range("Win")
        Set notesRange = .Range("Notes")
        Set allLogsRange = .Range("AllLogs")
    End With
        
    Dim i As Long
    Dim g As Game
    For i = 1 To allLogsRange.Rows.count
        'Read until an empty timestamp
        If timeRange(i).Value2 = "" Then Exit For
        
        If Not PassesFilters(allLogsRange.Rows(i)) Then
            ' Strikethrough on invalid row
            allLogsRange.Rows(i).Font.StrikeThrough = True
        Else
            allLogsRange.Rows(i).Font.StrikeThrough = False
            
            Set g = Factory.CreateGame(allLogsRange.Rows(i))
            If g.myDeck.IsValid And g.oppDeck.IsValid Then
                ' Write data to Deck sheets
                g.WriteData
                allLogsRange.Rows(i).Interior.Color = xlNone
            Else
                ' Red text on non-existant Deck names
                allLogsRange.Rows(i).Interior.Color = vbRed
            End If
        End If
    Next i
End Function

' Checks that specified row passes all relevant filters
Private Function PassesFilters(logRange As Range) As Boolean
    Dim tempRank As String
    Dim tempRankI As Integer
    
    PassesFilters = True
    If CDate(logRange.Cells(1, LogCfgCol_Date)) < minDate Then
        PassesFilters = False
    Else
        ' Use temp/tempI to avoid repeated calls to CInt(Range(...))
        tempRank = logRange.Cells(1, LogCfgCol_MyRank).Value2
        If tempRank <> "" Then
            ' If ranks are omitted, don't return false as a result
            tempRankI = CInt(tempRank)
            If tempRankI > myMinRank Or tempRankI < myMaxRank Then PassesFilters = False
        End If
        If PassesFilters Then
            tempRank = logRange.Cells(1, LogCfgCol_OppRank).Value2
            If tempRank <> "" Then
                tempRankI = CInt(tempRank)
                If tempRankI > oppMinRank Or tempRankI < oppMaxRank Then PassesFilters = False
            End If
        End If
    End If
End Function

' Sets colors/deletes sheets according to configured values
Public Function CleanUpDeckSheets()
    Dim curSheet As Worksheet
    
    ' Load settings from Log sheet
    Dim minGamesRed As Integer, minGamesYellow As Integer, minGamesBlack As Integer
    minGamesRed = CInt(Log.Range("MinGamesRed").Value2)
    minGamesYellow = CInt(Log.Range("MinGamesYellow").Value2)
    minGamesBlack = CInt(Log.Range("MinGamesBlack").Value2)
    
    ' Manually keep track of sheet count since xlCalculation will be turned off
    Dim shtCount As Integer, i As Integer, gameCount As Integer
    shtCount = Sheets.count
    i = 1
    While i <= shtCount
        Set curSheet = Sheets(i)
        If IsDeckSheet(curSheet.Name) Then
            gameCount = CInt(curSheet.Range("TotalGames").Value2)
            If gameCount < minGamesRed Then
                Application.DisplayAlerts = False
                curSheet.Delete
                Application.DisplayAlerts = True
                i = i - 1
                shtCount = shtCount - 1
            ElseIf gameCount < minGamesYellow Then
                curSheet.Tab.Color = vbRed
            ElseIf gameCount < minGamesBlack Then
                curSheet.Tab.Color = vbYellow
            Else
                curSheet.Tab.Color = vbBlack
            End If
        End If
        
        i = i + 1
    Wend

    SortDeckSheets
    'UpdateDefaultRates
    UpdateAllBestMatchups
End Function

' An insertion sort on the number of games played, taking care to skip over non-deck sheets.
Public Function SortDeckSheets()
    Dim i As Integer, j As Integer
    Dim doSwap As Boolean
    For i = 1 To Sheets.count
        If IsDeckSheet(Sheets(i).Name) Then
            j = i - 1
            doSwap = IsDeckSheet(Sheets(j).Name) And _
                            CInt(Sheets(i).Range("TotalGames").Value2) > CInt(Sheets(j).Range("TotalGames").Value2)
            
            While doSwap
                j = j - 1
                doSwap = IsDeckSheet(Sheets(j).Name) And _
                                CInt(Sheets(i).Range("TotalGames").Value2) > CInt(Sheets(j).Range("TotalGames").Value2)
            Wend
            
            Sheets(i).Move After:=Sheets(j)
        End If
    Next
End Function

' Replaced by formulas in the Deck sheets; put this back if its too much
'Sub UpdateDefaultRates()
'    For Each sht In sheets
'        If IsDeckSheet(sht.Name) Then
'            sht.Unprotect
'            tokens = Split(sht.Name, " ")
'            myClass = tokens(0)
'            myDeck = tokens(1)
'
'            sourceRow = 4 + 3 * GetColOffset(sht.Name) + GetRowOffset(sht.Name)
'            For classOffset = 0 To 8
'                For deckOffset = 0 To 5
'                    sourceCol = 4 + 6 * classOffset + deckOffset
'                    If Priors.Cells(3, sourceCol) <> "" Then
'                        defaultRate = 0.5
'                        priorValue = Priors.Cells(sourceRow, sourceCol)
'                        If priorValue <> "" Then defaultRate = CDbl(priorValue)
'
'                        sht.Cells(14 + deckOffset, 12 + 2 * classOffset) = defaultRate
'                    Else
'                        sht.Cells(14 + deckOffset, 12 + 2 * classOffset) = ""
'                    End If
'                Next
'            Next
'        End If
'    Next
'End Sub

Public Function UpdateAllBestMatchups()
    Dim sht As Worksheet
    For Each sht In Sheets
        If IsDeckSheet(sht.Name) Then
            ComputeBestMatchups sht
        End If
    Next
End Function

' Populates BestMatchups table on this sheet, which in turns populates the win rates for the current meta table.
' TODO: To speed up: Prepopulating Decks instead of parsing each one, maybe a more efficient way of looping through meta decks
Public Function ComputeBestMatchups(sheet As Worksheet)
    Dim metaNames As Range, matchupsNames, matchupsValues As Range, expWinRates As Range
    Set metaNames = sheet.Range("Meta_Names")
    Set matchupsNames = sheet.Range("BestMatchups_Names")
    Set matchupsValues = sheet.Range("BestMatchups_Values")
    Set expWinRates = sheet.Range("ExpRates")
    
    ' Clear current data
    matchupsNames.Value2 = ""
    matchupsValues.Value2 = ""
    
    ' Get number of decks that will be processed
    Dim numDecks As Integer
    numDecks = nMetaDecks - WorksheetFunction.CountIf(sheet.Range("Meta_Names"), "")
    
    ' Loop over decks in "Current Meta" table until numDecks have been copied. Each iteration add decks with max win rate.
    Dim i As Integer
    Dim curCell As Range
    Dim curStr As String
    Dim curVal As Double, curMax As Double, nextMax As Double
    Dim d As Deck
    
    i = 1
    curMax = 1#
    
    ' TODO: Speed this up by reading the whole range ahead of time, then writing it all at once
    While i <= numDecks
        nextMax = 0#
        For Each curCell In metaNames
            Set d = Factory.CreateDeck(curCell.Value2)
            ' Look up value in expected win rate table, convert to Double
            curStr = expWinRates.Cells(d.rowOffset, 2 * d.colOffset - 1).Value2
            If curStr = "" Then
                curVal = -1
            Else
                curVal = CDbl(curStr)
            End If
            
            If curVal = curMax Then
                matchupsNames.Cells(i).Value2 = curCell.Value2
                matchupsValues.Cells(i).Value2 = curVal
                i = i + 1
            ElseIf curVal > nextMax And curVal < curMax Then
                nextMax = curVal
            End If
        Next curCell
        
        curMax = nextMax
    Wend
End Function


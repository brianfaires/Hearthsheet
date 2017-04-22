Attribute VB_Name = "LogStuff"
Option Explicit

Dim minDate As Date
Dim myMinRank As Integer
Dim myMaxRank As Integer
Dim oppMinRank As Integer
Dim oppMaxRank As Integer

Public Function ReadSettings()
    Dim shtLog As Worksheet
    Set shtLog = Sheets("Log")
    
    minDate = CDate(shtLog.Cells(rLogCfg_Date, cLogCfg))
    myMinRank = CInt(shtLog.Cells(rLogCfg_MyMinRank, cLogCfg))
    myMaxRank = CInt(shtLog.Cells(rLogCfg_MyMaxRank, cLogCfg))
    oppMinRank = CInt(shtLog.Cells(rLogCfg_OppMinRank, cLogCfg))
    oppMaxRank = CInt(shtLog.Cells(rLogCfg_OppMaxRank, cLogCfg))
End Function

Public Function ClearDeckSheetGames()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsDeckSheet(ws.Name) Then
            ws.Range(ws.Cells(rDeckSheet_WLTable + 1, cDeckSheet_WLTable), _
                ws.Cells(rDeckSheet_WLTable + nDecksPerClass, cDeckSheet_WLTable + 2 * nClasses - 1)).Value2 = 0
            'Application.DisplayAlerts = False
            'ws.Delete
            'Application.DisplayAlerts = True
        End If
    Next
End Function

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

' This function is a little wasteful; it naively creates a sheet for each Deck, even if that sheet will be deleted in CleanUpDeckSheets()
Public Function CreateMissingDeckSheets()
    Dim d As Deck
    Dim shtLog As Worksheet
    Dim shtTemplate As Worksheet
    Set shtLog = Sheets("Log")
    Set shtTemplate = Sheets("Template")
    
    shtTemplate.Visible = True
    shtTemplate.Unprotect
    
    Dim Class As String
    Dim col As Integer
    Dim Row As Integer
    For col = cDeckGrid To cDeckGrid + nClasses - 1
        Class = shtLog.Cells(rDeckGrid, col)
        
        For Row = rDeckGrid + 1 To rDeckGrid + nDecksPerClass
            Set d = Factory.CreateDeck(shtLog.Cells(Row, col).Value2 & " " & Class)
        
            If d.IsValid And Not SheetExists(d.FullName) Then
                'Create a new Deck sheet, move it to the end
                shtTemplate.Copy After:=Sheets(Sheets.count)
                ActiveSheet.Name = d.FullName
            End If
        Next Row
    Next col
    
    shtTemplate.Protect
    shtTemplate.Visible = False
End Function

Public Function ProcessLogs()
    Dim g As Game
    Dim shtLog As Worksheet
    Set shtLog = Sheets("Log")
    
    Dim i As Long
    For i = rLogs To shtLog.Rows.count
            'Read until an empty timestamp
            If shtLog.Cells(i, cLogs_Date) = "" Then Exit For
            
            If Not PassesFilters(i) Then
                ' Strikethrough on invalid row
                shtLog.Rows(i).Font.StrikeThrough = True
            Else
                shtLog.Rows(i).Font.StrikeThrough = False
                
                Set g = Factory.CreateGame(shtLog.Rows(i))
                
                If g.myDeck.IsValid And g.oppDeck.IsValid Then
                    ' Write data to Deck sheets
                    g.WriteData
                    Range(shtLog.Cells(i, cLogs_Date), shtLog.Cells(i, cLogs_Notes)).Interior.Color = xlNone
                Else
                    ' Red text on non-existant Deck names
                    Range(shtLog.Cells(i, cLogs_Date), shtLog.Cells(i, cLogs_Notes)).Interior.Color = vbRed
                End If
            End If
        Next i
End Function

Private Function PassesFilters(Row As Long) As Boolean
    Dim shtLog As Worksheet
    Set shtLog = Sheets("Log")

    PassesFilters = True
    ' If ranks are omitted, don't fail the test
    If CDate(shtLog.Cells(Row, cLogs_Date)) < minDate Then PassesFilters = False
    If shtLog.Cells(Row, cLogs_MyRank) <> "" Then
        If CInt(shtLog.Cells(Row, cLogs_MyRank)) > myMinRank Then PassesFilters = False
        If CInt(shtLog.Cells(Row, cLogs_MyRank)) < myMaxRank Then PassesFilters = False
    End If
    If shtLog.Cells(Row, cLogs_OppRank) <> "" Then
        If CInt(shtLog.Cells(Row, cLogs_OppRank)) > oppMinRank Then PassesFilters = False
        If CInt(shtLog.Cells(Row, cLogs_OppRank)) < oppMaxRank Then PassesFilters = False
    End If
End Function

Public Function CleanUpDeckSheets()
    Dim sheet As Worksheet
    Dim shtLog As Worksheet
    Set shtLog = Worksheets("Log")
    
    ' Load settings from Log sheet
    Dim minGamesRed As Integer
    Dim minGamesYellow As Integer
    Dim minGamesBlack As Integer
    minGamesRed = CInt(Sheets("Log").Cells(rDeckSheetCfg_Red, cDeckSheetCfg).Value2)
    minGamesYellow = CInt(Sheets("Log").Cells(rDeckSheetCfg_Yellow, cDeckSheetCfg).Value2)
    minGamesBlack = CInt(Sheets("Log").Cells(rDeckSheetCfg_Black, cDeckSheetCfg).Value2)
    
    Dim gameCount As Integer
    Dim i As Integer
    For i = 1 To Sheets.count
        Set sheet = Sheets(i)
        If IsDeckSheet(sheet.Name) Then
            gameCount = CInt(sheet.Cells(rDeckSheet_GameCount, cDeckSheet_GameCount).Value2)
            If gameCount < minGamesRed Then
                Application.DisplayAlerts = False
                sheet.Delete
                Application.DisplayAlerts = True
                i = i - 1
            ElseIf gameCount < minGamesYellow Then
                sheet.Tab.Color = vbRed
            ElseIf gameCount < minGamesBlack Then
                sheet.Tab.Color = vbYellow
            Else
                sheet.Tab.Color = vbWhite
            End If
        End If
    Next i

    SortDeckSheets
    'UpdateDefaultRates
    UpdateAllBestMatchups
End Function

' A simple insertion sort by the number of games played
Public Function SortDeckSheets()
    Dim i, j As Integer
    Dim doSwap As Boolean
    For i = 1 To Sheets.count
        If IsDeckSheet(Sheets(i).Name) Then
            j = i - 1
            doSwap = False
            If IsDeckSheet(Sheets(j).Name) Then
                doSwap = CInt(Sheets(i).Cells(rDeckSheet_GameCount, cDeckSheet_GameCount).Value2) > _
                                CInt(Sheets(j).Cells(rDeckSheet_GameCount, cDeckSheet_GameCount).Value2)
            End If
            
            While doSwap
                j = j - 1
                doSwap = False
                If IsDeckSheet(Sheets(j).Name) Then
                    doSwap = CInt(Sheets(i).Cells(rDeckSheet_GameCount, cDeckSheet_GameCount).Value2) > _
                                    CInt(Sheets(j).Cells(rDeckSheet_GameCount, cDeckSheet_GameCount).Value2)
                End If
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
'                    If sheets("Priors").Cells(3, sourceCol) <> "" Then
'                        defaultRate = 0.5
'                        priorValue = sheets("Priors").Cells(sourceRow, sourceCol)
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
            ComputeBestMatchups (sht.Name)
        End If
    Next
End Function

Public Function ComputeBestMatchups(ByVal sheetName As String)
    Dim sheet As Worksheet
    Set sheet = Sheets(sheetName)
    sheet.Unprotect
    
    ' Clear current data
    Range(sheet.Cells(rDeckSheetBM, cDeckSheetBM_Name), _
                sheet.Cells(rDeckSheetBM + nMetaDecks - 1, cDeckSheetBM_Value)).Value2 = ""
    
    ' Get number of decks that will be processed
    Dim i, maxI, numDecks, currentMax, nextMax, metaOffset, rowOffset, colOffset As Integer
    Dim thisVal As Double
    
    i = rDeckSheetBM
    numDecks = Application.WorksheetFunction.CountIf( _
            Range(sheet.Cells(rDeckSheetMeta, cDeckSheetMeta_Value), _
                        sheet.Cells(rDeckSheetMeta + nMetaDecks - 1, cDeckSheetMeta_Value)), ">0")
    maxI = i + numDecks - 1
    currentMax = 1#
    
    ' Iterate on best decks in "Current Meta" table until maxI has been reached
    Dim d As Deck
    Dim strVal As String
    While i <= maxI
        nextMax = 0#
        For metaOffset = 0 To numDecks - 1
            Set d = Factory.CreateDeck(sheet.Cells(rDeckSheetMeta + metaOffset, 6).Value2)
            rowOffset = d.rowOffset
            colOffset = d.colOffset
            ' Pull win rate from expectedWinRate grid on this sheet, count as 0 if empty, so user notices data is missing
            ' EDIT: The expected win rates will now set themselves to .5 if no data exists on priors or in W/L table of this decksheet
            strVal = sheet.Cells(rDeckSheet_ExpWinRate + rowOffset, cDeckSheet_ExpWinRate + 2 * colOffset).Value2
            If strVal = "" Then
                thisVal = 0
            Else
                thisVal = CDbl(strVal)
            End If
            
            If thisVal = currentMax Then
                sheet.Cells(i, cDeckSheetBM_Name).Value2 = sheet.Cells(rDeckSheetMeta + metaOffset, cDeckSheetMeta_Name).Value2
                sheet.Cells(i, cDeckSheetBM_Value).Value2 = thisVal
                i = i + 1
            ElseIf thisVal > nextMax And thisVal < currentMax Then
                nextMax = thisVal
            End If
        Next
        currentMax = nextMax
    Wend
    
    sheet.Protect
End Function

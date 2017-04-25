Attribute VB_Name = "UserIO"
Option Explicit
' This module contains all the access points for the user. It should only contain subs and be the only module with Subs (to hide functionality on the Macros list)

' This is fired when the user presses the big button on the Log sheet. It will process all logs and update the deck sheets.
Sub Click_ProcessLogs()
    ' Disable visibility and calculation
    Application.ScreenUpdating = False
    Application.StatusBar = False
    Application.Calculation = xlCalculationManual
    
    ReadSettings
    SetProtectionAllDeckSheets False
    ClearDeckSheetGames
    CreateMissingDeckSheets
    ProcessLogs
    
    ' Re-enable early because Cleanup depends on it working
    Application.Calculation = xlCalculationAutomatic
    CleanUpDeckSheets
    SetProtectionAllDeckSheets True
    
    ' Restore visibility/calculation settings
    Sheets("Meta").Select
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

' This is fired when the user changes the number of games at the default win rate. This will change the win rates and best matchups.
' This fires when the user clicks the "Recompute" button on each deck sheet.
' TODO: Should this also fire when the user edits win/loss counts on deck sheets?
Sub RecalculateBestMatchupsForCurrentSheet()
    ' Disable visibility and calculation
    If IsDeckSheet(ActiveSheet.Name) Then
        Application.ScreenUpdating = False
        Application.StatusBar = False
        Application.Calculation = xlCalculationManual
        ActiveSheet.Unprotect
        
        ComputeBestMatchups ActiveSheet.Name
        
        ' Restore visibility/calculation settings
        ActiveSheet.Protect
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = True
        Application.ScreenUpdating = True
    End If
End Sub

' This is fired when the user clicks the "Cleanup" button on the Priors tab
Sub Click_CleanUpPriors()
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    
    ' Disable visibility and calculation
    Application.ScreenUpdating = False
    Application.StatusBar = False
    Application.Calculation = xlCalculationManual
    sht.Unprotect
    
    'Adjust hidden columns
    SetPriorVisibility
    
    ' Various checks, keep track of how many errors
    Dim errorsFound As Integer
    errorsFound = CheckPriorDiagonal
    errorsFound = errorsFound + CheckPriorHalfs
    errorsFound = errorsFound + CheckPriorValues

    DisplayPriorErrors errorsFound
    
    ' Restore visibility/calculation settings
    sht.Protect
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

' This fires when the user clicks the "Meta from Logs" button on the Meta sheet.
' TODO: Turn off calculations where allowable. This function is way slower than it should be.
Public Sub Click_UpdateMetaFromLogs()
    ' Disable visibility and calculation
    Application.ScreenUpdating = False
    Application.StatusBar = False
    Meta.Unprotect
    SetProtectionAllDeckSheets False
    
    ClearMetaData
    LoadMetaFromLogs
    
    UpdateMostPlayedClasses
    UpdateMostPlayedDecks
    UpdateAllBestMatchups
    UpdateBestMetaDecks
    
    ' Restore visibility/calculation settings
    SetProtectionAllDeckSheets True
    Meta.Protect
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

' This fires when the user clicks the "Manual Meta" button on the Meta sheet.
Public Sub Click_UpdateMetaFromMetaSheet()
    Dim shtMeta As Worksheet
    Set shtMeta = Sheets("Meta")
    
    'Disable visibility And Calculation
    Application.ScreenUpdating = False
    Application.StatusBar = False
    shtMeta.Unprotect
    SetProtectionAllDeckSheets False
    
    UpdateMostPlayedClasses
    UpdateMostPlayedDecks
    UpdateAllBestMatchups
    UpdateBestMetaDecks
    
    ' Restore visibility/calculation settings
    SetProtectionAllDeckSheets True
    shtMeta.Protect
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

' This fires when the user clicks the "Process all lineups" button on the Conquest sheet.
Public Sub Click_ProcessAllConquestLineups()
    Dim sht As Worksheet
    Dim match3 As Matchup3
    Dim match4 As Matchup4
    Dim myDecks As Lineup3, oppDecks As Lineup3
    Dim myLineup As Lineup4, oppLineup As Lineup4
    Dim i As Integer, j As Integer
    Dim rngOppLineups As Range, rngOppBan As Range, rngWinRate As Range, rngStrat00 As Range, rngStrat10a As Range
    Dim rngStrat10b As Range, rngStrat10c As Range, rngStrat01a As Range, rngStrat01b As Range, rngStrat01c As Range
     
    Set sht = Sheets("Conquest")
    Set myLineup = Factory.CreateLineup4(sht.Range("MyLineup"))
    Set rngOppLineups = sht.Range("OppLineups")
    Set rngOppBan = sht.Range("OppBans")
    Set rngWinRate = sht.Range("WinRates")
    Set rngStrat00 = sht.Range("Strat00")
    Set rngStrat10a = sht.Range("Strat10A")
    Set rngStrat10b = sht.Range("Strat10B")
    Set rngStrat10c = sht.Range("Strat10C")
    Set rngStrat01a = sht.Range("Strat01A")
    Set rngStrat01b = sht.Range("Strat01B")
    Set rngStrat01c = sht.Range("Strat01C")
   
    ' Disable screen update
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False
    sht.Unprotect
    
    ' Clear current results
    sht.Range("AllOppLineupData").Value2 = ""
    
    ' For each row of the lineups table
    Dim curRow As Range
    For i = 1 To rngOppLineups.Rows.count
        Set curRow = rngOppLineups.Rows(i)
        ' Clean up names, Strikethrough if cannot sanitize
        If curRow.Cells(1).Value2 <> "" Then
            If Not SanitizeDeckNames(curRow) Then
                curRow.Font.StrikeThrough = True
            Else
                ' Read lineups, determine best bans, then run matchup with the best lineups
                Set oppLineup = Factory.CreateLineup4(curRow)
                Set match4 = Factory.CreateMatchup4(myLineup, oppLineup)
                match4.ComputeBannedWinRates
                Set match3 = match4.TakeMaxMinBans
                match3.SetWinRatesFromDecks
                match3.RunThroughConquestSheet
                
                ' Determine which decks were banned
                Dim myBan As String, oppBan As String, myDeck As String, oppDeck As String
                For j = 1 To 4
                    myDeck = match4.lineupA.GetDeck(j).fullName
                    oppDeck = match4.lineupB.GetDeck(j).fullName
                    If match3.lineupA.GetDeck(1).fullName <> myDeck And _
                    match3.lineupA.GetDeck(2).fullName <> myDeck And match3.lineupA.GetDeck(3).fullName <> myDeck Then
                        oppBan = myDeck
                    End If
                    
                    If match3.lineupB.GetDeck(1).fullName <> oppDeck And _
                    match3.lineupB.GetDeck(2).fullName <> oppDeck And match3.lineupB.GetDeck(3).fullName <> oppDeck Then
                        myBan = oppDeck
                    End If
                Next j
                
                ' Show my ban via Strikethrough
                For j = 1 To 4
                    With curRow.Cells(j)
                        .Font.StrikeThrough = (.Value2 = myBan)
                    End With
                Next j
                
                ' Write opp ban and other results
                rngOppBan.Cells(i) = oppBan
                rngWinRate.Cells(i) = match3.winRate
                rngStrat00.Cells(i) = match3.strat00
                rngStrat10a.Cells(i) = match3.strat10a
                rngStrat10b.Cells(i) = match3.strat10b
                rngStrat10c.Cells(i) = match3.strat10c
                rngStrat01a.Cells(i) = match3.strat01a
                rngStrat01b.Cells(i) = match3.strat01b
                rngStrat01c.Cells(i) = match3.strat01c
            End If
        End If
    Next i
    
    ' Restore visibility/calculation settings
    sht.Protect
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

' This fires when user clicks "Pull Win Rates" on Conquest tab. Populates the matchup3 table.
Public Sub Click_PullWinRates()
    ' Disable visibility/calculation
    Application.ScreenUpdating = False
    Application.StatusBar = False
    Application.Calculation = xlCalculationManual
    
    Dim sht As Worksheet
    Set sht = Sheets("Conquest")
    
    Dim myDecks(3) As Deck, oppDecks(3) As Deck
    Dim i As Integer, j As Integer
    Dim rngMyDecks As Range, rngOppDecks As Range, rngWinRates As Range
    Set rngMyDecks = sht.Range("M3_MyDecks")
    Set rngOppDecks = sht.Range("M3_OppDecks")
    Set rngWinRates = sht.Range("M3_WinRates")
    
    For i = 1 To 3
        Set myDecks(i) = Factory.CreateDeck(rngMyDecks(i).Value2)
        Set oppDecks(i) = Factory.CreateDeck(rngOppDecks(i).Value2)
    Next i
    
    For i = 1 To 3
        For j = 1 To 3
            rngWinRates.Cells(i, j).Value2 = myDecks(i).GetWinRateVs(oppDecks(j))
        Next j
    Next i
    
    ' Restore visibility/calculation settings
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

' This fires when user clicks "Process Matchups" on Conquest tab. Uses the existing win rates to process a 3v3 matchup
Public Sub Click_ProcessMatchup()
    'Disable visibility / Calculation
    Application.ScreenUpdating = False
    Application.StatusBar = False
    Application.Calculation = xlCalculationManual
    Sheets("Conquest").Unprotect
    
    Dim match As Matchup3
    Set match = New Matchup3
    ' Can skip the other steps since these values are already in the Conquest/ConquestHidden sheets
    match.PullResults
    match.OutputConquestResults Conquest.Range("M3_WinRate")

    ' Restore visibility/calculation settings
    Sheets("Conquest").Protect
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

' This fires when user clicks "Compute Bans" on Conquest tab. It computes the outcomes of all bans for the given lineups.
Public Sub Click_ComputeBans()
    'Disable visibility / Calculation
    Application.ScreenUpdating = False
    Application.StatusBar = False
    'Application.Calculation = xlCalculationManual
    Sheets("Conquest").Unprotect
    
    Dim sht As Worksheet
    Set sht = Sheets("Conquest")
    Dim match As Matchup4
    Dim myLineup As Lineup4, oppLineup As Lineup4
    Dim outputA_BansA As Range, outputA_BansB As Range, outputB_BansA As Range, outputB_BansB As Range
    Dim outputA_WinRate As Range, outputB_WinRate As Range
    Set outputA_BansA = sht.Range("BansA_A")
    Set outputA_BansB = sht.Range("BansA_B")
    Set outputB_BansA = sht.Range("BansB_A")
    Set outputB_BansB = sht.Range("BansB_B")
    Set outputA_WinRate = sht.Range("BansA_WinRates")
    Set outputB_WinRate = sht.Range("BansB_WinRates")
    
    Set myLineup = Factory.CreateLineup4(sht.Range("Bans_LineupA"))
    Set oppLineup = Factory.CreateLineup4(sht.Range("Bans_LineupB"))
    Set match = Factory.CreateMatchup4(myLineup, oppLineup)
    match.ComputeBannedWinRates
    
    Dim i As Integer, j As Integer, rowNum As Integer
    For i = 1 To 4
        For j = 1 To 4
            rowNum = 4 * (i - 1) + j
            ' First table (A's perspective)
            outputA_BansA(rowNum).Value2 = match.lineupB.GetDeck(i).fullName
            outputA_BansB(rowNum).Value2 = match.lineupA.GetDeck(j).fullName
            outputA_WinRate(rowNum).Value2 = match.GetBanWinRate(j, i)
            
            ' Second table (B's perspective)
            outputB_BansA(rowNum).Value2 = match.lineupB.GetDeck(j).fullName
            outputB_BansB(rowNum).Value2 = match.lineupA.GetDeck(i).fullName
            outputB_WinRate(rowNum).Value2 = 1 - match.GetBanWinRate(i, j)
        Next j
    Next i
    
    ' Restore visibility/calculation settings
    Sheets("Conquest").Protect
    'Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = True
    Application.ScreenUpdating = True
End Sub

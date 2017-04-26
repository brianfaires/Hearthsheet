VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matchup3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public lineupA As Lineup3
Public lineupB As Lineup3

' Match win rate
Public winRate As Double

' Text describing the optimal mixed strategies for deck queue order
Public strat00, strat10a, strat10b, strat10c, strat01a, strat01b, strat01c As String

' Deck vs deck win rates
Private winRates(3, 3) As Double
Public Function GetWinRate(a As Integer, b As Integer) As Double
    GetWinRate = winRates(a, b)
End Function

Public Function PopulateValues(a As Lineup3, b As Lineup3)
    Set lineupA = a
    Set lineupB = b
End Function

' Set the deck vs deck win rates
Public Function SetWinRatesFromDecks()
    Dim i As Integer, j As Integer
    With Conquest.Range("M3_WinRates")
        For i = 1 To 3
            For j = 1 To 3
                winRates(i, j) = lineupA.GetDeck(i).GetWinRateVs(lineupB.GetDeck(j))
                .Cells(i, j).Value2 = winRates(i, j)
            Next j
        Next i
    End With
End Function

' Pull win rates from the matchup table of the Conquest sheet
Public Function SetWinRatesFromConquestSheet()
    Dim i As Integer, j As Integer
    
    With Conquest.Range("M3_WinRates")
        For i = 1 To 3
            For j = 1 To 3
                winRates(i, j) = .Cells(rConqMatchup, cConqMatchup)
            Next j
        Next i
    End With
End Function

' Use ConquestHidden sheet to do the heavy lifting
Public Function RunThroughConquestSheet()
    With Conquest.Range("M3_WinRates")
        Dim i As Integer, j As Integer
        
        ' Populate ConquestHidden via the plugin on Conquest sheet
        For i = 1 To 3
            For j = 1 To 3
                .Cells(i, j).Value2 = winRates(i, j)
            Next
        Next
    End With
    
    ConquestHidden.Calculate
    PullResults
End Function

' Pull results from ConquestHidden
' TODO: Clean up magic numbers by naming cells in ConquestHidden
Public Function PullResults()
    With ConquestHidden
        winRate = Round(.Cells(1, 8), 3)
        strat00 = Round(.Cells(3, 12), 2) & ", " & Round(.Cells(4, 12), 2) & ", " & Round(.Cells(5, 12), 2)
        strat10a = Round(.Cells(37, 5), 2) & ", " & Round(.Cells(38, 5), 2)
        strat10b = Round(.Cells(37, 11), 2) & ", " & Round(.Cells(38, 11), 2)
        strat10c = Round(.Cells(37, 17), 2) & ", " & Round(.Cells(38, 17), 2)
        strat01a = Round(.Cells(151, 4), 2) & ", " & Round(.Cells(152, 4), 2) & ", " & Round(.Cells(153, 4), 2)
        strat01b = Round(.Cells(151, 10), 2) & ", " & Round(.Cells(152, 10), 2) & ", " & Round(.Cells(153, 10), 2)
        strat01c = Round(.Cells(151, 16), 2) & ", " & Round(.Cells(152, 16), 2) & ", " & Round(.Cells(153, 16), 2)
    End With
End Function

' Writes results to a specified row of the lineups table
Public Function OutputConquestResults(outRange As Range)
    With outRange
        .Cells(1, 1).Value2 = winRate
        .Cells(1, 2).Value2 = strat00
        .Cells(1, 3).Value2 = strat10a
        .Cells(1, 4).Value2 = strat10b
        .Cells(1, 5).Value2 = strat10c
        .Cells(1, 6).Value2 = strat01a
        .Cells(1, 7).Value2 = strat01b
        .Cells(1, 8).Value2 = strat01c
    End With
End Function

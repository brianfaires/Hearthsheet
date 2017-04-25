VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matchup4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public lineupA As Lineup4
Public lineupB As Lineup4

' The deck vs deck win rates
Private winRates(4, 4) As Double
'Matchup win rates, when 1st index is deck banned BY A, away from B, and 2nd index banned by B away from A
Private banWinRates(4, 4) As Double

Public Function GetWinRate(a As Integer, b As Integer) As Double
    GetWinRate = winRates(a, b)
End Function

Public Function GetBanWinRate(a As Integer, b As Integer) As Double
    GetBanWinRate = banWinRates(a, b)
End Function

Public Function PopulateValues(a As Lineup4, b As Lineup4)
    Set lineupA = a
    Set lineupB = b
End Function

' Set the deck vs deck win rates
Public Function SetWinRates()
    Dim i As Integer, j As Integer
    For i = 1 To 4
        For j = 1 To 4
            winRates(i, j) = lineupA.GetDeck(i).GetWinRateVs(lineupB.GetDeck(j))
        Next j
    Next i
End Function

' Counts all 16 possible arrangements and records the resulting match win rates
Public Function ComputeBannedWinRates()
    SetWinRates
    Dim match As Matchup3
    
    Dim i As Integer, j As Integer
    For i = 1 To 4
        For j = 1 To 4
            Set match = Factory.CreateMatchup3(lineupA.Ban(lineupA.GetDeck(i).fullName), lineupB.Ban(lineupB.GetDeck(j).fullName))
            match.SetWinRatesFromDecks
            match.RunThroughConquestSheet
            banWinRates(i, j) = match.winRate
        Next j
    Next i
End Function

' Bans with maximin choices (maximize the minimum win rate)
' TODO: This could copy the win rates into the Matchup3 object, to avoid a call to SetWinRates... afterward
Public Function TakeMaxMinBans() As Matchup3
    Dim banA As String, banB As String
    Dim minWinRate(4) As Double
    Dim curMin As Double
    Dim i As Integer, j As Integer
    
    ' Determine A's ban
    curMin = 1
    For i = 1 To 4
        ' Assuming A bans deck(i), find the minimum match win rate
        For j = 1 To 4
            If winRates(i, j) < curMin Then curMin = winRates(i, j)
        Next j
        minWinRate(i) = curMin
    Next i
    
    curMin = 0
    For i = 1 To 4
        If minWinRate(i) > curMin Then
            curMin = minWinRate(i)
            banB = lineupB.GetDeck(i).fullName
        End If
    Next i
    
    
    ' Repeat for B's ban; remember that B's win rate = 1-winRates()
    curMin = 0
    For i = 1 To 4
        'Assuming B bans deck(i), find the minimum match win rate
        For j = 1 To 4
            If winRates(j, i) > curMin Then curMin = winRates(i, j)
        Next j
        minWinRate(i) = curMin
    Next i
    
    curMin = 1
    For i = 1 To 4
        If minWinRate(i) < curMin Then
            curMin = minWinRate(i)
            banA = lineupA.GetDeck(i).fullName
        End If
    Next i
    
    Set TakeMaxMinBans = Factory.CreateMatchup3(lineupA.Ban(banA), lineupB.Ban(banB))
End Function


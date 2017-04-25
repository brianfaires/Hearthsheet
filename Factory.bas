Attribute VB_Name = "Factory"
' Factory Module: Provides convenient "constructors"
Option Explicit

Public Function CreateDeck(rawName As String) As Deck
    Dim retVal As Deck
    Set retVal = New Deck
    retVal.PopulateValues rawName
    Set CreateDeck = retVal
End Function

Public Function CreateGame(inputRow As Range) As Game
    Dim retVal As Game
    Set retVal = New Game
    retVal.PopulateValues inputRow
    Set CreateGame = retVal
End Function

Public Function CreateLineup4(inputRange As Range) As Lineup4
    Dim retVal As Lineup4
    Set retVal = New Lineup4
    retVal.PopulateValues inputRange
    Set CreateLineup4 = retVal
End Function

Public Function CreateMatchup3(lineupA As Lineup3, lineupB As Lineup3)
    Dim retVal As Matchup3
    Set retVal = New Matchup3
    retVal.PopulateValues lineupA, lineupB
    Set CreateMatchup3 = retVal
End Function

Public Function CreateMatchup4(lineupA As Lineup4, lineupB As Lineup4) As Matchup4
    Dim retVal As Matchup4
    Set retVal = New Matchup4
    retVal.PopulateValues lineupA, lineupB
    Set CreateMatchup4 = retVal
End Function

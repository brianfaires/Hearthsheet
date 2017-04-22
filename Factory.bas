Attribute VB_Name = "Factory"
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


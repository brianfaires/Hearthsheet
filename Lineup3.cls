VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lineup3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private decks(3) As Deck

Public Function GetDeck(i As Integer) As Deck
    Set GetDeck = decks(i)
End Function

Public Function SetDeck(i As Integer, d As Deck)
    Set decks(i) = d
End Function


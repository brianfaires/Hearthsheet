VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lineup4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private decks(4) As Deck
Public Function GetDeck(i As Integer) As Deck
    Set GetDeck = decks(i)
End Function

Public Function PopulateValues(inputRange As Range)
    Dim i As Integer
    For i = 1 To 4
        Set decks(i) = Factory.CreateDeck(inputRange.Cells(i).Value2)
    Next
End Function

Public Function Ban(banned As String) As Lineup3
    Dim i As Integer
    Dim cur As Integer
    Dim retVal As Lineup3
    Set retVal = New Lineup3
    
    cur = 1
    For i = 1 To 4
        If decks(i).fullName <> banned Then
            retVal.SetDeck cur, decks(i)
            cur = cur + 1
        End If
    Next i
    
    Set Ban = retVal
End Function

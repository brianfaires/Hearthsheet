VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public timestamp As Date
Public myDeck As Deck
Public myRank As Integer
Public oppDeck As Deck
Public oppRank As Integer
Public won As Boolean
Public notes As String

Public Function PopulateValues(logRow As Range)
    timestamp = logRow.Cells(cLogs_Date).Value2
    Set myDeck = Factory.CreateDeck(logRow.Cells(cLogs_MyDeck).Value2)
    myRank = logRow.Cells(cLogs_MyRank).Value2
    Set oppDeck = Factory.CreateDeck(logRow.Cells(cLogs_OppDeck).Value2)
    oppRank = logRow.Cells(cLogs_OppRank).Value2
    won = logRow.Cells(cLogs_Won).Value2 = 1 Or _
               StartsWith(LCase(logRow.Cells(cLogs_Won)), "t") Or _
               StartsWith(logRow.Cells(cLogs_Won), "y")
    notes = logRow.Cells(cLogs_Notes).Value2
End Function
 
 
' Increments the win/loss count in the Deck sheets
Public Function WriteData()
    Dim curSheet As Worksheet
    Dim colOffset As Integer
    Dim initValue As Integer
    
    If SheetExists(myDeck.FullName) Then
        Set curSheet = Sheets(myDeck.FullName)
        colOffset = 2 * oppDeck.colOffset
        If Not won Then colOffset = colOffset + 1 ' Move to LOSS column
        initValue = CInt(curSheet.Cells(rDeckSheet_WLTable + oppDeck.rowOffset, cDeckSheet_WLTable + colOffset).Value2)
        curSheet.Cells(rDeckSheet_WLTable + oppDeck.rowOffset, cDeckSheet_WLTable + colOffset).Value2 = initValue + 1
    End If
        
    If SheetExists(oppDeck.FullName) Then
        Set curSheet = Sheets(oppDeck.FullName)
        colOffset = 2 * myDeck.colOffset
        If won Then colOffset = colOffset + 1 ' Move to LOSS column
        initValue = CInt(curSheet.Cells(rDeckSheet_WLTable + myDeck.rowOffset, cDeckSheet_WLTable + colOffset).Value2)
        curSheet.Cells(rDeckSheet_WLTable + myDeck.rowOffset, cDeckSheet_WLTable + colOffset) = initValue + 1
    End If
End Function

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

' Accepts an entire row as input
Public Function PopulateValues(logRow As Range)
    With logRow
        timestamp = .Cells(LogCfgCol_Date).Value2
        Set myDeck = Factory.CreateDeck(.Cells(LogCfgCol_MyDeck).Value2)
        myRank = .Cells(LogCfgCol_MyRank).Value2
        Set oppDeck = Factory.CreateDeck(.Cells(LogCfgCol_OppDeck).Value2)
        oppRank = .Cells(LogCfgCol_OppRank).Value2
        won = .Cells(LogCfgCol_Win).Value2 = 1 Or StartsWith(LCase(.Cells(LogCfgCol_Win)), "t") Or StartsWith(.Cells(LogCfgCol_Win), "y")
        notes = .Cells(LogCfgCol_Notes).Value2
    End With
End Function
 
 
' Increments the win/loss count in the Deck sheets
Public Function WriteData()
    Dim curSheet As Worksheet
    Dim wlTable As Range
    Dim colOffset As Integer
    Dim initValue As Integer
    
    If SheetExists(myDeck.fullName) Then
        Set curSheet = Sheets(myDeck.fullName)
        colOffset = 2 * oppDeck.colOffset
        If won Then colOffset = colOffset - 1 ' Move to WIN column
        Set wlTable = curSheet.Range("WL_Table")
        initValue = CInt(wlTable.Cells(oppDeck.rowOffset, colOffset).Value2)
        wlTable.Cells(oppDeck.rowOffset, colOffset).Value2 = initValue + 1
    End If
        
    If SheetExists(oppDeck.fullName) Then
        Set curSheet = Sheets(oppDeck.fullName)
        colOffset = 2 * myDeck.colOffset
        If Not won Then colOffset = colOffset - 1 ' Move to WIN column
        Set wlTable = curSheet.Range("WL_Table")
        initValue = CInt(wlTable.Cells(myDeck.rowOffset, colOffset).Value2)
        wlTable.Cells(myDeck.rowOffset, colOffset) = initValue + 1
    End If
End Function

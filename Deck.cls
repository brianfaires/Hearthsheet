VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Deck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ClassName As String
Public DeckName As String

' This is used primarily for parsing user input to a proper Deck name
' TODO: Could change this or something upstream to store already-made decks and reference them instead of creating from scratch
Public Function PopulateValues(rawName As String)
    Dim i As Integer, j As Integer
    Dim foundClass As Boolean
    Dim classHeader As String, tempName As String, shortName As String
    
    Dim classNames As Range, deckNames As Range
    Set classNames = Log.Range("Classes")
    Set deckNames = Log.Range("Decks")
    
    tempName = LCase(Trim(rawName))
    For i = 1 To nClasses
        classHeader = LCase(classNames.Cells(i).Value2)
        foundClass = EndsWith(tempName, classHeader) Or (classHeader = "warlock" And EndsWith(tempName, "lock"))
        If Not foundClass Then
            ' Also check each Deck name for exact match, e.g. user entered "Zoo" instead of "Zoolock"
            For j = 1 To nDecksPerClass
                If LCase(deckNames.Cells(j, i)) = tempName Then
                    foundClass = True
                End If
            Next j
        End If

        If foundClass Then
            ' shortName as the lowercased Deck name without class suffix
            If EndsWith(tempName, classHeader) Then
                shortName = Trim(Left(tempName, Len(tempName) - Len(classHeader)))
            ElseIf classHeader = "warlock" And EndsWith(tempName, "lock") Then
                shortName = Trim(Left(tempName, Len(tempName) - Len("lock")))
            Else
                shortName = Trim(tempName)
            End If
            
            For j = 1 To nDecksPerClass
                If LCase(deckNames.Cells(j, i)) = shortName Then
                    ' Capitalize first letter
                    DeckName = UCase(Left(shortName, 1)) & LCase(Mid(shortName, 2))
                    ClassName = UCase(Left(classHeader, 1)) & LCase(Mid(classHeader, 2))
                    Exit For
                End If
            Next j
        End If
    Next i
End Function

Public Property Get fullName() As String
    fullName = DeckName + " " + ClassName
End Function

' From 0 to nDecks-1
Public Property Get AbsoluteOffset() As Integer
    Dim classOffset As Integer
    Dim deckOffset As Integer
    
    ' Return -1 if invalid
    AbsoluteOffset = -1
        
    classOffset = colOffset()
    If classOffset <> -1 Then
        deckOffset = rowOffset
        If deckOffset <> -1 Then
            AbsoluteOffset = nDecksPerClass * (classOffset - 1) + deckOffset
        End If
    End If
End Property

Public Property Get IsValid() As Boolean
    IsValid = DeckName <> "" And AbsoluteOffset <> -1
End Property

' TODO: These are pretty wasteful, should store value after doing computation, maybe even computing this for all somewhere else
' and accessing the results here.
Public Property Get colOffset() As Integer
    Dim deckGrid As Range
    Set deckGrid = Log.Range("Classes")
    
    ' Return -1 if not found
    colOffset = -1
    
    Dim i As Integer
    For i = 1 To nClasses
        If ClassName = deckGrid.Cells(i).Value2 Then
            colOffset = i
            Exit For
        End If
    Next i
End Property

Public Property Get rowOffset() As Integer
    ' Only look this up once
    Dim fastColOffset As Integer
    fastColOffset = colOffset
    
    ' Return -1 if not found
    rowOffset = -1
    
    With Log.Range("Decks")
        Dim i As Integer
        For i = 1 To nDecksPerClass
            If StrComp(DeckName, .Cells(i, fastColOffset).Value2, vbTextCompare) = 0 Then
                rowOffset = i
                Exit For
            End If
        Next i
    End With
End Property

'Gets the best known win rate for this matchup. First checks deck sheets, then goes to priors tab
Public Function GetWinRateVs(opp As Deck) As Double
    If SheetExists(fullName) Then
        ' Look up in deck sheet
        GetWinRateVs = Round(Sheets(fullName).Range("ExpRates").Cells(opp.rowOffset, 2 * opp.colOffset - 1).Value2, 3)
    ElseIf SheetExists(opp.fullName) Then
        ' Look up in opponent's deck sheet
        GetWinRateVs = Round(1 - Sheets(opp.fullName).Range("ExpRates").Cells(rowOffset, 2 * colOffset - 1).Value2, 3)
    Else
        ' Pull win rate from Priors sheet
        Dim cellVal As String
        cellVal = Round(Priors.Range("AllPriors").Cells(AbsoluteOffset, opp.AbsoluteOffset).Value2, 3)
        If cellVal = "" Then
            ' Check other half of Priors table
            cellVal = Round(1 - Priors.Range("AllPriors").Cells(opp.AbsoluteOffset, AbsoluteOffset).Value2, 3)
            If cellVal = "" Then
                GetWinRateVs = 0.5
            Else
                GetWinRateVs = CDbl(cellVal)
            End If
        Else
            GetWinRateVs = 0.5
        End If
    End If
End Function


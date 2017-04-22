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
' Could change this or something upstream to store already-made decks and reference them instead of creating from scratch
Public Function PopulateValues(rawName As String)
    Dim shtLog As Worksheet
    Set shtLog = Sheets("Log")
    
    Dim col, i As Integer
    Dim foundClass As Boolean
    Dim classHeader As String
    Dim tempName As String
    
    tempName = LCase(Trim(rawName))
    For col = cDeckGrid To cDeckGrid + nClasses - 1
        classHeader = LCase(shtLog.Cells(rDeckGrid, col).Text)
        foundClass = EndsWith(tempName, classHeader) Or (classHeader = "warlock" And EndsWith(tempName, "lock"))
        If Not foundClass Then
            ' Also check each Deck name for exact match, e.g. user entered "Zoo" instead of "Zoolock"
            For i = 1 To nDecksPerClass
                If LCase(shtLog.Cells(rDeckGrid + i, col)) = tempName Then
                    foundClass = True
                End If
            Next i
        End If

        If foundClass Then
            Dim shortName As String
            ' shortName as the lowercased Deck name without class suffix
            If EndsWith(tempName, classHeader) Then
                shortName = Trim(Left(tempName, Len(tempName) - Len(classHeader)))
            ElseIf classHeader = "warlock" And EndsWith(tempName, "lock") Then
                shortName = Trim(Left(tempName, Len(tempName) - Len("lock")))
            Else
                shortName = Trim(tempName)
            End If
            
            Dim Row As Integer
            For Row = rDeckGrid To rDeckGrid + nDecksPerClass - 1
                If LCase(shtLog.Cells(Row, col)) = shortName Then
                    ' Capitalize first letter
                    DeckName = UCase(Left(shortName, 1)) & LCase(Mid(shortName, 2))
                    ClassName = UCase(Left(classHeader, 1)) & LCase(Mid(classHeader, 2))
                    Exit For
                End If
            Next Row
        End If
    Next col
End Function

Public Property Get FullName() As String
    FullName = DeckName + " " + ClassName
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
            AbsoluteOffset = nDecksPerClass * classOffset + deckOffset
        End If
    End If
End Property

Public Property Get IsValid() As Boolean
    IsValid = DeckName <> "" And AbsoluteOffset <> -1
End Property

' TODO: These are pretty wasteful, should store value after doing computation, maybe even computing this for all somewhere else
' and accessing the results here.
Public Property Get colOffset() As Integer
    ' Return -1 if not found
    colOffset = -1
    
    Dim col As Integer
    For col = cDeckGrid To cDeckGrid + nClasses - 1
        If ClassName = Log.Cells(rDeckGrid, col) Then
            colOffset = col - cDeckGrid
            Exit For
        End If
    Next col
End Property

Public Property Get rowOffset() As Integer
    ' Return -1 if not found
    rowOffset = -1
    
    Dim Row As Integer
    For Row = rDeckGrid + 1 To rDeckGrid + nDecksPerClass
        If StrComp(DeckName, Log.Cells(Row, cDeckGrid + colOffset).Value2, vbTextCompare) = 0 Then
            rowOffset = Row - rDeckGrid
            Exit For
        End If
    Next Row
End Property

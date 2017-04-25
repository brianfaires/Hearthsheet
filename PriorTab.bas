Attribute VB_Name = "PriorTab"
Option Explicit

' TODO: This should perhaps be in a Class module
Private allWinRates(nDecks, nDecks) As Double

' Reads from the allWinRates() array, avoiding repeated reads to the Priors spreadsheet
Public Function GetWinRate(myDeck As Deck, oppDeck As Deck) As Double
    GetWinRate = allWinRates(myDeck.AbsoluteOffset, oppDeck.AbsoluteOffset)
End Function

' Populates the allWinRates() array, to avoid repeated reads to the Priors spreadsheet
Public Function LoadWinRatesFromPriorsSheet()
    Dim i As Integer, j As Integer
    
    With Priors.Range("AllPriors")
        ' We could only read half of the grid and infer the rest, but doing it this way allows for asymmetrical grids if the user thinks that's appropriate
        For i = 1 To nDecks
            For j = 1 To nDecks
                allWinRates(i, j) = CDbl(.Cells(i, j).Value2)
            Next j
        Next i
    End With
End Function


'Hide rows/cols for missing decks
Public Function SetPriorVisibility()
    Dim i, j, rPriors, cPriors As Integer
    Dim wasHidden, isHidden As Boolean
    
    With Priors
        rPriors = .Range("AllPriors").Row
        cPriors = .Range("AllPriors").Column
        
        For i = 0 To nDecks - 1
            wasHidden = .Columns(rPriors + i).Hidden
            isHidden = Trim(.Cells(rPriors + i, cPriors - 1).Value2) = ""
            .Columns(cPriors + i).Hidden = isHidden
            .Rows(rPriors + i).Hidden = isHidden
            
            'If unhiding a row/col, initialize the values
            If wasHidden And Not isHidden Then
                For j = 1 To nDecks
                    If i = j Then
                        .Cells(rPriors + i, cPriors + j).Value2 = 0.5
                    Else
                        .Cells(rPriors + i, cPriors + j).Value2 = ""
                        .Cells(rPriors + j, cPriors + i).Value2 = ""
                    End If
                Next
            End If
        Next
    End With
End Function

Public Function CheckPriorDiagonal() As Integer
    Dim errorsFound, i As Integer
    Dim deckNames As Range
    Set deckNames = Priors.Range("DecksA")
    'Returns number of errors marked
    errorsFound = 0
    
    With Priors.Range("AllPriors")
        'Ensure all .50's along the diagonal, strikethrough errors
        For i = 1 To nDecks
            If Trim(deckNames(i).Value2) <> "" Then
                If .Cells(i, i).Value2 <> 0.5 Then
                    .Cells(i, i).Font.StrikeThrough = True
                    errorsFound = errorsFound + 1
                Else
                    .Cells(i, i).Font.StrikeThrough = False
                End If
            End If
        Next
    End With
    
    CheckPriorDiagonal = errorsFound
End Function

Public Function CheckPriorHalfs() As Integer
    Dim errorsFound, i, j As Integer
    
    'Returns number of errors marked
    errorsFound = 0
    
    With Priors.Range("AllPriors")
        'Ensure both halves of matrix agree, strikethrough errors
        For i = 1 To nDecks
            For j = i + 1 To nDecks
                If .Cells(i, j).Value2 = "" And .Cells(j, i).Value2 <> "" Then
                    ' Fill in other half of matrix to match
                    .Cells(i, j).Value2 = 1 - CDbl(.Cells(j, i).Value2)
                ElseIf .Cells(j, i).Value2 = "" And .Cells(i, j).Value2 <> "" Then
                    ' Fill in other half of matrix to match
                    .Cells(j, i).Value2 = 1 - CDbl(.Cells(i, j).Value2)
                    .Cells(i, j).Font.StrikeThrough = False
                    .Cells(j, i).Font.StrikeThrough = False
                ElseIf .Cells(i, j).Value2 <> "" And CDbl(.Cells(i, j).Value2) + CDbl(.Cells(j, i).Value2) <> 1 Then
                    ' Mismatch; error
                    .Cells(i, j).Font.StrikeThrough = True
                    .Cells(j, i).Font.StrikeThrough = True
                    errorsFound = errorsFound + 1
                Else
                    ' Match; no error
                    .Cells(i, j).Font.StrikeThrough = False
                    .Cells(j, i).Font.StrikeThrough = False
                End If
            Next
        Next
    End With
    
    CheckPriorHalfs = errorsFound
End Function

Public Function CheckPriorValues() As Integer
    Dim errorsFound, i, j As Integer
    
    'Returns number of errors marked
    errorsFound = 0
    
    With Priors.Range("AllPriors")
        'Check for <0 or >1, strikethrough errors
        For i = 1 To nDecks
            For j = i + 1 To nDecks
                If CDbl(.Cells(i, j).Value2) < 0 Or _
                CDbl(.Cells(i, j).Value2) > 1 Then
                    ' Since the matrix is diagonally symmetric at this point, mark both cells as errors
                    .Cells(i, j).Font.StrikeThrough = True
                    .Cells(j, i).Font.StrikeThrough = True
                    errorsFound = errorsFound + 1
                End If
            Next
        Next
    End With
    
    CheckPriorValues = errorsFound
End Function

Public Function DisplayPriorErrors(errorsFound As Integer)
    With Priors.Range("Output")
        If errorsFound > 0 Then
            .Value2 = errorsFound & " error(s) were found"
            .Font.Color = vbRed
        Else
            .Value2 = "Enter estimated win rates below"
            .Font.Color = vbBlack
        End If
    End With
End Function

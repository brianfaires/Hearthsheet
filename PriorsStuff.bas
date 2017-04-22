Attribute VB_Name = "PriorsStuff"
Option Explicit

Private allWinRates(nDecks, nDecks) As Double

Public Function GetWinRate(myDeck As Deck, oppDeck As Deck) As Double
    GetWinRate = allWinRates(myDeck.AbsoluteOffset, oppDeck.AbsoluteOffset)
End Function

Public Function LoadWinRatesFromPriorsSheet()
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    
    ' We could only read half of the grid and infer the rest, but doing it this way allows for asymmetrical grids if the user thinks that's appropriate
    For i = 1 To nDecks
        For j = 1 To nDecks
            allWinRates(i, j) = CDbl(sht.Cells(rPriors + i, cPriors + j).Value2)
        Next j
    Next i
    
    sht = Nothing
End Function

Public Function SetPriorVisibility()
    'Hide rows/cols for missing decks
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    
    Dim i, j As Integer
    Dim wasHidden, isHidden As Boolean
    For i = 1 To nDecks
        wasHidden = sht.Columns(rPriors + i).Hidden
        isHidden = Trim(sht.Cells(rPriors + i, cPriors).Value2) = ""
        sht.Columns(cPriors + i).Hidden = isHidden
        sht.Rows(rPriors + i).Hidden = Trim(sht.Cells(rPriors + i, cPriors).Value2) = ""
        
        'If unhiding a row/col, initialize the values
        If wasHidden And Not isHidden Then
            For j = 1 To nDecks
                If i = j Then
                    sht.Cells(rPriors + i, cPriors + j).Value2 = 0.5
                Else
                    sht.Cells(rPriors + i, cPriors + j).Value2 = ""
                    sht.Cells(rPriors + j, cPriors + i).Value2 = ""
                End If
            Next
        End If
    Next
End Function

Public Function CheckPriorDiagonal() As Integer
    'Returns number of errors marked
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    Dim errorsFound, i As Integer
    errorsFound = 0
    
    'Ensure all .50's along the diagonal, strikethrough errors
    For i = 1 To nDecks
        If Trim(sht.Cells(rPriors + i, cPriors).Value2) <> "" Then
            If sht.Cells(rPriors + i, cPriors + i).Value2 <> 0.5 Then
                sht.Cells(rPriors + i, cPriors + i).Font.StrikeThrough = True
                errorsFound = errorsFound + 1
            Else
                sht.Cells(rPriors + i, cPriors + i).Font.StrikeThrough = False
            End If
        End If
    Next
    
    CheckPriorDiagonal = errorsFound
End Function

Public Function CheckPriorHalfs() As Integer
    'Returns number of errors marked
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    Dim errorsFound, i, j As Integer
    errorsFound = 0
    
    'Ensure both halves of matrix agree, strikethrough errors
    For i = 1 To nDecks
        For j = i + 1 To nDecks
            If sht.Cells(rPriors + i, cPriors + j).Value2 = "" And sht.Cells(rPriors + j, cPriors + i).Value2 <> "" Then
                ' Fill in other half of matrix to match
                sht.Cells(rPriors + i, cPriors + j).Value2 = 1 - CDbl(sht.Cells(rPriors + j, cPriors + i).Value2)
            ElseIf sht.Cells(rPriors + j, cPriors + i).Value2 = "" And sht.Cells(rPriors + i, cPriors + j).Value2 <> "" Then
                ' Fill in other half of matrix to match
                sht.Cells(rPriors + j, cPriors + i).Value2 = 1 - CDbl(sht.Cells(rPriors + i, cPriors + j).Value2)
                sht.Cells(rPriors + i, cPriors + j).Font.StrikeThrough = False
                sht.Cells(cPriors + j, rPriors + i).Font.StrikeThrough = False
            ElseIf sht.Cells(rPriors + i, cPriors + j).Value2 <> "" And _
            sht.Cells(rPriors + j, cPriors + i).Value2 <> "" And _
            CDbl(sht.Cells(rPriors + i, cPriors + j).Value2) + CDbl(sht.Cells(rPriors + j, cPriors + i).Value2) <> 1 Then
                ' Mismatch; error
                sht.Cells(rPriors + i, cPriors + j).Font.StrikeThrough = True
                sht.Cells(cPriors + j, rPriors + i).Font.StrikeThrough = True
                errorsFound = errorsFound + 1
            Else
                ' Match; no error
                sht.Cells(rPriors + i, cPriors + j).Font.StrikeThrough = False
                sht.Cells(cPriors + j, rPriors + i).Font.StrikeThrough = False
            End If
        Next
    Next
    
    CheckPriorHalfs = errorsFound
End Function

Public Function CheckPriorValues() As Integer
    'Returns number of errors marked
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    Dim errorsFound, i, j As Integer
    errorsFound = 0
    
    'Check for <0 or >1, strikethrough errors
    For i = 1 To nDecks
        For j = i + 1 To nDecks
            If CDbl(sht.Cells(rPriors + i, cPriors + j).Value2) < 0 Or _
            CDbl(sht.Cells(rPriors + i, cPriors + j).Value2) > 1 Then
                ' Since the matrix is symmetric at this point, mark both cells as error'd
                sht.Cells(rPriors + i, cPriors + j).Font.StrikeThrough = True
                sht.Cells(rPriors + j, cPriors + i).Font.StrikeThrough = True
                errorsFound = errorsFound + 1
            End If
        Next
    Next
    
    CheckPriorValues = errorsFound
End Function

Public Function DisplayPriorErrors(errorsFound As Integer)
    'Display error count
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    Dim errorMsg As String
    errorMsg = ""
    If errorsFound > 0 Then
        errorMsg = errorsFound & " error(s) were found"
        sht.Cells(rPriors_Output, cPriors_Output).Value2 = errorMsg
        sht.Cells(rPriors_Output, cPriors_Output).Font.Color = vbRed
    Else
        sht.Cells(rPriors_Output, cPriors_Output).Value2 = "Enter estimated win rates below"
        sht.Cells(rPriors_Output, cPriors_Output).Font.Color = vbBlack
    End If
End Function

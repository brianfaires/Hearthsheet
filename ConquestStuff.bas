Attribute VB_Name = "ConquestStuff"
Option Explicit

Public Function ConquestMulti()
    Dim myLineup() As String
    Dim myBans() As String
    Dim oppLineup() As String
    Dim oppBans() As String
    Dim myThree() As String
    Dim oppThree() As String
    
    Application.ScreenUpdating = False
    Sheets("Conquest").Unprotect
    Sheets("Conquest").Range("G4:N33") = ""
    
    myLineup = LoadLineup(3)
    myBans = GetBans(3)
    
    For i = 4 To 33
        If Sheets("Conquest").Cells(i, 2) <> "" Then
            oppLineup = LoadLineup(i)
            oppBans = GetBans(i)
            myThree = ProcessBans(myLineup, oppBans)
            oppThree = ProcessBans(oppLineup, myBans)
            ProcessMatchup i, myThree, oppThree
        End If
    Next
    
    Sheets("Conquest").Range("G37:N37") = ""
    Sheets("Conquest").Protect
    Application.ScreenUpdating = True
End Function

Public Function LoadLineup(ByVal rowNum As Integer) As String()
    Dim retVal(0 To 3) As String
    For i = 0 To 3
        retVal(i) = FixDeckName(Sheets("Conquest").Cells(rowNum, 2 + i))
    Next
    
    LoadLineup = retVal
End Function

Public Function GetBans(ByVal rowNum As Integer) As String()
    Dim retVal(0 To 8) As String
    
    txt = Sheets("Conquest").Cells(rowNum, 6)
    tokens = Split(txt, ",")
    For i = 0 To UBound(tokens)
        retVal(i) = Trim(tokens(i))
    Next i
    
    GetBans = retVal
End Function

Public Function ProcessBans(lineup() As String, bans() As String) As String()
    Dim retVal(0 To 2) As String
    
    ban = ""
    For i = 0 To UBound(bans)
        If ban = "" Then
            For j = 0 To UBound(lineup)
                If EndsWith(lineup(j), bans(i)) Then
                    ban = bans(i)
                End If
            Next
        End If
    Next
    
    If ban <> "" Then
        deckCount = 0
        For j = 0 To UBound(lineup)
            If Not EndsWith(lineup(j), ban) Then
                retVal(deckCount) = lineup(j)
                deckCount = deckCount + 1
            End If
        Next
    End If
    
    ProcessBans = retVal
End Function

Sub ProcessMatchup(ByVal rowNum As Integer, myLineup() As String, oppLineup() As String)
    Dim shtConquest
    Set shtConquest = Sheets("Conquest")
    
    For i = 0 To 2
        shtConquest.Cells(36, 3 + i) = oppLineup(i)
        shtConquest.Cells(37 + i, 2) = myLineup(i)
    Next
    
    UpdateConquestWinRates
    WriteConquestData rowNum
End Sub

Public Function ConquestProcessMatchup()
    Application.ScreenUpdating = False
    Sheets("Conquest").Unprotect
    
    WriteConquestData 36
    
    Sheets("Conquest").Protect
    Application.ScreenUpdating = True
End Function

Public Function UpdateConquestWinRates_Button()
    Application.ScreenUpdating = False
    'sheets("Conquest").Unprotect

    UpdateConquestWinRates
    
    'sheets("Conquest").Protect
    Application.ScreenUpdating = True
End Function

Public Function UpdateConquestWinRates()
    For Row = 37 To 39
            myDeck = Sheets("Conquest").Cells(Row, 2)
        For col = 3 To 5
            oppDeck = Sheets("Conquest").Cells(36, col)
            Sheets("Conquest").Cells(Row, col) = GetPrior(myDeck, oppDeck)
        Next
    Next
End Function

'NOTE: Update this to check the Deck sheets instead of just pulling priors
Public Function GetPrior(ByVal myDeck As String, ByVal oppDeck As String) As Double
    Row = 4 + 3 * GetColOffset(myDeck) + GetRowOffset(myDeck)
    col = 4 + 3 * GetColOffset(oppDeck) + GetRowOffset(oppDeck)
    
    cellVal = Sheets("Priors").Cells(Row, col)
    If cellVal = "" Then
        GetPrior = 0.5
    Else
        GetPrior = CDbl(cellVal)
    End If
End Function

Public Function WriteConquestData(rowNum As Integer)
    Dim shtHidden
    Dim shtConquest
    Set shtHidden = Sheets("ConquestHidden")
    Set shtConquest = Sheets("Conquest")
    
    shtConquest.Cells(rowNum, 7) = Round(shtHidden.Cells(1, 8), 3)
    
    strat = Round(shtHidden.Cells(3, 12), 2) & ", " & Round(shtHidden.Cells(4, 12), 2) & ", " & Round(shtHidden.Cells(5, 12), 2)
    shtConquest.Cells(rowNum, 8) = strat
    
    shtConquest.Cells(rowNum, 9) = Round(shtHidden.Cells(37, 5), 2) & ", " & Round(shtHidden.Cells(38, 5), 2)
    shtConquest.Cells(rowNum, 10) = Round(shtHidden.Cells(37, 11), 2) & ", " & Round(shtHidden.Cells(38, 11), 2)
    shtConquest.Cells(rowNum, 11) = Round(shtHidden.Cells(37, 17), 2) & ", " & Round(shtHidden.Cells(38, 17), 2)
    
    strat = Round(shtHidden.Cells(151, 4), 2) & ", " & Round(shtHidden.Cells(152, 4), 2) & ", " & Round(shtHidden.Cells(153, 4), 2)
    shtConquest.Cells(rowNum, 12) = strat

    strat = Round(shtHidden.Cells(151, 10), 2) & ", " & Round(shtHidden.Cells(152, 10), 2) & ", " & Round(shtHidden.Cells(153, 10), 2)
    shtConquest.Cells(rowNum, 13) = strat
    
    strat = Round(shtHidden.Cells(151, 16), 2) & ", " & Round(shtHidden.Cells(152, 16), 2) & ", " & Round(shtHidden.Cells(153, 16), 2)
    shtConquest.Cells(rowNum, 14) = strat
End Function

Public Function ComputeBansForOneMatchup()
    Dim shtConquest
    Set shtConquest = Sheets("Conquest")
    
    Application.ScreenUpdating = False
    Sheets("Conquest").Unprotect
    
    For i = 0 To 3
        For j = 0 To 3
            myBan = shtConquest.Cells(45, 2 + i)
            oppBan = shtConquest.Cells(44, 2 + j)
            shtConquest.Cells(48 + 4 * i + j, 2) = myBan
            shtConquest.Cells(48 + 4 * i + j, 3) = oppBan
            
            'Fill in Deck names
            curOffset = 0
            For k = 0 To 3
                If shtConquest.Cells(44, 2 + k) <> oppBan Then
                    shtConquest.Cells(37 + curOffset, 2) = shtConquest.Cells(44, 2 + k)
                    curOffset = curOffset + 1
                End If
            Next
            
            curOffset = 0
            For k = 0 To 3
                If shtConquest.Cells(45, 2 + k) <> myBan Then
                    shtConquest.Cells(36, 3 + curOffset) = shtConquest.Cells(45, 2 + k)
                    curOffset = curOffset + 1
                End If
            Next
            
            UpdateConquestWinRates
            WriteConquestData 36
            shtConquest.Cells(48 + 4 * i + j, 4) = shtConquest.Cells(36, 7)
        Next
    Next
    
    Sheets("Conquest").Protect
    Application.ScreenUpdating = True
End Function


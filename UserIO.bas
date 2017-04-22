Attribute VB_Name = "UserIO"
Option Explicit
' This module contains all the access points for the user. It should only contain subs and be the only module with Subs (to hide functionality on the Macros list)

' This is fired when the user presses the big button on the Log sheet. It will process all logs and update the deck sheets.
Sub Click_ProcessLogs()
    ' Disable visibility and calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ReadSettings
    SetProtectionAllDeckSheets False
    ClearDeckSheetGames
    CreateMissingDeckSheets
    ProcessLogs
    
    ' Re-enable early because Cleanup depends on it working
    Application.Calculation = xlCalculationAutomatic
    CleanUpDeckSheets
    SetProtectionAllDeckSheets True
    
    ' Restore visibility/calculation settings
    Sheets("Meta").Select
    Application.ScreenUpdating = True
End Sub

' This is fired when the user changes the number of games at the default win rate. This will change the win rates and best matchups.
' This fires when the user clicks the "Recompute" button on each deck sheet.
' TODO: Should this also fire when the user edits win/loss counts on deck sheets?
Sub RecalculateBestMatchupsForCurrentSheet()
    ' Disable visibility and calculation
    If IsDeckSheet(ActiveSheet.Name) Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        ActiveSheet.Unprotect
        
        ComputeBestMatchups ActiveSheet.Name
        
        ' Restore visibility/calculation settings
        ActiveSheet.Protect
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
End Sub

' This is fired when the user clicks the "Cleanup" button on the Priors tab
Sub Click_CleanUpPriors()
    ' Disable visibility and calculation
    Dim sht As Worksheet
    Set sht = Sheets("Priors")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    sht.Unprotect
    
    'Adjust hidden columns
    SetPriorVisibility
    
    ' Various checks, keep track of how many errors
    Dim errorsFound As Integer
    errorsFound = CheckPriorDiagonal
    errorsFound = errorsFound + CheckPriorHalfs
    errorsFound = errorsFound + CheckPriorValues

    DisplayPriorErrors errorsFound
    
    ' Restore visibility/calculation settings
    sht.Protect
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' TODO: Turn off calculations where allowable. This function is way slower than it should be.
Public Sub Click_UpdateMetaFromLogs()
    ' Disable visibility and calculation
    Application.ScreenUpdating = False
    Sheets("Meta").Unprotect
    
    ClearMetaData
    LoadMetaFromLogs
    
    UpdateMostPlayedClasses
    UpdateMostPlayedDecks
    UpdateAllBestMatchups
    UpdateBestMetaDecks
    
    ' Restore visibility/calculation settings
    Sheets("Meta").Protect
    Application.ScreenUpdating = True
End Sub

Public Sub Click_UpdateMetaFromMetaSheet()
    'Disable visibility And Calculation
    Application.ScreenUpdating = False
    Sheets("Meta").Unprotect
    
    ClearClassMatchups
    
    UpdateMostPlayedClasses
    UpdateMostPlayedDecks
    UpdateAllBestMatchups
    UpdateBestMetaDecks
    
    ' Restore visibility/calculation settings
    Sheets("Meta").Protect
    Application.ScreenUpdating = True
End Sub

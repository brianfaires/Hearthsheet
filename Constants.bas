Attribute VB_Name = "Constants"
Option Explicit
' This file contains all the locations of various tables and lists throughout the sheet

''''Globals
Public Const nClasses = 9
Public Const nDecksPerClass = 6
Public Const nDecks = nClasses * nDecksPerClass
Public Const nMetaDecks = 25

''''Log Sheet
'Settings
Public Const cLogCfg = 9
Public Const rLogCfg_Date = 2
Public Const rLogCfg_MyMinRank = 3
Public Const rLogCfg_MyMaxRank = 4
Public Const rLogCfg_OppMinRank = 5
Public Const rLogCfg_OppMaxRank = 6
'Deck Grid (on Log sheet)
Public Const cDeckGrid = 11
Public Const rDeckGrid = 2
'Log rows; columns for each game
Public Const rLogs = 9
Public Const cLogs_Date = 1
Public Const cLogs_MyDeck = 2
Public Const cLogs_MyRank = 3
Public Const cLogs_OppDeck = 4
Public Const cLogs_OppRank = 5
Public Const cLogs_Won = 6
Public Const cLogs_Notes = 7
' Deck sheets settings
Public Const cDeckSheetCfg = 6
Public Const rDeckSheetCfg_Red = 3
Public Const rDeckSheetCfg_Yellow = 4
Public Const rDeckSheetCfg_Black = 5

''''Priors Sheet
Public Const cPriors_Output = 4
Public Const rPriors_Output = 1
Public Const cPriors = 3
Public Const rPriors = 3

''''Meta Sheet
'Config Section for Log-based meta
Public Const cMetaCfg = 4
Public Const rMetaCfg_MaxGames = 34
Public Const rMetaCfg_MinDate = 35
Public Const rMetaCfg_MaxDate = 36
Public Const rMetaCfg_MyMinRank = 37
Public Const rMetaCfg_MyMaxRank = 38
Public Const rMetaCfg_OppMinRank = 39
Public Const rMetaCfg_OppMaxRank = 40
'Table with counts of each deck
Public Const cMetaCounts = 2
Public Const rMetaCounts = 11
'Table with percents of each deck
Public Const cMetaPerc = 2
Public Const rMetaPerc_ClassName = 21
Public Const rMetaPerc_Deck = 22
Public Const rMetaPerc_Class = 28
'Table of class matchups
Public Const cMetaClassMatchups = 8
Public Const rMetaClassMatchups = 32
'Table of most played classes
Public Const rMetaMPC = 32
Public Const cMetaMPC_Name = 21
Public Const cMetaMPC_Value = 22
'Master meta table (copied by deck sheets)
Public Const rMMeta = 3
Public Const cMMeta_Name = 21
Public Const cMMeta_Value = 22
'Best decks table
Public Const rBestDecks = 3
Public Const cBestDecks_Name = 24
Public Const cBestDecks_Value = 25
Public Const cBestDecksCfg_MinGames = 25
Public Const rBestDecksCfg_MinGames = 28
'Table of deck names
Public Const cMeta_DeckName = 2
Public Const rMeta_DeckName = 2

''''Conquest Sheet

''''Deck Sheets
'Game count/win rate
Public Const cDeckSheet_GameCount = 5
Public Const rDeckSheet_GameCount = 2
Public Const cDeckSheet_WinRate = 9
Public Const rDeckSheet_WinRate = 2
'Win/Loss Table
Public Const rDeckSheet_WLTable = 3
Public Const cDeckSheet_WLTable = 12
'Expected win rate table
Public Const rDeckSheet_ExpWinRate = 22
Public Const cDeckSheet_ExpWinRate = 12
'Copies of Master meta table
Public Const cDeckSheetMeta_Name = 6
Public Const cDeckSheetMeta_Value = 8
Public Const rDeckSheetMeta = 5
'Best matchups table
Public Const rDeckSheetBM = 5
Public Const cDeckSheetBM_Name = 2
Public Const cDeckSheetBM_Value = 4



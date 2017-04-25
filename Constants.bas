Attribute VB_Name = "Constants"
Option Explicit
''''Globals
Public Const nClasses = 9
Public Const nDecksPerClass = 6
Public Const nDecks = nClasses * nDecksPerClass
Public Const nMetaDecks = 25

'Log config columns, to index into a single row instead of several diff ranges
Public Const LogCfgCol_Date = 1
Public Const LogCfgCol_MyDeck = 2
Public Const LogCfgCol_MyRank = 3
Public Const LogCfgCol_OppDeck = 4
Public Const LogCfgCol_OppRank = 5
Public Const LogCfgCol_Win = 6
Public Const LogCfgCol_Notes = 7


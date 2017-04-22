Attribute VB_Name = "Util"
Option Explicit

Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Public Function StartsWith(str As String, prefix As String) As Boolean
     Dim startingLen As Integer
     startingLen = Len(prefix)
     StartsWith = (Left(Trim(UCase(str)), startingLen) = UCase(prefix))
End Function

Public Function SheetExists(sheetName As String) As Boolean
    Dim sht As Worksheet
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function

Public Function IsDeckSheet(sheetName As String) As Boolean
    Dim d As Deck
    Set d = Factory.CreateDeck(sheetName)
    IsDeckSheet = d.IsValid
End Function

Public Function CellToInt(Cell As Range) As Integer
    On Error GoTo NOT_AN_INTEGER
    ConvertToInteger = CInt(Cell.Value2)
    Exit Function
NOT_AN_INTEGER:
    ConvertToInteger = 0
End Function

Public Function IntToColName(ByVal i As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, i).Address(True, False), "$")
    IntToColName = vArr(0)
    vArr = Nothing
End Function


Attribute VB_Name = "Module90_Public_Sub"
Option Explicit

Function GetWorkPath() As String

GetWorkPath = ""

Dim sTEAMPath As String
'sTEAMPath = "https://psav-my.sharepoint.com/personal/jlu_psav_com/Documents"
sTEAMPath = "https://mckessoncorpca-my.sharepoint.com/personal/" & NameHTTP & "_mckesson_ca/Documents"

Dim sNormPath As String
'sNormPath = "C:\Users\JiataoLu\OneDrive - Encore"
sNormPath = "C:\Users\" & NameHD & "\OneDrive - McKesson Corporation"

Dim sPath As String
sPath = ThisWorkbook.Path

If UCase(Left(sPath, 2)) = "C:" Then
    GetWorkPath = sPath
    Exit Function
End If

sPath = Replace(sPath, sTEAMPath, sNormPath)
'Debug.Print sPath
sPath = Replace(sPath, "/", "\")
'Debug.Print sPath

GetWorkPath = sPath

End Function



Sub DeleteUnusedFormats()
    Dim lLastRow As Long, lLastColumn As Long
    Dim lRealLastRow As Long, lRealLastColumn As Long
    
    With Range("a1").SpecialCells(xlCellTypeLastCell)
        lLastRow = .Row
        lLastColumn = .Column
    End With
    lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
    lRealLastColumn = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
    
    If lRealLastRow < lLastRow Then
        Range(Cells(lRealLastRow + 1, 1), Cells(lLastRow, 1)).EntireRow.Delete
    End If
    If lRealLastColumn < lLastColumn Then
        Range(Cells(1, lRealLastColumn + 1), Cells(1, lLastColumn)).EntireColumn.Delete
    End If
    ActiveSheet.UsedRange
    
End Sub

'Delete all data, only leave header name in first row

Sub Initialize_Sheet(SheetName As String)
'Dim SheetName As String
'SheetName = "FIS & PeopleSoft"

Dim wsWork As Worksheet
Dim iLastRowWorkSheet As Long

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsWork = Worksheets(SheetName)
wsWork.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iLastRowWorkSheet = lRealLastRow
If iLastRowWorkSheet < 2 Then Exit Sub

Rows("2:" & iLastRowWorkSheet).Delete
Call DeleteUnusedFormats

End Sub

Function Remove_Leading_Zero(InputNumber)
'Dim InputNumber As String
'InputNumber = "001234566"

Dim sInput As String
Dim iLen As Integer
Dim i As Integer
Dim sLeftLetter As String

sInput = InputNumber
sLeftLetter = Left(sInput, 1)
iLen = Len(sInput)

While iLen > 1 And sLeftLetter = "0"
    sInput = Right(sInput, iLen - 1)
    iLen = Len(sInput)
    sLeftLetter = Left(sInput, 1)
Wend

Remove_Leading_Zero = sInput

End Function

Function Read_BUGL(OrigInfo As String)
Read_BUGL = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

strPattern = "\d+"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    Read_BUGL = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function



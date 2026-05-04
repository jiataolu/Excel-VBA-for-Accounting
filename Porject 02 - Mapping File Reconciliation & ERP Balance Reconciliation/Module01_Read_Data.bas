Attribute VB_Name = "Module01_Read_Data"
Option Explicit

Sub Read_Data_from_Mapping_and_FCCS_Report_and_Treasury_Report()

Dim wkbFmWorkBook As Workbook
Dim wsFmWorkSheet As Worksheet

Dim wkbToWorkBook As Workbook
Dim wsToWorkSheet As Worksheet

'Read FCCS data: 3 sheets - FCCS (SAP), JDE, NetSuite
'Read FCCS-SAP
Set wkbToWorkBook = ThisWorkbook
wkbToWorkBook.Activate
Set wsToWorkSheet = Worksheets("FCCS")

Set wkbFmWorkBook = Workbooks.Open(GetWorkPath & "\" & FileNameFCCS, UpdateLinks:=0)
wkbFmWorkBook.Activate
Set wsFmWorkSheet = Worksheets("SAP")
Call Read_One_sheet(wkbFmWorkBook, wsFmWorkSheet, wkbToWorkBook, wsToWorkSheet)

'Read JDE
wkbToWorkBook.Activate
Set wsToWorkSheet = Worksheets("JDE")

wkbFmWorkBook.Activate
Set wsFmWorkSheet = Worksheets("JDE")
Call Read_One_sheet(wkbFmWorkBook, wsFmWorkSheet, wkbToWorkBook, wsToWorkSheet)


'Read NetSuite
wkbToWorkBook.Activate
Set wsToWorkSheet = Worksheets("NetSuite")

wkbFmWorkBook.Activate

Set wsFmWorkSheet = Worksheets("NetSuite")
Call Read_One_sheet(wkbFmWorkBook, wsFmWorkSheet, wkbToWorkBook, wsToWorkSheet)

wkbFmWorkBook.Close savechanges:=False

'Read Mapping File
wkbToWorkBook.Activate
Set wsToWorkSheet = Worksheets("Mapping")

Set wkbFmWorkBook = Workbooks.Open(GetWorkPath & "\" & FileNameMacroOne, UpdateLinks:=0)
wkbFmWorkBook.Activate
Set wsFmWorkSheet = Worksheets("Mapping Consolidated")
Call Read_One_sheet(wkbFmWorkBook, wsFmWorkSheet, wkbToWorkBook, wsToWorkSheet)

wkbFmWorkBook.Close savechanges:=False

Call Read_FIS

End Sub

Sub Read_One_sheet(fmWorkBook As Workbook, fmWorkSheet As Worksheet, toWorkBook As Workbook, toWorkSheet As Worksheet)

'Dim fmWorkBook As Workbook
'Dim fmWorkSheet As Worksheet

'Dim toWorkBook As Workbook
'Dim toWorkSheet As Worksheet

'Set fmWorkBook = ThisWorkbook

Dim rngCopy As Range
Dim rngPaste As Range

'Set toWorkBook = ThisWorkbook
'toWorkBook.Activate
'Set toWorkSheet = Worksheets("FCCS")
toWorkBook.Activate
toWorkSheet.Select
Cells.Select
Selection.Delete
Selection.ClearFormats
Cells(1, 1).Select

Set rngPaste = Cells(1, 1)

'Set fmWorkBook = Workbooks.Open(GetWorkPath & "\" & "To read 02 - FCCS-SAP-JDE-NetSuite.xlsx")
'fmWorkBook.Activate
'Set fmWorkSheet = Worksheets("SAP")
fmWorkBook.Activate
fmWorkSheet.Select
Cells.Select
Set rngCopy = Selection

rngCopy.Copy Destination:=rngPaste
Cells(1, 1).Select

toWorkBook.Activate
toWorkSheet.Select
Call DeleteUnusedFormats

End Sub


Sub Read_FIS()
Application.DisplayAlerts = False

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsFIS As Worksheet
Dim iMaxRowFIS As Integer
Dim str4HeaderCombined As String
Dim strCell As String
Dim iStartRowHeader As Integer
Dim iTotalRow As Integer
Dim iCountAcct As Integer

Dim i As Integer
Dim j As Integer

Dim wkbCashPosition As Workbook
Dim wsFormatting As Worksheet
Dim iFoundwsFormatting As Integer
Dim sFileNameTreasury As String
Dim iFormattingTotalRow As Integer
Dim iMaxRowFormatting As Integer
Dim iMaxColFormatting As Integer


Dim rngCopy As Range
Dim rngPaste As Range

'Clear FIS Sheet
ThisWorkbook.Activate
Set wsFIS = Worksheets("FIS")
wsFIS.Select
Cells.Select
Selection.Delete
Selection.ClearFormats
Cells(1, 1).Select
Set rngPaste = Cells(1, 1)

sFileNameTreasury = GetWorkPath & "\" & FileNameTreausry
'Debug.Print sFileNameTreasury
Set wkbCashPosition = Workbooks.Open(sFileNameTreasury)

'Debug.Print Sheets.Count()
iFoundwsFormatting = 0
For i = 1 To Sheets.Count()
    If Sheets(i).Name = "Formatting" Then
        iFoundwsFormatting = 1
        Exit For
    End If
Next i
If iFoundwsFormatting = 0 Then
    MsgBox "The file from Treasury, Sheet-""Formatting"" is missing."
    Exit Sub
End If


Set wsFormatting = wkbCashPosition.Worksheets("Formatting")
wsFormatting.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFormatting = lRealLastRow
iMaxColFormatting = lRealLastCol
If iMaxRowFormatting < 2 Then
    MsgBox "There is some problem with the file from Treasury, Sheet-Formatting."
    Exit Sub
End If

iFormattingTotalRow = 0
For i = 1 To iMaxRowFormatting
    strCell = Cells(i, 1)
    strCell = Replace(strCell, " ", "")
    strCell = UCase(strCell)
    If InStr(strCell, "TOTAL") > 0 Then
        iFormattingTotalRow = i
        Exit For
    End If
Next i
If iFormattingTotalRow = 0 Then
    MsgBox "The file from Treasury, Sheet-Formatting is missing ""Total"" Lines."
    Exit Sub
End If
'Debug.Print iFormattingTotalRow

Set rngCopy = Range(Cells(1, 1), Cells(iFormattingTotalRow, iMaxColFormatting))
rngCopy.Copy
rngPaste.PasteSpecial xlPasteValues
rngPaste.PasteSpecial xlPasteFormats
'rngCopy.Copy Destination:=rngPaste

Application.CutCopyMode = False
wkbCashPosition.Close savechanges:=False
wsFIS.Select
Cells.Select
Selection.EntireColumn.AutoFit
Cells(1, 1).Select

Call DeleteUnusedFormats

'lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
'lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
'iMaxRowFIS = lRealLastRow


'find header row
'iStartRowHeader = 0

'For i = 1 To iMaxRowFIS
'    str4HeaderCombined = ""
'    For j = 1 To 4
'        str4HeaderCombined = str4HeaderCombined & Cells(i, j)
'    Next j
'    str4HeaderCombined = Replace(str4HeaderCombined, " ", "")
    'Debug.Print str4HeaderCombined
'    If str4HeaderCombined = FIS4Header Then
'        iStartRowHeader = i
'        Exit For
'    End If
'Next i

'If iStartRowHeader = 0 Then
'    MsgBox "Please check FIS Sheet for columns"
'    Exit Sub
'End If


'find last row, which is Total
'iTotalRow = 0
'For i = iStartRowHeader To iMaxRowFIS
'    Cells(i, 1) = Replace(Cells(i, 1), " ", "")
'    If UCase(Cells(i, 1)) = "TOTAL" Then
'        iTotalRow = i
'        Exit For
'    End If
'Next i

'If iTotalRow = 0 Then
'    Cells(iMaxRowFIS + 2, 1) = "Total"
'    iMaxRowFIS = iMaxRowFIS + 2
'    iTotalRow = iMaxRowFIS
'End If

'count the number of account
'iCountAcct = 0
'For i = iStartRowHeader + 1 To iTotalRow - 1
'    strCell = Cells(i, 1)
'    strCell = Replace(strCell, " ", "")
'    If strCell <> "" Then iCountAcct = iCountAcct + 1
'Next i
'Cells(iTotalRow, 2) = CStr(iCountAcct)


Application.DisplayAlerts = True

End Sub




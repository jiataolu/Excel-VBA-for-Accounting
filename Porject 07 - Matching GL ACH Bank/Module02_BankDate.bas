Attribute VB_Name = "Module02_BankDate"
Option Explicit

Sub Find_Unique_Bank_Date()

Dim wsBankDate As Worksheet
Dim wsDataBank As Worksheet
Dim rngCopy As Range
Dim rngPaste As Range

Dim iMaxRowBankDate As Integer
Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsBankDate = Worksheets(SheetNameBankDate)
wsBankDate.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select

Set rngPaste = Cells(1, 1)

Set wsDataBank = Worksheets(SheetNameDataBank)
wsDataBank.Select
Set rngCopy = Columns(ColDataBankValueDate)

rngCopy.Copy Destination:=rngPaste
wsBankDate.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBankDate = lRealLastRow
wsBankDate.Range("A1:A" & iMaxRowBankDate).RemoveDuplicates Columns:=1, Header:=xlYes

Rows(1).Delete
Call DeleteUnusedFormats
End Sub

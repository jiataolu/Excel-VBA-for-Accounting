Attribute VB_Name = "Module12_OutputReport"
Option Explicit


Sub Output_report(CompanyName As String)

'Dim CompanyName As String
'CompanyName = "MSD"

Dim wkbReport As Workbook
Dim strFileFullReport As String


Dim wkbMacro As Workbook


strFileFullReport = GetWorkPath & "\" & SubFolderOutput & "\" & CompanyName & " PAP clearing.xlsx"

Set wkbMacro = ThisWorkbook

Set wkbReport = Workbooks.Open(strFileFullReport)

Call Copy_Paste_Sheet(wkbMacro, wkbReport, "Bank Statement")
Call Copy_Paste_Sheet(wkbMacro, wkbReport, "FBL5N")
Call Copy_Paste_Sheet(wkbMacro, wkbReport, "PAP Invoices")
Call Copy_Paste_Sheet_Keep_Formula(wkbMacro, wkbReport, "Validation")

If CompanyName = "SPS" Then
    Call Copy_Paste_Sheet(wkbMacro, wkbReport, "DISCOUNT INFO")
End If

wkbReport.Close SaveChanges:=True

End Sub

Sub Copy_Paste_Sheet(FromWorkBook As Workbook, ToWorkBook As Workbook, SheetName As String)

Dim wsFrom As Worksheet
Dim wsTo As Worksheet
Dim rngCopy As Range
Dim rngPaste As Range

ToWorkBook.Activate
Set wsTo = Worksheets(SheetName)
wsTo.Select
Cells.Select
Selection.Delete
Set rngPaste = Cells(1, 1)

FromWorkBook.Activate
Set wsFrom = Worksheets(SheetName)
wsFrom.Select
Cells.Select
Set rngCopy = Selection

rngCopy.Copy Destination:=rngPaste

FromWorkBook.Activate
End Sub

Sub Copy_Paste_Sheet_Keep_Formula(FromWorkBook As Workbook, ToWorkBook As Workbook, SheetName As String)

Dim wsFrom As Worksheet
Dim iMaxFrom As Integer
Dim jMaxFrom As Integer

Dim iFrom As Integer
Dim jFrom As Integer

Dim strCell As String

Dim wsTo As Worksheet

Dim lRealLastRow As Long
Dim lRealLastCol As Long
'Dim rngCopy As Range
'Dim rngPaste As Range

ToWorkBook.Activate
Set wsTo = Worksheets(SheetName)
wsTo.Select
Cells.Select
Selection.Delete
'Set rngPaste = Cells(1, 1)

FromWorkBook.Activate
Set wsFrom = Worksheets(SheetName)
wsFrom.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxFrom = lRealLastRow
jMaxFrom = lRealLastCol

For iFrom = 1 To iMaxFrom
    For jFrom = 1 To jMaxFrom
        strCell = wsFrom.Cells(iFrom, jFrom).Formula
        wsTo.Cells(iFrom, jFrom).Value = strCell
    Next jFrom
Next iFrom

ToWorkBook.Activate
wsTo.Select
Columns("B").AutoFit
Cells(4, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Cells(5, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Cells(7, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

FromWorkBook.Activate
End Sub


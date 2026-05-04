Attribute VB_Name = "Module11_SheetValidation"
Option Explicit

Sub Make_Validation_Sheet(CompanyName As String)

'Dim CompanyName As String
'CompanyName = "SPS"

Dim wsBS As Worksheet
Dim iMaxBS As Integer
Dim iBS As Integer
Dim strFormula As String

Dim wsValid As Worksheet

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsBS = Worksheets("Bank Statement")
wsBS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxBS = lRealLastRow

strFormula = "="
For iBS = 2 To iMaxBS
    If Cells(iBS, ColBSEntity) = CompanyName Then
        strFormula = strFormula & "'Bank Statement'!F" & CStr(iBS) & "+"
    End If
Next iBS
strFormula = Left(strFormula, Len(strFormula) - 1)

Set wsValid = Worksheets("Validation")
wsValid.Select
Cells.Select
Selection.Delete

Cells(4, 2) = "Bank Statement"
Cells(5, 2) = "SAP Invoices"
Cells(7, 2) = "Difference"

Cells(4, 3) = strFormula
Cells(5, 3) = "=SUM('PAP Invoices'!K:K)"
Cells(7, 3) = "=C4-C5"

Columns("B").AutoFit
Cells(4, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Cells(5, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Cells(7, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
End Sub

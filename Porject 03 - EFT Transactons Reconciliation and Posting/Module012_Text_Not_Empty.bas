Attribute VB_Name = "Module012_Text_Not_Empty"
Option Explicit

'Check Text Field, if it is empty, then add / sign
'Force Text Field non-empty, otherwise offset step will have problem.
Sub Text_Field_Can_Not_Be_Empty()

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim sText As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow

If iMaxRowSAP < 2 Then Exit Sub

For iRowSAP = 2 To iMaxRowSAP
    sText = wsSAP.Cells(iRowSAP, iColSAPText)
    If Replace(sText, " ", "") = "" Then wsSAP.Cells(iRowSAP, iColSAPText) = "/"
Next iRowSAP


End Sub


Attribute VB_Name = "Module10_ClearPivotTSheet"
Option Explicit

Sub Generate_Pivot_Table_01_Clear_Pivot_Sheet()

Dim wsPT As Worksheet

Set wsPT = Worksheets(SheetNamePivotTableGLBank)
wsPT.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select
End Sub



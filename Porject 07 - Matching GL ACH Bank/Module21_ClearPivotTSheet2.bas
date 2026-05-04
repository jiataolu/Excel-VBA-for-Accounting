Attribute VB_Name = "Module21_ClearPivotTSheet2"
Option Explicit


Sub Generate_Pivot_Table_01_Clear_Pivot_Sheet_GL_ACH()

Dim wsPT As Worksheet

Set wsPT = Worksheets(SheetNamePivotTableGLACH)
wsPT.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select
End Sub


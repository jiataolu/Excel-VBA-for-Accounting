Attribute VB_Name = "Module04_PivotTable"
Option Explicit

Sub Generate_Pivot_Table()

Dim ws03_PT As Worksheet

Dim ws02JEData As Worksheet
Dim iMaxRowData As Long
Dim iMaxColData As Integer
Dim rngData As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim PTCacheFinal As PivotCache
Dim PTTableFinal As PivotTable
Dim sNamePivotTable As String

Set ws03_PT = Worksheets(Sheet04Name_Pivot)
ws03_PT.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select

'Create Pivot table, data from "02-Data for JE", in "03-Pivot"
Set ws02JEData = Worksheets(Sheet03Name_JEDataClean1ZBA)
ws02JEData.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowData = lRealLastRow
iMaxColData = lRealLastCol

Set rngData = Range(Cells(1, 1), Cells(iMaxRowData, iMaxColData))

'define Pivot Tabel Cache
Set PTCacheFinal = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData)
sNamePivotTable = "WD01"
Set PTTableFinal = PTCacheFinal.CreatePivotTable(TableDestination:=ws03_PT.Cells(1, 1), TableName:=sNamePivotTable)


'add Row (and Columns)
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("BU_1")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 1
End With


'add Row (and Columns)
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("Bank_Code_1")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 2
End With

'add Row (and Columns)
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("GL_1")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 3
End With

'add Row (and Columns)
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("Bank_Code_2")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 4
End With

'add Row (and Columns)
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("BU_2")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 5
End With

'add Row (and Columns)
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("GL_2")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 6
End With


'add Row (and Columns)
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("Ccy")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 7
End With


'Add data into PivotTable
With ws03_PT.PivotTables(sNamePivotTable).PivotFields("Amount_ADJ")
    .Orientation = xlDataField
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Total Amount"
End With

' Set the report layout to tabular form
' Repeat all item labels
With ws03_PT.PivotTables(sNamePivotTable)
    .RowAxisLayout xlTabularRow
    .RepeatAllLabels xlRepeatLabels
End With

'Do not show subtotal
With ws03_PT.PivotTables(sNamePivotTable)
    .PivotFields("Bank_Code_1").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("BU_1").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("GL_1").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Bank_Code_2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("BU_2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("GL_2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With


ws03_PT.Select
Columns("A:F").HorizontalAlignment = xlCenter
Columns("G:G").Style = "Comma"
Cells.Select
Columns.AutoFit
Cells(1, 1).Select

End Sub

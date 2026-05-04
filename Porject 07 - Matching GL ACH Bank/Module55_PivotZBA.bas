Attribute VB_Name = "Module55_PivotZBA"
Option Explicit

Sub Generate_Pivot_Table()

Dim wsPT As Worksheet

Dim wsZBA As Worksheet
Dim iMaxRowZBA As Long
Dim iMaxColZBA As Integer
Dim rngZBA As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim PTCacheFinal As PivotCache
Dim PTTableFinal As PivotTable
Dim sNamePivotTable As String

Set wsPT = Worksheets(SheetNamePivotZBA)
wsPT.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select

'Create Pivot table, data from "02-Data for JE", in "03-Pivot"
Set wsZBA = Worksheets(SheetNameKyribaZBAMMS)
wsZBA.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowZBA = lRealLastRow
iMaxColZBA = lRealLastCol

Set rngZBA = Range(Cells(1, 1), Cells(iMaxRowZBA, iMaxColZBA))

'define Pivot Tabel Cache
Set PTCacheFinal = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngZBA)
sNamePivotTable = "WD01"
Set PTTableFinal = PTCacheFinal.CreatePivotTable(TableDestination:=wsPT.Cells(1, 1), TableName:=sNamePivotTable)


'add Row (and Columns)
With wsPT.PivotTables(sNamePivotTable).PivotFields("Account")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 1
End With


'add Row (and Columns)
With wsPT.PivotTables(sNamePivotTable).PivotFields("Concentration-BU")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 2
End With

'add Row (and Columns)
With wsPT.PivotTables(sNamePivotTable).PivotFields("Concentration-GL")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 3
End With

'add Row (and Columns)
With wsPT.PivotTables(sNamePivotTable).PivotFields("Offset Kyriba Code")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 4
End With

'add filter


'add Row (and Columns)
With wsPT.PivotTables(sNamePivotTable).PivotFields("Offset-BU")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 5
End With

'add Row (and Columns)
With wsPT.PivotTables(sNamePivotTable).PivotFields("Offset-GL")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 6
End With


''add Row (and Columns)
With wsPT.PivotTables(sNamePivotTable).PivotFields("Account cur.")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 7
End With


'Add data into PivotTable
With wsPT.PivotTables(sNamePivotTable).PivotFields("Net Amount")
    .Orientation = xlDataField
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Total Amount"
End With

' Set the report layout to tabular form
' Repeat all item labels
With wsPT.PivotTables(sNamePivotTable)
    .RowAxisLayout xlTabularRow
    .RepeatAllLabels xlRepeatLabels
End With

'Do not show subtotal
With wsPT.PivotTables(sNamePivotTable)
    .PivotFields("Account").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Concentration-BU").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Concentration-GL").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Offset Kyriba Code").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Offset-BU").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Offset-GL").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With


wsPT.Select
Columns("A:F").HorizontalAlignment = xlCenter
Columns("G:G").Style = "Comma"
Cells.Select
Columns.AutoFit
Columns("A:F").ColumnWidth = 15
Columns("G").ColumnWidth = 7
Range(Cells(1, 1), Cells(1, 3)).Interior.ColorIndex = 38
Range(Cells(1, 4), Cells(1, 6)).Interior.ColorIndex = 43
Cells(1, 1).Select
ActiveWindow.FreezePanes = True


End Sub


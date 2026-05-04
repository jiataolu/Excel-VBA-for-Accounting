Attribute VB_Name = "Module14_PivotTBank"
Option Explicit

Sub Generate_Pivot_Table_03_Bank()

Dim wsPT As Worksheet
Dim iColPT As Integer

Dim wsDataBank As Worksheet
Dim iMaxRowDataBank As Long
Dim iMaxColDataBank As Integer
Dim rngDataBank As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim PTCacheFinal As PivotCache
Dim PTTableFinal As PivotTable
Dim sNamePivotTable As String

Set wsPT = Worksheets(SheetNamePivotTableGLBank)
wsPT.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iColPT = lRealLastCol + 4

' 2 - Create Pivot table, data from "Data_Bank"
Set wsDataBank = Worksheets(SheetNameDataBank)
wsDataBank.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataBank = lRealLastRow
iMaxColDataBank = lRealLastCol

Set rngDataBank = Range(Cells(1, 1), Cells(iMaxRowDataBank, iMaxColDataBank))

'define Pivot Tabel Cache
Set PTCacheFinal = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngDataBank)
sNamePivotTable = "WDBank"
Set PTTableFinal = PTCacheFinal.CreatePivotTable(TableDestination:=wsPT.Cells(3, iColPT), TableName:=sNamePivotTable)


'add Row (and Columns)

With wsPT.PivotTables(sNamePivotTable).PivotFields("Trans_Type")
    '.Orientation = xlPageField
    .Orientation = xlRowField
    .Position = 1
    .AutoSort xlAscending, "Trans_Type"
End With

With wsPT.PivotTables(sNamePivotTable).PivotFields("Recon_Date")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .AutoSort xlAscending, "Recon_Date"
    .Position = 2
End With


With wsPT.PivotTables(sNamePivotTable).PivotFields("Flow code")
    '.Orientation = xlRowField
    .Orientation = xlColumnField
    .Position = 1
End With



'Add data into PivotTable
With wsPT.PivotTables(sNamePivotTable).PivotFields("Amount")
    .Orientation = xlDataField
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Sum. of Amount"
End With

' Set the report layout to tabular form
' Repeat all item labels
With wsPT.PivotTables(sNamePivotTable)
    .RowAxisLayout xlTabularRow
    .RepeatAllLabels xlRepeatLabels
End With

'Do not show subtotal
With wsPT.PivotTables(sNamePivotTable)
    .PivotFields("Trans_Type").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Value Date").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("REDEPOSIT YES/ NO").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Flow code").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With

wsPT.Cells(1, iColPT) = "Bank"




wsPT.Select
'Columns("A:F").HorizontalAlignment = xlCenter
'Columns("G:G").Style = "Comma"
'Cells.Select
'Columns.AutoFit
'Columns("A:F").ColumnWidth = 15
'Columns("G").ColumnWidth = 7
'Range(Cells(1, 1), Cells(1, 3)).Interior.ColorIndex = 38
'Range(Cells(1, 4), Cells(1, 6)).Interior.ColorIndex = 43
Cells(1, 1).Select
'ActiveWindow.FreezePanes = True


End Sub





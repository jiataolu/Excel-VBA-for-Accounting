Attribute VB_Name = "Module12_PivotTGL"
Option Explicit

Sub Generate_Pivot_Table_02_GL()

Dim wsPT As Worksheet

Dim wsDataGL As Worksheet
Dim iMaxRowDataGL As Long
Dim iMaxColDataGL As Integer
Dim rngDataGL As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim PTCacheFinal As PivotCache
Dim PTTableFinal As PivotTable
Dim sNamePivotTable As String

Set wsPT = Worksheets(SheetNamePivotTableGLBank)
wsPT.Select
'Cells.Select
'Selection.Delete
'Cells(1, 1).Select


' 1 - Create Pivot table, data from "Data_GL", in "03-Pivot"
Set wsDataGL = Worksheets(SheetNameDataGL)
wsDataGL.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataGL = lRealLastRow
iMaxColDataGL = lRealLastCol

Set rngDataGL = Range(Cells(1, 1), Cells(iMaxRowDataGL, iMaxColDataGL))

'define Pivot Tabel Cache
Set PTCacheFinal = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngDataGL)
sNamePivotTable = "WDGL"
Set PTTableFinal = PTCacheFinal.CreatePivotTable(TableDestination:=wsPT.Cells(3, 1), TableName:=sNamePivotTable)


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


With wsPT.PivotTables(sNamePivotTable).PivotFields("Document Type")
    '.Orientation = xlRowField
    .Orientation = xlColumnField
    .Position = 1
End With



'Add data into PivotTable
With wsPT.PivotTables(sNamePivotTable).PivotFields("Amount in doc. curr.")
    .Orientation = xlDataField
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Sum. of Amount in doc. curr."
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
    .PivotFields("Posting Date").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Return Yes / No").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    .PivotFields("Document Type").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)


End With

wsPT.Cells(1, 1) = "GL"




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




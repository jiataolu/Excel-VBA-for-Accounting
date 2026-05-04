Attribute VB_Name = "Module63_PivotACH"
Option Explicit

Sub Generate_Pivot_Table_04_ACH()

Dim wsPT As Worksheet
Dim iColPT As Integer
Dim irowpt As Integer

Dim wsDataACH1115 As Worksheet
Dim iMaxRowDataACH1115 As Long
Dim iMaxColDataACH1115 As Integer
Dim rngDataACH1115 As Range

Dim wsDataACH1127 As Worksheet
Dim iMaxRowDataACH1127 As Long
Dim iMaxColDataACH1127 As Integer
Dim rngDataACH1127 As Range


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim PTCacheFinal As PivotCache
Dim PTTableFinal As PivotTable
Dim sNamePivotTable As String

Set wsPT = Worksheets(SheetNamePivotTableGLACH)
wsPT.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iColPT = lRealLastCol + 4

' 1 - Create Pivot table, data from "ACH_1115"
Set wsDataACH1115 = Worksheets(sheetNameDataACH1115)
wsDataACH1115.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataACH1115 = lRealLastRow
iMaxColDataACH1115 = lRealLastCol

Set rngDataACH1115 = Range(Cells(1, 1), Cells(iMaxRowDataACH1115, iMaxColDataACH1115))

'define Pivot Tabel Cache
Set PTCacheFinal = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngDataACH1115)
sNamePivotTable = "WDACH1115"
Set PTTableFinal = PTCacheFinal.CreatePivotTable(TableDestination:=wsPT.Cells(3, iColPT), TableName:=sNamePivotTable)


'add Row (and Columns)


With wsPT.PivotTables(sNamePivotTable).PivotFields("Effective Date")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .AutoSort xlAscending, "Effective Date"
    .Position = 1
End With


'Add data into PivotTable
With wsPT.PivotTables(sNamePivotTable).PivotFields("Debit Amount")
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
    .PivotFields("Effective Date").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With

wsPT.Cells(1, iColPT) = "ACH_1115"




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


' 2 - Create Pivot table, data from "ACH_1127"

irowpt = Range(Cells(65536, iColPT), Cells(65536, iColPT)).End(xlUp).Row
irowpt = irowpt + 6

Set wsDataACH1127 = Worksheets(sheetNameDataACH1127)
wsDataACH1127.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataACH1127 = lRealLastRow
iMaxColDataACH1127 = lRealLastCol

Set rngDataACH1127 = Range(Cells(1, 1), Cells(iMaxRowDataACH1127, iMaxColDataACH1127))

'define Pivot Tabel Cache
Set PTCacheFinal = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngDataACH1127)
sNamePivotTable = "WDACH1127"
Set PTTableFinal = PTCacheFinal.CreatePivotTable(TableDestination:=wsPT.Cells(irowpt, iColPT), TableName:=sNamePivotTable)


'add Row (and Columns)


With wsPT.PivotTables(sNamePivotTable).PivotFields("As of Date")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .AutoSort xlAscending, "As of Date"
    .Position = 1
End With

With wsPT.PivotTables(sNamePivotTable).PivotFields("Return Type Desc")
    .Orientation = xlPageField
    '.Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 1
End With

'Add data into PivotTable
With wsPT.PivotTables(sNamePivotTable).PivotFields("Debit Amount")
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
    .PivotFields("Return Type Desc").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With


'--- apply filter to show "Return" only
'This works for typical PivotTables where blank items appear as "(blank)"
With wsPT.PivotTables(sNamePivotTable).PivotFields("Return Type Desc")
    .ClearAllFilters
        
    'Try the simplest approach first (works in most cases)
    On Error Resume Next
    .CurrentPage = "Return"
    On Error GoTo 0
        
    'If CurrentPage fails (some pivot setups), fallback to manual visibility
    If .CurrentPage <> "Return" Then
        Dim pi As PivotItem
        On Error Resume Next
        For Each pi In .PivotItems
            pi.Visible = (pi.Name = "(blank)")
        Next pi
        On Error GoTo 0
    End If
End With




wsPT.Cells(irowpt - 4, iColPT) = "ACH_1127"




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



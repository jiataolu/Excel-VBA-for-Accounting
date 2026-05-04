Attribute VB_Name = "Module66_delete_Pivot"
Option Explicit

Sub Generate_Pivot_Table_8000_10115()

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
Cells.Select
Selection.Delete
Cells(1, 1).Select


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
Set PTTableFinal = PTCacheFinal.CreatePivotTable(TableDestination:=wsPT.Cells(1, 1), TableName:=sNamePivotTable)


'add Row (and Columns)

With wsPT.PivotTables(sNamePivotTable).PivotFields("Posting Date")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 1
End With

With wsPT.PivotTables(sNamePivotTable).PivotFields("Document Type")
    '.Orientation = xlRowField
    .Orientation = xlColumnField
    .Position = 1
End With

'add filter
With wsPT.PivotTables(sNamePivotTable).PivotFields("Return Yes / No")
    .Orientation = xlPageField
    '.Orientation = xlColumnField
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
'With wsPT.PivotTables(sNamePivotTable)
'    .PivotFields("Account").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'    .PivotFields("Concentration-BU").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'    .PivotFields("Concentration-GL").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'    .PivotFields("Offset Kyriba Code").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'    .PivotFields("Offset-BU").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'    .PivotFields("Offset-GL").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'End With


'turn off "Grand Total"
'With wsPT.PivotTables(sNamePivotTable)
'    .RowGrand = False      ' Removes Grand Total at the bottom
'    .ColumnGrand = False   ' Removes Grand Total on the right
'End With

'=== Apply filter to show only (blank) in "If Return" ===
With wsPT.PivotTables(sNamePivotTable).PivotFields("Return Yes / No")
    .ClearAllFilters
        
    On Error Resume Next   ' In case "(blank)" doesn't exist yet
        .PivotItems("(blank)").Visible = True
        ' Hide everything else
        Dim pi As PivotItem
        For Each pi In .PivotItems
            If pi.Name <> "(blank)" Then
                pi.Visible = False
            End If
        Next pi
    On Error GoTo 0
End With





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



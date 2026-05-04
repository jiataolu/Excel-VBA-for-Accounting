Attribute VB_Name = "Module100_Output_Report"
Option Explicit

Sub Mapping_100_Output_Mapping_Report()

Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
    
Dim strYearNumber As String
Dim strMonthNumber As String
Dim strMonthLetter As String
Dim strTodayDate As String

'Dim strFileNamexlsm As String
'Dim strFullFileNamexlsm As String
Dim strFileNamexlsx As String
Dim strFullFileNamexlsx As String

Dim wkbNewxlsxFile As Workbook
Dim iCountSheetNew As Integer
Dim strSheetNameTemp As String
Dim i As Integer
Dim wsNew As Worksheet
Dim wsCurrent As Worksheet
Dim iMaxRowNewSheet As Integer



Dim wkbMacro As Workbook
Dim wsMacro As Worksheet
Dim strSheetNameMacro As String
Dim iCountSheetMacro As Integer

Dim rngCopy As Range
Dim rngPaste As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long


strTodayDate = Date

strYearNumber = CStr(Format(strTodayDate, "yyyy"))
'Debug.Print strYearNumber

strMonthNumber = CStr(Format(strTodayDate, "MM"))
'Debug.Print strMonthNumber

strMonthLetter = Format(strTodayDate, "mmm")
'Debug.Print strMonthLetter

'strFileNamexlsm = strYearNumber & "_" & strMonthNumber & " Cash Position Report " & strMonthLetter & " " & strYearNumber & ".xlsm"
'Debug.Print strFileNamexlsm

'strFullFileNamexlsm = GetWorkPath & "\" & strFileNamexlsm
'Debug.Print strFullFileNamexlsm

'ThisWorkbook.SaveCopyAs strFullFileNamexlsm


strFileNamexlsx = strYearNumber & "_" & strMonthNumber & " Mapping " & strMonthLetter & " " & strYearNumber & ".xlsx"
'Debug.Print strFileNamexlsx

strFullFileNamexlsx = GetWorkPath & "\" & strFileNamexlsx
'Debug.Print strFullFileNamexlsx

strSheetNameTemp = "$$$Temp"
Set wkbNewxlsxFile = Workbooks.Add
wkbNewxlsxFile.Activate

iCountSheetNew = Worksheets.Count
Debug.Print iCountSheetNew
For i = 1 To iCountSheetNew
    Worksheets(i).Name = strSheetNameTemp & CStr(i)
Next i

Set wkbMacro = ThisWorkbook
wkbMacro.Activate
iCountSheetMacro = Worksheets.Count

For i = 1 To iCountSheetMacro
    wkbMacro.Activate
    Worksheets(i).Activate
    strSheetNameMacro = Worksheets(i).Name
    
    If strSheetNameMacro = SheetNameMapping Or strSheetNameMacro = SheetNameDeleted Or strSheetNameMacro = SheetNameFIS Then
        Cells.Select
        Set rngCopy = Selection
        'Selection.Interior.Color = xlNone
    
        wkbNewxlsxFile.Activate
        Set wsCurrent = ActiveSheet
        Set wsNew = Worksheets.Add(after:=wsCurrent)
        wsNew.Name = strSheetNameMacro
        wsNew.Select
        Set rngPaste = Cells(1, 1)
    
        rngCopy.Copy Destination:=rngPaste
        Cells(1, 1).Select
    End If
Next i

wkbNewxlsxFile.Activate

For i = 1 To iCountSheetNew
    Worksheets(strSheetNameTemp & CStr(i)).Delete
Next i

Worksheets("Mapping Consolidated").Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowNewSheet = lRealLastRow
If iMaxRowNewSheet > 1 Then Rows("2:" & iMaxRowNewSheet).Interior.Color = xlNone
Call DeleteUnusedFormats

Cells(1, 1).Select
Selection.AutoFilter


' ********** Pivot table **************************

Dim wsData As Worksheet
Dim wsPT As Worksheet
'to question PCFinal variable
'Dim PCFinal As PivotCache
'Dim PTFin As PivotTable

Dim PTCacheFinal As PivotCache
Dim PTTableFinal As PivotTable
Dim sNamePivotTable As String

Dim rngData As Range
Dim lLastRowData As Long
Dim lLastColData As Long

Dim sNameDataSheet As String
Dim sNamePTSheet As String
'Dim lRealLastRow As Long
'Dim lRealLastCol As Long

sNameDataSheet = "Mapping Consolidated"
sNamePTSheet = "Pivot Mapping"

'Define working sheets
'Dim i As Integer
'Application.DisplayAlerts = False
For i = 1 To Worksheets.Count
    If Worksheets(i).Name = sNamePTSheet Then
        Worksheets(sNamePTSheet).Delete
        Exit For
    End If
Next i
Worksheets(sNameDataSheet).Select
Worksheets.Add after:=ActiveSheet
ActiveSheet.Name = sNamePTSheet
'Application.DisplayAlerts = True

Set wsPT = Worksheets(sNamePTSheet)
Set wsData = Worksheets(sNameDataSheet)

'define working range"
wsData.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lLastRowData = lRealLastRow
lLastColData = lRealLastCol
Set rngData = Range(Cells(1, 1), Cells(lLastRowData, lLastColData))
wsPT.Select

'define Pivot Tabel Cache
'Set PTFin = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData).CreatePivotTable()
Set PTCacheFinal = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData)
sNamePivotTable = "PivotTable01"
Set PTTableFinal = PTCacheFinal.CreatePivotTable(Tabledestination:=wsPT.Cells(1, 1), TableName:=sNamePivotTable)

'add Row (and ColumnAccount Name
With wsPT.PivotTables(sNamePivotTable).PivotFields("FIS Code")
    .Orientation = xlRowField
    '.Orientation = xlColumnField
    .Position = 1
End With

'With wsPT.PivotTables(sNamePivotTable).PivotFields("Month")
'    .Orientation = xlRowField
'    '.Orientation = xlColumnField
'    .Position = 2
'End With

'Add data into PivotTable
'With wsPT.PivotTables(sNamePivotTable).PivotFields("Amount to Invoice in LC")
'    .Orientation = xlDataField
'    .Function = xlSum
'    .NumberFormat = "$#,##0.00"
'    .Name = "Amount"
'End With

'Define the format of PivotTable
wsPT.PivotTables(sNamePivotTable).ShowTableStyleRowStripes = True
wsPT.PivotTables(sNamePivotTable).TableStyle = "PivotStyleMedium9"





'  *************** Pivot Table


















Worksheets("Mapping Consolidated").Select
Range("A1").Select
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True

Worksheets("Mapping Consolidated").Protect Password:="banking", Contents:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

Worksheets("Deleted").Select
Cells(1, 1).Select
Selection.AutoFilter

Range("A1").Select
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True
Worksheets("Deleted").Protect Password:="banking", Contents:=True, AllowFiltering:=True, AllowUsingPivotTables:=True


Worksheets("Mapping Consolidated").Select
wkbNewxlsxFile.SaveCopyAs strFullFileNamexlsx
wkbNewxlsxFile.Close savechanges:=False

ThisWorkbook.Activate
Worksheets("Mapping Consolidated").Select
Cells(1, 1).Select

' Re-enable macro execution
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

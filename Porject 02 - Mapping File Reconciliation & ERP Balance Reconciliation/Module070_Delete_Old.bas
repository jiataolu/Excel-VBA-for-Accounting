Attribute VB_Name = "Module070_Delete_Old"
Option Explicit
Sub Mapping_070_Delete_Lines()

Dim wsDelete As Worksheet
Dim iMaxRowDelete As Integer
Dim iCurrentRowDelete As Integer

Dim wsMap As Worksheet
Dim iMaxRowMap As Integer
Dim iRowMap As Integer
Dim sMapRemark As String
Dim sMapBankAcctFull As String

Dim wsFIS As Worksheet


Dim rngCopy As Range
Dim rngPaste As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsDelete = Worksheets(SheetNameDeleted)
wsDelete.Select
Call DeleteUnusedFormats

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDelete = lRealLastRow
iCurrentRowDelete = iMaxRowDelete

Set wsMap = Worksheets(SheetNameMapping)
wsMap.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMap = lRealLastRow

For iRowMap = 2 To iMaxRowMap
    wsMap.Select
    sMapRemark = wsMap.Cells(iRowMap, ColMapRemark)
    sMapBankAcctFull = wsMap.Cells(iRowMap, ColMapBankAcctFull)
    If Replace(sMapRemark, " ", "") = "" And sMapBankAcctFull <> "752-82605" Then
        iCurrentRowDelete = iCurrentRowDelete + 1
        
        Set rngCopy = Range(Cells(iRowMap, 1), Cells(iRowMap, ColMapComment))
        'wsDelete.Select
        Set rngPaste = wsDelete.Cells(iCurrentRowDelete, 1)
        rngCopy.Copy Destination:=rngPaste
        
        wsDelete.Cells(iCurrentRowDelete, ColDeletedDeletedData) = "Deleted at " & Format(Date, "MMM DD, YYYY")
        
        Set rngCopy = Nothing
        Set rngPaste = Nothing
    End If

Next iRowMap

'Count how many lines are new, and how many lines are deleted,
LineDeleted = 0
LineNew = 0
For iRowMap = iMaxRowMap To 2 Step -1
    wsMap.Select
    sMapRemark = wsMap.Cells(iRowMap, ColMapRemark)
    sMapBankAcctFull = wsMap.Cells(iRowMap, ColMapBankAcctFull)
    If UCase(Replace(sMapRemark, " ", "")) = "NEW" Then LineNew = LineNew + 1
    If Replace(sMapRemark, " ", "") = "" And sMapBankAcctFull <> "752-82605" Then
        LineDeleted = LineDeleted + 1
        Rows(iRowMap).Delete
    End If
Next iRowMap

Call DeleteUnusedFormats
wsMap.Select
Cells(1, 1).Select

Debug.Print LineDeleted
Debug.Print LineNew

MsgBox "Deleted acounts - " & CStr(LineDeleted)
MsgBox "New Added Accounts - " & CStr(LineNew)


End Sub




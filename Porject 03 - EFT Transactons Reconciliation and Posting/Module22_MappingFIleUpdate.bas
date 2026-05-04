Attribute VB_Name = "Module22_MappingFIleUpdate"
Option Explicit

Sub Mapping_File_Update()
Application.DisplayAlerts = False

Dim wkbMap As Workbook
Dim wsMapEFT As Worksheet
Dim iRowMapEFT As Integer
Dim iMaxRowMapEFT As Integer
Dim sBankAcct As String


Dim iSheet As Integer
Dim iFoundTempSheet As Integer
Dim sSheetNameMapEFT As String

Dim wsMap As Worksheet

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim rngCopy As Range
Dim rngPaste As Range

sSheetNameMapEFT = "Mapping EFT"

Set wkbMap = Workbooks.Open(Map_File_Full_Name())
wkbMap.Activate

'Check if new tab- "Mapping EFT" exist. If exist, then delete and re-create.
'If not exist, then create it.
iFoundTempSheet = 0
For iSheet = 1 To Worksheets.Count()
    If Worksheets(iSheet).Name = sSheetNameMapEFT Then
        iFoundTempSheet = 1
        Exit For
    End If
Next iSheet

If iFoundTempSheet = 1 Then Worksheets(sSheetNameMapEFT).Delete
Set wsMapEFT = Worksheets.Add
wsMapEFT.Name = sSheetNameMapEFT


'Copy data from original mapping sheet to new Mapping EFT sheet
wsMapEFT.Select
Set rngPaste = Cells(1, 1)

Set wsMap = Worksheets("Mapping Consolidated")
wsMap.Select
wsMap.Unprotect Password:="banking"
wsMap.AutoFilterMode = False

Cells.Select
Set rngCopy = Selection
rngCopy.Copy Destination:=rngPaste


'Work on Mapping EFT sheet, to delete all leading zero of account number
wsMapEFT.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMapEFT = lRealLastRow
'Debug.Print iMaxRowMapEFT
If iMaxRowMapEFT < 2 Then Exit Sub

For iRowMapEFT = 2 To iMaxRowMapEFT
    sBankAcct = wsMapEFT.Cells(iRowMapEFT, iColMapBankAcct)
    sBankAcct = CStr(Number_wt_Leading_Zero(sBankAcct))
    wsMapEFT.Cells(iRowMapEFT, iColMapBankAcct) = sBankAcct
Next iRowMapEFT


Columns(iColMapBankAcct).Select
Selection.NumberFormat = "0"
Selection.HorizontalAlignment = xlLeft
Cells(1, 1).Select

'Set wsMapEFT = Nothing

wkbMap.Close SaveChanges:=True
Application.DisplayAlerts = True

    
End Sub



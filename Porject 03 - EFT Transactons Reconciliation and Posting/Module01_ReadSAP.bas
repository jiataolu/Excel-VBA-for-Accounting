Attribute VB_Name = "Module01_ReadSAP"
Option Explicit

Sub Read_SAP_File()

Dim strSAPFileFullName As String
Dim wkbSAP As Workbook
Dim wsSAPCopy As Worksheet

Dim wsSAPPaste As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim sClearNote As String

Dim rngCopy As Range
Dim rngPaste As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAPPaste = Worksheets("1-SAP")
wsSAPPaste.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select
Set rngPaste = Cells(1, 1)

strSAPFileFullName = GetWorkPath() & "\" & SubFolderInput & "\" & FileNameSAP
Set wkbSAP = Workbooks.Open(strSAPFileFullName)
wkbSAP.Activate
Set wsSAPCopy = Worksheets(1)
Cells.Select
Set rngCopy = Selection
Selection.Copy Destination:=rngPaste

wkbSAP.Close SaveChanges:=False

wsSAPPaste.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow
If iMaxRowSAP < 2 Then Exit Sub

For iRowSAP = iMaxRowSAP To 2 Step -1
    sClearNote = wsSAPPaste.Cells(iRowSAP, iColSAPClear)
    If Replace(sClearNote, " ", "") <> "" Then Rows(iRowSAP).Delete
Next iRowSAP

Call DeleteUnusedFormats

Cells(1, 1).Select
End Sub

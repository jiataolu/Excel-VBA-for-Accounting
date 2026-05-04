Attribute VB_Name = "Module02_ReadPAPBankStatement"
Option Explicit

Sub Read_Bank_Statement()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

Dim strFileFullBankStatementPAP As String
Dim wkbBankStatementPAP As Workbook
Dim wsBankstatementPAPCopy As Worksheet
Dim iMaxRowBankStatementPAPCopy As Integer
Dim iMaxColBankStatementPAPCopy As Integer

Dim wkbMacro As Workbook
Dim wsBankstatementPAPPaste As Worksheet

Dim rngCopy As Range
Dim rngPaste As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim strFileFullReconciledPAPBankStatement As String
Set wkbMacro = ThisWorkbook
Set wsBankstatementPAPPaste = Worksheets("Bank Statement")
wsBankstatementPAPPaste.Select
Cells.Select
Selection.Delete
Set rngPaste = Cells(1, 1)


strFileFullBankStatementPAP = GetWorkPath & "\" & SubFolder & "\" & FileBankStatementPAP
'Debug.Print strFileFullBankStatementPAP

Set wkbBankStatementPAP = Workbooks.Open(strFileFullBankStatementPAP)
wkbBankStatementPAP.Activate
Set wsBankstatementPAPCopy = Worksheets(1)
wsBankstatementPAPCopy.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBankStatementPAPCopy = lRealLastRow
iMaxColBankStatementPAPCopy = lRealLastCol

Set rngCopy = Range(Cells(1, 1), Cells(iMaxRowBankStatementPAPCopy, iMaxColBankStatementPAPCopy))

rngCopy.Copy Destination:=rngPaste

wkbBankStatementPAP.Activate

strFileFullReconciledPAPBankStatement = GetWorkPath & "\" & SubFolderOutput & "\" & FileReconPAPBankStatement
wkbBankStatementPAP.SaveAs strFileFullReconciledPAPBankStatement, xlOpenXMLWorkbook
wkbBankStatementPAP.Close SaveChanges:=False
Cells(1, 1).Select
Cells.Select
Cells.EntireColumn.AutoFit
Cells(1, 1).Select

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub


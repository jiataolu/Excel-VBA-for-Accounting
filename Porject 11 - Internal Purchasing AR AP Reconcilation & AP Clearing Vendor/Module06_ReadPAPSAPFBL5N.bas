Attribute VB_Name = "Module06_ReadPAPSAPFBL5N"
Option Explicit

Sub Read_SAP_FBL5N(CompanyName As String)

'Dim CompanyName As String
'CompanyName = "MSD"

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

Dim strFileFullSAP As String
Dim wkbSAP As Workbook
Dim wsSAPCopy As Worksheet
Dim iMaxRowSAPCopy As Integer
Dim iMaxColSAPCopy As Integer

Dim wkbMacro As Workbook
Dim wsSAPPaste As Worksheet
'Dim iMaxRowSAPPaste As Integer
'Dim iSAPPaste As Integer

Dim rngCopy As Range
Dim rngPaste As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

'Dim strFileFullReconciledPAPBankStatement As String
Set wkbMacro = ThisWorkbook
Set wsSAPPaste = Worksheets("FBL5N")
wsSAPPaste.Select
Cells.Select
Selection.Delete
Set rngPaste = Cells(1, 1)


strFileFullSAP = GetWorkPath & "\" & SubFolder & "\" & "SAP-" & CompanyName & ".xlsx"
'Debug.Print strFileFullBankStatementPAP

    Set wkbSAP = Workbooks.Open(strFileFullSAP)
wkbSAP.Activate
Set wsSAPCopy = Worksheets("Sheet1")
wsSAPCopy.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAPCopy = lRealLastRow
iMaxColSAPCopy = lRealLastCol

Set rngCopy = Range(Cells(1, 1), Cells(iMaxRowSAPCopy, iMaxColSAPCopy))

rngCopy.Copy Destination:=rngPaste

wkbSAP.Activate


wkbSAP.Close SaveChanges:=False
Cells(1, 1).Select



Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.EnableEvents = True



End Sub

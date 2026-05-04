Attribute VB_Name = "Module14_FinalBankRecon"
Option Explicit

Sub Final_Bank_Reconciliation()

Dim wkbFinalBankRecon As Workbook
Dim wsBankTo As Worksheet
Dim strFileFullFinalBankRecon As String
Dim rngPaste As Range

Dim wkbMacro As Workbook
Dim wsBankFrom As Worksheet
Dim iMaxRowBankFrom As Integer
Dim rngCopy As Range

Dim wkbReport As Workbook
Dim strFileFullReport As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wkbMacro = ThisWorkbook

strFileFullFinalBankRecon = GetWorkPath & "\" & SubFolderOutput & "\" & FileReconPAPBankStatement
Set wkbFinalBankRecon = Workbooks.Open(strFileFullFinalBankRecon)
wkbFinalBankRecon.Activate
Set wsBankTo = Worksheets(1)
wsBankTo.Select
wsBankTo.Cells(1, ColBSEntity) = "Entity"
wsBankTo.Cells(1, ColBSAMTPAP) = "Amount PAP"
wsBankTo.Cells(1, ColBSTradingPart) = "Trading Partner"
wsBankTo.Cells(1, ColBSCustomer) = "Customer ID"
wsBankTo.Cells(1, ColBSBranch) = "Branch"

Set rngPaste = wsBankTo.Cells(2, ColBSEntity)

wkbMacro.Activate
Set wsBankFrom = Worksheets("Bank Statement")
wsBankFrom.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBankFrom = lRealLastRow
Set rngCopy = Range(Cells(2, ColBSEntity), Cells(iMaxRowBankFrom, ColBSEntity))

rngCopy.Copy Destination:=rngPaste

wkbFinalBankRecon.Activate
wsBankTo.Select
'Range(Cells(2, ColBSEntity), Cells(iMaxRowBankFrom, ColBSEntity)).Font.Color = RGB(0, 0, 0)

strFileFullReport = GetWorkPath & "\" & SubFolderOutput & "\" & "MSD PAP clearing.xlsx"
Set wkbReport = Workbooks.Open(strFileFullReport)
wkbReport.Activate
Set wsBankFrom = Worksheets("Bank Statement")

Call One_Bank_starement_PAP_Info_Copy_Paste(wkbReport, wsBankFrom, wkbFinalBankRecon, wsBankTo)

wkbReport.Close SaveChanges:=False


strFileFullReport = GetWorkPath & "\" & SubFolderOutput & "\" & "SPS PAP clearing.xlsx"
Set wkbReport = Workbooks.Open(strFileFullReport)
wkbReport.Activate
Set wsBankFrom = Worksheets("Bank Statement")

Call One_Bank_starement_PAP_Info_Copy_Paste(wkbReport, wsBankFrom, wkbFinalBankRecon, wsBankTo)

wkbReport.Close SaveChanges:=False

strFileFullReport = GetWorkPath & "\" & SubFolderOutput & "\" & "Well.ca PAP clearing.xlsx"
Set wkbReport = Workbooks.Open(strFileFullReport)
wkbReport.Activate
Set wsBankFrom = Worksheets("Bank Statement")

Call One_Bank_starement_PAP_Info_Copy_Paste(wkbReport, wsBankFrom, wkbFinalBankRecon, wsBankTo)

wkbReport.Close SaveChanges:=False


wkbFinalBankRecon.Close SaveChanges:=True

End Sub

Sub One_Bank_starement_PAP_Info_Copy_Paste(ReportBSBook As Workbook, ReportBSSheet As Worksheet, ReconBSBook As Workbook, ReconBSSheet As Worksheet)

Dim iMaxRow As Integer
Dim iRow As Integer
Dim jCol As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

ReportBSBook.Activate
ReportBSSheet.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRow = lRealLastRow
'Debug.Print iMaxRow

For iRow = 2 To iMaxRow
'For iRow = 8 To 8
    If Not IsEmpty(ReportBSSheet.Cells(iRow, ColBSAMTPAP)) Then
    'Debug.Print iRow
        For jCol = ColBSAMTPAP To ColBSBranch
            ReconBSSheet.Cells(iRow, jCol) = ReportBSSheet.Cells(iRow, jCol)
        Next jCol
        ReconBSSheet.Rows(iRow).Font.Color = RGB(255, 0, 0)
    End If
Next iRow

End Sub

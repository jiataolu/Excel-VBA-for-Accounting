Attribute VB_Name = "Module04_ProcessDataGLBank"
Option Explicit
Option Base 1

'0_ZBA: GL vs Bank same date (Posting date vs Value date)
'1_Regular ACH: GL posting date mapped to next day in bank statement.
'2_Return: GL posting date mapping to next day in bank statement,
'  -ACH also mapped to next day in bank stataement
'  +ACH does not change the date


Sub Process_GL_Bank_Data()

Dim wsDataGL As Worksheet
Dim iMaxRowDataGL As Long
Dim iRowDataGL As Long
Dim sReference As String
Dim sIfReturn As String
Dim sDocType As String
Dim sTransType As String
Dim dPostingDate As Date
Dim dReconDate As Date
Dim sACHNumber As String
Dim sDocNumber As String

Dim wsDataBank As Worksheet
Dim iMaxRowDataBank As Long
Dim iRowDataBank As Long
Dim sDes As String
Dim sIfRedeposit As String
Dim dDebitAmt As Double
Dim dCreditAmt As Double
Dim sFlowCode As String
Dim sBankType As String
Dim dValueDate As Date
Dim dBankReconDate As Date

Dim wsBankDate As Worksheet
Dim iMaxRowBankDate As Integer
Dim arrBankDate As Variant

Dim lRealLastRow As Long
Dim lRealLastCol As Long



'------------------------------------------------------------------------------------------------
'0 - Bank date list
Set wsBankDate = Worksheets(SheetNameBankDate)
wsBankDate.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBankDate = lRealLastRow
'Debug.Print iMaxRowBankDate

arrBankDate = Range(Cells(1, 1), Cells(iMaxRowBankDate, 1)).Value

'-------------------------------------------------------------------------------------
' 1- process GL Data
Set wsDataGL = Worksheets(SheetNameDataGL)
wsDataGL.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataGL = lRealLastRow
'Debug.Print iMaxRowDataGL

Columns(ColDataGLReturn).Delete
Columns(ColDataGLReturn).Delete
Columns(ColDataGLReturn).Delete
Columns(ColDataGLReturn).Delete
Columns(ColDataGLReturn).Delete
Columns(ColDataGLReturn).Delete

Call DeleteUnusedFormats
Cells(1, ColDataGLReturn) = "Return Yes / No"
Cells(1, ColDataGLTransType) = "Trans_Type"
Cells(1, ColDataGLReconDate) = "Recon_Date"
Cells(1, ColDataGLACHNumber) = "ACH Number"

For iRowDataGL = 2 To iMaxRowDataGL
'For iRowDataGL = 21 To 21
    sIfReturn = ""
    sReference = wsDataGL.Cells(iRowDataGL, ColDataGLReference)
    sDocNumber = wsDataGL.Cells(iRowDataGL, ColDataGLDocNumber)
    If UCase(Right(sReference, 2)) = "RT" Then sIfReturn = "Yes"
    wsDataGL.Cells(iRowDataGL, ColDataGLReturn) = sIfReturn
    
    sDocType = wsDataGL.Cells(iRowDataGL, ColDataGLDocType)
    dPostingDate = wsDataGL.Cells(iRowDataGL, ColDataGLPostingDate)
    
    sTransType = ""
    
    ' ZBA transactions
    If UCase(sDocType) = "9Y" Then
        sTransType = "2_ZBA"
        dReconDate = dPostingDate
    'regular non-return
    ElseIf sIfReturn = "" Then
        sTransType = "0_ACH"
        dReconDate = Find_Next_Date_in_Bank(dPostingDate, arrBankDate)
        'Debug.Print Find_Next_Date_in_Bank(dPostingDate, arrBankDate)
    'Return
    Else
        sTransType = "1_Return"
        dReconDate = Find_Next_Date_in_Bank(dPostingDate, arrBankDate)
    End If
    
    
    wsDataGL.Cells(iRowDataGL, ColDataGLTransType) = sTransType
    wsDataGL.Cells(iRowDataGL, ColDataGLReconDate) = dReconDate
    
    sACHNumber = ""
    'process ACH number
    If sIfReturn = "" And sDocType = "DZ" Then sACHNumber = sDocNumber
    If sIfReturn = "" And sDocType = "Z8" Then sACHNumber = "ACH" & Right(sReference, 8)

    wsDataGL.Cells(iRowDataGL, ColDataGLACHNumber) = sACHNumber
Next iRowDataGL

'-------------------------------------------------------------------------------------
' 2 - process Bank Data
Set wsDataBank = Worksheets(SheetNameDataBank)
wsDataBank.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataBank = lRealLastRow
'Debug.Print iMaxRowDataBank

Columns(ColDataBankAmount).Delete
Columns(ColDataBankAmount).Delete
Columns(ColDataBankAmount).Delete
Columns(ColDataBankAmount).Delete
Call DeleteUnusedFormats

Cells(1, ColDataBankAmount) = "Amount"
Cells(1, ColDataBankRedeposit) = "REDEPOSIT YES/ NO"
Cells(1, ColDataBankType) = "Trans_Type"
Cells(1, ColDataBankReconDate) = "Recon_Date"

For iRowDataBank = 2 To iMaxRowDataBank

    dCreditAmt = wsDataBank.Cells(iRowDataBank, ColDataBankCreditAmt)
    dDebitAmt = wsDataBank.Cells(iRowDataBank, ColDataBankDebitAmt)
    wsDataBank.Cells(iRowDataBank, ColDataBankAmount) = dCreditAmt - dDebitAmt
    
    sIfRedeposit = ""
    sDes = wsDataBank.Cells(iRowDataBank, ColDataBankDescription)
    If InStr(sDes, "REDEPOSITS") > 0 Then sIfRedeposit = "Yes"
    wsDataBank.Cells(iRowDataBank, ColDataBankRedeposit) = sIfRedeposit
    
    sFlowCode = wsDataBank.Cells(iRowDataBank, ColDataBankFlowCode)
    dValueDate = wsDataBank.Cells(iRowDataBank, ColDataBankValueDate)
    
    ' ZBA
    If InStr(sFlowCode, "ZBA") > 0 Then
        sBankType = "2_ZBA"
        dBankReconDate = dValueDate
    ' Regular ACH deposit
    ElseIf sFlowCode = "+ACH" And sIfRedeposit = "" Then
        sBankType = "0_ACH"
        dBankReconDate = dValueDate
    'Return
    ElseIf sFlowCode = "+ACH" And sIfRedeposit = "Yes" Then
        sBankType = "1_Return"
        dBankReconDate = dValueDate
    'Return
    ElseIf sFlowCode = "-ACH" Then
        sBankType = "1_Return"
        dBankReconDate = Find_Next_Date_in_Bank(dValueDate, arrBankDate)
    Else
        sBankType = "9_Other"
        dBankReconDate = dValueDate
    End If

    wsDataBank.Cells(iRowDataBank, ColDataBankType) = sBankType
    wsDataBank.Cells(iRowDataBank, ColDataBankReconDate) = dBankReconDate

Next iRowDataBank




End Sub

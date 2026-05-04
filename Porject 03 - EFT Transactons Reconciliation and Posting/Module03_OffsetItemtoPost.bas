Attribute VB_Name = "Module03_OffsetItemtoPost"
Option Explicit

'In Kyriba bank statement, add net amount
Sub Process_Kyriba_Bank_Statement()

Dim strBSFileFullName As String
Dim wkbBS As Workbook
Dim wsBS As Worksheet
Dim iMaxRowBS As Integer
Dim iRowBS As Integer
Dim sAmountDeposit As String
Dim sAmountPayment As String
Dim dAmountDeposit As Double
Dim dAmountPayment As Double
Dim sNetAmount As String
Dim dNetAmount As Double

Dim lRealLastRow As Long
Dim lRealLastCol As Long

strBSFileFullName = Bank_Statement_File_Full_Name
'Debug.Print strBSFileFullName

If strBSFileFullName = "" Then Exit Sub

Set wkbBS = Workbooks.Open(strBSFileFullName)
wkbBS.Activate
Set wsBS = Worksheets(1)
wsBS.Select
'wsBS.Columns(iColBSInsertNetAmt).EntireColumn.Insert
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBS = lRealLastRow
If iMaxRowBS < 2 Then Exit Sub


For iRowBS = iMaxRowBS To 2 Step -1
    If wsBS.Cells(iRowBS, 1) = "Account" And wsBS.Cells(iRowBS, 2) = "Bank" And wsBS.Cells(iRowBS, 3) = "Account cur." Then
        'Debug.Print iRowBS
        Rows(iRowBS).Delete
    End If
Next iRowBS
Call DeleteUnusedFormats


wsBS.Select
'wsBS.Columns(iColBSInsertNetAmt).EntireColumn.Insert
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBS = lRealLastRow
If iMaxRowBS < 2 Then Exit Sub


wsBS.Cells(1, iColBSAMT) = "Net Amount"
'For iRowBS = 2 To 2
For iRowBS = 2 To iMaxRowBS
    'Debug.Print "Row-" & iRowBS
    sAmountDeposit = Replace(wsBS.Cells(iRowBS, iColBSDepositAMT), " ", "")
    sAmountPayment = Replace(wsBS.Cells(iRowBS, iColBSPaymentAMT), " ", "")
    
    If sAmountDeposit = "" Then
        dAmountDeposit = 0
    Else
        dAmountDeposit = CDbl(sAmountDeposit)
    End If
    
    If sAmountPayment = "" Then
        dAmountPayment = 0
    Else
        dAmountPayment = CDbl(sAmountPayment)
    End If
    
    dNetAmount = dAmountDeposit - dAmountPayment
    wsBS.Cells(iRowBS, iColBSAMT) = dNetAmount

    wsBS.Cells(iRowBS, iColBSAMT).Style = "Currency"
    wsBS.Cells(iRowBS, iColBSDepositAMT).Style = "Currency"
    wsBS.Cells(iRowBS, iColBSPaymentAMT).Style = "Currency"
Next iRowBS


wkbBS.Close SaveChanges:=True

End Sub


Sub Activate_Offset_Items_to_Read()

'keyword in bank statement: "ATHENS" + "ID: 001233113647"

Dim strBSFileFullName As String
Dim wkbBS As Workbook
Dim wsBS As Worksheet
Dim iMaxRowBS As Integer
Dim iRowBS As Integer
Dim strBSBankCode As String
Dim dBSAMT As Double
Dim strBSComment As String


Dim wsSAP As Worksheet
Dim iRowSAP As Integer
Dim iMaxRowSAP As Integer
Dim dSAPAmt As Double
Dim sSAPGL As String


Dim wsConClear As Worksheet
Dim iRowConClear As Integer
Dim iMaxRowConClear As Integer
Dim sGL As String
Dim sConClearBankCode As String

'Dim wsSAP As Worksheet
'Dim iMaxRowSAP As Integer
'Dim iRowSAP As Integer

Dim iFoundATHENS As Integer
Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsConClear = Worksheets("Concentration & Clearing GL")
wsConClear.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowConClear = lRealLastRow


Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow


strBSFileFullName = Bank_Statement_File_Full_Name
If strBSFileFullName = "" Then Exit Sub

Set wkbBS = Workbooks.Open(strBSFileFullName)
wkbBS.Activate
Set wsBS = Worksheets(1)
wsBS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBS = lRealLastRow
If iMaxRowBS < 2 Then
    MsgBox "Please check bank statement"
    wkbBS.Close SaveChanges:=False
    Exit Sub
End If

iFoundATHENS = 0
For iRowBS = 2 To iMaxRowBS
    wkbBS.Activate
    wsBS.Select
    strBSComment = wsBS.Cells(iRowBS, iColBSComment)
    strBSComment = UCase(Replace(strBSComment, " ", ""))
    
    sGL = ""
    If InStr(strBSComment, "ATHENS") > 0 And InStr(strBSComment, "ID:001233113647") > 0 Then
        
        iFoundATHENS = 1
        dBSAMT = wsBS.Cells(iRowBS, iColBSAMT)
        strBSBankCode = wsBS.Cells(iRowBS, iColBSBankCode)
        'Debug.Print "In Line:" & iRowBS
        'Debug.Print "In Line:" & dBSAMT
        'Debug.Print "In Line:" & strBSBankCode
        
        sGL = ""
        For iRowConClear = 2 To iMaxRowConClear
            sConClearBankCode = wsConClear.Cells(iRowConClear, iColConBankCode)
            If InStr(sConClearBankCode, strBSBankCode) > 0 Then
                sGL = wsConClear.Cells(iRowConClear, iColConcenClear)
                Exit For
            End If
        Next iRowConClear
    End If
    
    If sGL = "" Then GoTo CONTINUEBANKSTATEMENT
    
    'Debug.Print sGL
    'Debug.Print dBSAMT
    
    For iRowSAP = 2 To iMaxRowSAP
    'For iRowSAP = 14 To 14
        sSAPGL = wsSAP.Cells(iRowSAP, iColSAPGL)
        dSAPAmt = wsSAP.Cells(iRowSAP, iColSAPAMT)
        'Debug.Print sSAPGL
        'Debug.Print dSAPAMT
        
        'If absolute amount from SAP is equal to "ATHENS" amount, and GL is same
        If sSAPGL = sGL And Abs(dSAPAmt) = Abs(dBSAMT) Then
            ThisWorkbook.Activate
            wsSAP.Select
            Debug.Print "Found"
            'to check if it is already highlithed
            wsSAP.Cells(iRowSAP, iColSAPClear) = ""
            wsSAP.Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.Pattern = xlNone
        
        End If
    Next iRowSAP
    
    
CONTINUEBANKSTATEMENT:
Next iRowBS


wkbBS.Close SaveChanges:=False

If iFoundATHENS = 1 Then MsgBox "There is bank transaction for ""ATHENS ID:001233113647""."
End Sub


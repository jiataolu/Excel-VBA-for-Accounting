Attribute VB_Name = "Module040_Consolidate_FIS_PS"
Option Explicit

Sub Mapping_040_Consolidate_FIS_PS()

Dim wsFIS As Worksheet
Dim iMaxRowFIS As Integer
Dim iRowFIS As Integer
Dim sBankAcctFull As String
Dim sBankAcctKey As String
Dim sFISFISCode As String
Dim iLen As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsFIS = Worksheets(SheetNameFIS)
wsFIS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
If iMaxRowFIS < 2 Then Exit Sub

For iRowFIS = 2 To iMaxRowFIS
'For iRowFIS = 438 To 450
    
    'When macro read bank account (long), single quotation is gone.
    sFISFISCode = wsFIS.Cells(iRowFIS, ColFISFISCode)

    sBankAcctFull = wsFIS.Cells(iRowFIS, ColFISBankAcct)
    'Debug.Print sBankAcctFull
    
    'All accounts compare with bank account number, the last 9 digits or defind by variable LenKeyBankAcctNo
    'except, Deal with 3 lines whose bank account is only "x". Add FIS bank code
    If UCase(Replace(sBankAcctFull, " ", "")) = "X" Then sBankAcctFull = Replace(sBankAcctFull & sFISFISCode, " ", "")
        
    iLen = Len(sBankAcctFull)
    

    'Internal key code: Key-9 digits
    If iLen < LenKeyBankAcctNo Then
        sBankAcctKey = "Key-" & sBankAcctFull
    Else
        sBankAcctKey = "Key-" & Right(sBankAcctFull, LenKeyBankAcctNo)
    End If
    
    wsFIS.Cells(iRowFIS, ColFISKeyNumber) = sBankAcctKey
Next iRowFIS


'remove empty line
'For iRowFIS = iMaxRowFIS To 2 Step -1
For iRowFIS = 445 To 445
    sBankAcctFull = wsFIS.Cells(iRowFIS, ColFISBankAcct)
    sBankAcctFull = Replace(sBankAcctFull, "'", "")
    sBankAcctFull = Replace(sBankAcctFull, " ", "")
    
    If sBankAcctFull = "" Then Rows(iRowFIS).Delete
    
Next iRowFIS

Call DeleteUnusedFormats

End Sub



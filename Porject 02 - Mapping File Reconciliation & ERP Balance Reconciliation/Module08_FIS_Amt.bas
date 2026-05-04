Attribute VB_Name = "Module08_FIS_Amt"
Option Explicit

Sub Read_FIS_Data_into_Cash_Porject()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsFIS As Worksheet
Dim iStartRowFIS As Integer
Dim iMaxRowFIS As Integer
Dim iFIS As Integer
Dim strBankCode As String

Dim wsCP As Worksheet
Dim iMaxRowCP As Integer
Dim iCP As Integer
Dim strCPBankCode As String
Dim strCPFISAmt As String
Dim strNumber As String

Dim wsBankVar As Worksheet
Dim rngBankVar As Range
Dim rngFound As Range

Set wsCP = Worksheets("Cash Project")
wsCP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCP = lRealLastRow
'Debug.Print iMaxRowCP

Set wsFIS = Worksheets("FIS")
wsFIS.Select
Call DeleteUnusedFormats
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
If iMaxRowFIS < 2 Then Exit Sub
'Debug.Print iMaxRowFIS

Columns(iFISCheck).Select
Selection.ClearContents


'To see if last row has "Total"
strBankCode = wsFIS.Cells(iMaxRowFIS, iFISBankCode)
strBankCode = Replace(strBankCode, " ", "")
strBankCode = UCase(strBankCode)
If strBankCode = "TOTAL" Then iMaxRowFIS = iMaxRowFIS - 1
'Debug.Print iMaxRowFIS

'To see if the second last row is count number
strBankCode = wsFIS.Cells(iMaxRowFIS, iFISBankCode)
strBankCode = CStr(strBankCode)
If Len(strBankCode) < 5 Then iMaxRowFIS = iMaxRowFIS - 1
'Debug.Print iMaxRowFIS

iStartRowFIS = 2
For iFIS = 2 To iMaxRowFIS
    strBankCode = wsFIS.Cells(iFIS, iFISBankCode)
    strBankCode = Replace(strBankCode, " ", "")
    strBankCode = UCase(strBankCode)
    If strBankCode = "FISCODE" Then
        iStartRowFIS = iFIS + 1
        Exit For
    End If
Next iFIS


Cells(iStartRowFIS - 1, iFISCheck) = "Is Read"
'Debug.Print iStartRowFIS
'For iFIS = 371 To 371
For iFIS = iStartRowFIS To iMaxRowFIS

    'If wsFIS.Cells(iFIS, iFISAmt) <> 0 Then
        strBankCode = wsFIS.Cells(iFIS, iFISBankCode)
    
        For iCP = 2 To iMaxRowCP
            strCPBankCode = wsCP.Cells(iCP, iCPBankCode)

            If InStr(strCPBankCode, strBankCode) > 0 Then
                'Debug.Print iFIS
                'Debug.Print iCP
                'Debug.Print InStr(strCPBankCode, strBankCode)
                wsFIS.Cells(iFIS, iFISCheck) = wsFIS.Cells(iFIS, iFISCheck) + 1
                
                strCPFISAmt = CStr(wsCP.Cells(iCP, iCPAmtBank).Formula)
                'Debug.Print strCPFISAmt
                If strCPFISAmt = "" Then
                    'Debug.Print "Emtpy"
                    wsCP.Cells(iCP, iCPAmtBank) = wsFIS.Cells(iFIS, iFISAmt)
                    GoTo FISContinue
                Else
                    'strNumber = CStr(wsCP.Cells(iCP, iCPAmtBank))
                    If InStr(strCPFISAmt, "=") > 0 Then
                        wsCP.Cells(iCP, iCPAmtBank) = strCPFISAmt & "+" & CStr(wsFIS.Cells(iFIS, iFISAmt))
                        GoTo FISContinue
                    Else
                        wsCP.Cells(iCP, iCPAmtBank) = "=" & strCPFISAmt & "+" & CStr(wsFIS.Cells(iFIS, iFISAmt))
                        GoTo FISContinue
                    End If
                End If
            End If
        Next iCP
        wsFIS.Cells(iFIS, iFISCheck) = "Missing"
    'End If
FISContinue:
    
Next iFIS

Set wsBankVar = Worksheets("Bank Code Variance")
wsBankVar.Select
Set rngBankVar = Columns(iBankVarAcct)

wsFIS.Select
For iFIS = iStartRowFIS To iMaxRowFIS
    If Cells(iFIS, iFISCheck) = "Missing" Then
        strBankCode = Cells(iFIS, iFISBankCode)
        Set rngFound = rngBankVar.Find(strBankCode, LookIn:=xlValues, lookat:=xlWhole)
        If Not rngFound Is Nothing Then Cells(iFIS, iFISCheck) = "Var"
    End If
Next iFIS



wsCP.Select
End Sub


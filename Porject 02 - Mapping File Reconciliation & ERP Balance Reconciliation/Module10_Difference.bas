Attribute VB_Name = "Module10_Difference"
Option Explicit

Sub Write_Difference()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsCP As Worksheet
Dim iMaxRowCP As Integer
Dim iMaxColCP As Integer
Dim i As Integer
Dim iCPRowTotal As Long
Dim iCPRowERPFIS As Long
Dim iCPRowDifference As Long
Dim dCPERPAmt As Double
Dim dCPFISAmt As Double


Dim wsERP As Worksheet
Dim iMaxRowERP As Long

Dim wsFIS As Worksheet
Dim iMaxRowFIS As Integer
Dim iFISTotalRow As Integer
Dim strCell As String

Set wsCP = Worksheets("Cash Project")
wsCP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCP = lRealLastRow
iMaxColCP = lRealLastCol
If iMaxRowCP < 2 Then Exit Sub

'delete all zero line, zero for both ERP and FIS
For i = iMaxRowCP To 2 Step -1
    dCPERPAmt = CDbl(wsCP.Cells(i, iCPAmtERP))
    dCPFISAmt = CDbl(wsCP.Cells(i, iCPAmtBank))
    
    If Abs(dCPERPAmt) < 0.01 And Abs(dCPFISAmt) < 0.01 Then Rows(i).Delete
Next i

Call DeleteUnusedFormats
Cells(1, 1).Select

wsCP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCP = lRealLastRow
iMaxColCP = lRealLastCol
If iMaxRowCP < 2 Then Exit Sub
Debug.Print iMaxRowCP

'Generate formula for each line, to calculate difference
For i = 2 To iMaxRowCP
    Cells(i, iCPDiff) = "=" & NumberToLetter(iCPAmtERP) & CStr(i) & "-" & NumberToLetter(iCPAmtBank) & CStr(i)
Next i

'To add summary at bottom
iCPRowTotal = iMaxRowCP + 3
iCPRowERPFIS = iCPRowTotal + 2
iCPRowDifference = iCPRowERPFIS + 2

'To add border
With wsCP.Range(Cells(iMaxRowCP + 2, 1), Cells(iMaxRowCP + 2, iMaxColCP)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
End With


'Write total line
wsCP.Cells(iCPRowTotal, 1) = "TOTAL (Cash Position)"
wsCP.Cells(iCPRowTotal, iCPAmtERP) = "=SUM(" & NumberToLetter(iCPAmtERP) & "2:" & NumberToLetter(iCPAmtERP) & CStr(iMaxRowCP) & ")"
wsCP.Cells(iCPRowTotal, iCPAmtBank) = "=SUM(" & NumberToLetter(iCPAmtBank) & "2:" & NumberToLetter(iCPAmtBank) & CStr(iMaxRowCP) & ")"
wsCP.Cells(iCPRowTotal, iCPDiff) = "=SUM(" & NumberToLetter(iCPDiff) & "2:" & NumberToLetter(iCPDiff) & CStr(iMaxRowCP) & ")"

'write ERP FIS total
wsCP.Cells(iCPRowERPFIS, 1) = "ERP & FIS"

'find max row of ERP sheet
Set wsERP = Worksheets("ERP")
wsERP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowERP = lRealLastRow
wsCP.Select
If iMaxRowERP > 1 Then
    wsCP.Cells(iCPRowERPFIS, iCPAmtERP) = "=SUM(ERP!" & NumberToLetter(iERPAmt) & "2:" & NumberToLetter(iERPAmt) & CStr(iMaxRowERP) & ")"
End If

'find total in FIS sheet
Set wsFIS = Worksheets("FIS")
wsFIS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow

iFISTotalRow = 0
For i = 2 To iMaxRowFIS
    strCell = wsFIS.Cells(i, 1)
    strCell = UCase(Replace(strCell, " ", ""))
    If strCell = "TOTAL" Then
        iFISTotalRow = i
        Exit For
    End If
Next i

wsCP.Select
If iFISTotalRow > 0 Then
    wsCP.Cells(iCPRowERPFIS, iCPAmtBank) = "=FIS!" & NumberToLetter(iFISAmt) & CStr(iFISTotalRow)
Else
    wsCP.Cells(iCPRowERPFIS, iCPAmtBank) = "=SUM(FIS!" & NumberToLetter(iFISAmt) & "2:" & NumberToLetter(iFISAmt) & CStr(iMaxRowFIS) & ")"
End If

wsCP.Cells(iCPRowERPFIS, iCPDiff) = "=" & NumberToLetter(iCPAmtERP) & CStr(iCPRowERPFIS) & "-" & NumberToLetter(iCPAmtBank) & CStr(iCPRowERPFIS)


'write difference row
wsCP.Cells(iCPRowDifference, 1) = "Difference"
wsCP.Cells(iCPRowDifference, iCPAmtERP) = "=" & NumberToLetter(iCPAmtERP) & CStr(iCPRowTotal) & "-" & NumberToLetter(iCPAmtERP) & CStr(iCPRowERPFIS)
wsCP.Cells(iCPRowDifference, iCPAmtBank) = "=" & NumberToLetter(iCPAmtBank) & CStr(iCPRowTotal) & "-" & NumberToLetter(iCPAmtBank) & CStr(iCPRowERPFIS)
wsCP.Cells(iCPRowDifference, iCPDiff) = "=" & NumberToLetter(iCPAmtERP) & CStr(iCPRowDifference) & "-" & NumberToLetter(iCPAmtBank) & CStr(iCPRowDifference)

wsCP.Select
Cells(iCPRowDifference, 1).Select

End Sub

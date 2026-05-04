Attribute VB_Name = "Module20_Validation"
Option Explicit


Sub Validation()

Dim wsVal As Worksheet

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim dAcctTotalAMT As Double
Dim sGLClearing As String
Dim dAcctAMT As Double
Dim dAcctAMTMatched As Double
Dim dAcctAMTUnmatched As Double
Dim iIfMatched As Integer
Dim iRowMatchedGL As Integer

Dim wsJE As Worksheet
Dim iMaxRowJE As Integer
Dim iRowJE As Integer
Dim rngMatchedGL As Range
Dim rngFound As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim dDrAMT As Double
Dim dCrAMT As Double
Dim sKeyCode As String
Dim dAMT As Double

Set wsVal = Worksheets("Validation")
wsVal.Select
Cells.Select
Selection.Delete
Cells(2, 2) = "JE UPload"
Cells(4, 2) = "Debit"
Cells(5, 2) = "Credit"
Cells(8, 2) = "Difference"
Cells(1, 1).Select

Cells(2, 5) = "GL"
Cells(2, 6) = "Matched AMT"
Cells(2, 7) = "Unmatched AMT"

Set rngMatchedGL = Columns(5)
iRowMatchedGL = 2


'Verify Matching
Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow

For iRowSAP = 2 To iMaxRowSAP
'For iRowSAP = 2 To 3
    iIfMatched = 0
    'To check Row is higlighted, ie, matched,
    If Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = xlNone And InStr(UCase(Cells(iRowSAP, iColSAPClear)), "OFFSET") = 0 Then
        iIfMatched = 2
    Else
        iIfMatched = 1

    End If
    
    sGLClearing = Cells(iRowSAP, iColSAPGL)
    dAcctAMT = Cells(iRowSAP, iColSAPAMT)
    
    Set rngFound = rngMatchedGL.Find(sGLClearing, LookIn:=xlValues, lookat:=xlWhole)
    If rngFound Is Nothing Then
        iRowMatchedGL = iRowMatchedGL + 1
        wsVal.Cells(iRowMatchedGL, 5) = sGLClearing
        wsVal.Cells(iRowMatchedGL, 5 + iIfMatched) = wsVal.Cells(iRowMatchedGL, 5 + iIfMatched) + dAcctAMT
    Else
        rngFound.Cells(1, iIfMatched + 1) = rngFound.Cells(1, iIfMatched + 1) + dAcctAMT
    End If
    
    'Debug.Print sGLClearing
    'Debug.Print dAcctAMT
    
Next iRowSAP

dAcctAMTMatched = 0
dAcctAMTUnmatched = 0
For iRowSAP = 3 To iRowMatchedGL
    dAcctAMTMatched = wsVal.Cells(iRowSAP, 6) + dAcctAMTMatched
    dAcctAMTUnmatched = wsVal.Cells(iRowSAP, 7) + dAcctAMTUnmatched
Next iRowSAP

wsVal.Cells(iRowMatchedGL + 3, 5) = "Total"
wsVal.Cells(iRowMatchedGL + 3, 6) = dAcctAMTMatched
wsVal.Cells(iRowMatchedGL + 3, 7) = dAcctAMTUnmatched



'Verify JE
Set wsJE = Worksheets("3 - C-SAP Standard Template")
wsJE.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowJE = lRealLastRow
If iMaxRowJE < 5 Then Exit Sub


dDrAMT = 0
dCrAMT = 0
For iRowJE = 5 To iMaxRowJE
    sKeyCode = wsJE.Cells(iRowJE, 12)
    dAMT = CDbl(wsJE.Cells(iRowJE, 19))
    
    Select Case sKeyCode
        Case "40"
            dDrAMT = dDrAMT + dAMT
        Case "50"
            dCrAMT = dCrAMT + dAMT
        Case "21"
            dDrAMT = dDrAMT + dAMT
        Case "31"
            dCrAMT = dCrAMT + dAMT
    End Select
Next iRowJE

'Debug.Print dDrAMT
'Debug.Print dCrAMT

wsVal.Select
wsVal.Cells(4, 3) = dDrAMT
wsVal.Cells(5, 3) = dCrAMT
wsVal.Cells(8, 3) = "=C4-C5"

Range(Cells(iRowMatchedGL + 1, 5), Cells(iRowMatchedGL + 1, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(6, 2), Cells(6, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous

Cells(4, 3).Style = "Currency"
Cells(5, 3).Style = "Currency"
Cells(8, 3).Style = "Currency"

Columns(6).Style = "Currency"
Columns(7).Style = "Currency"

'wsJE.Select
wsVal.Select
Cells.Select
Selection.EntireColumn.AutoFit
Cells(1, 1).Select
End Sub

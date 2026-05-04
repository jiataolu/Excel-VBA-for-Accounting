Attribute VB_Name = "Module08_Validation"
Option Explicit

Sub Validation()

Call Validation_CAD
Call Validation_USD
End Sub
Sub Validation_CAD()

Dim wsVal As Worksheet

Dim wsJE As Worksheet
Dim iMaxRowJE As Integer
Dim iRowJE As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim dDrAMT As Double
Dim dCrAMT As Double
Dim sKeyCode As String
Dim dAMT As Double

Set wsVal = Worksheets("Validation CAD")
wsVal.Select
Cells.Select
Selection.Delete
Cells(2, 2) = "JE UPload"
Cells(4, 2) = "Debit"
Cells(5, 2) = "Credit"
Cells(8, 2) = "Difference"
Cells(1, 1).Select

Set wsJE = Worksheets(Sheet05Name_JEUploadCAD)
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

Range(Cells(6, 2), Cells(6, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous

Cells(4, 3).Style = "Currency"
Cells(5, 3).Style = "Currency"
Cells(8, 3).Style = "Currency"
wsJE.Select
Cells(1, 1).Select
End Sub

Sub Validation_USD()

Dim wsVal As Worksheet

Dim wsJE As Worksheet
Dim iMaxRowJE As Integer
Dim iRowJE As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim dDrAMT As Double
Dim dCrAMT As Double
Dim sKeyCode As String
Dim dAMT As Double

Set wsVal = Worksheets("Validation USD")
wsVal.Select
Cells.Select
Selection.Delete
Cells(2, 2) = "JE UPload"
Cells(4, 2) = "Debit"
Cells(5, 2) = "Credit"
Cells(8, 2) = "Difference"
Cells(1, 1).Select

Set wsJE = Worksheets(Sheet05Name_JEUploadUSD)
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

Range(Cells(6, 2), Cells(6, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous

Cells(4, 3).Style = "Currency"
Cells(5, 3).Style = "Currency"
Cells(8, 3).Style = "Currency"
wsJE.Select
Cells(1, 1).Select
End Sub

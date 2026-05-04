Attribute VB_Name = "Module18_MatchingPT"
Option Explicit
Option Base 1

Sub Matching_Pivot_Table()

Dim wsPT As Worksheet
Dim lRealLastRow As Long
Dim lRealLastCol As Long
Dim iMaxRowPT As Integer
Dim iMaxColPT As Integer

Dim iColBankType As Integer
Dim iColBankDate As Integer
Dim iColBankAmt As Integer
Dim iColBankMatch As Integer
Dim iRowStartBank As Integer
Dim iRowEndBank As Integer
Dim sBankType As String
Dim dBankDate As Date
Dim dBankAmt As Double

Dim iColGLType As Integer
Dim iColGLDate As Integer
Dim iColGLAmt As Integer
Dim iColGLMatch As Integer
Dim iRowStartGL As Integer
Dim iRowEndGL As Integer
Dim sGLType As String
Dim dGLDate As Date
Dim dGLAmt As Double


Dim rngBank As Range
Dim rngGL As Range

Dim iGL As Integer
Dim jBank As Integer
Dim iFound As Integer
Dim iCountMatching As Long
Dim sMatchingNumber As String


Dim wsGLDetail As Worksheet
Dim iMaxRowGLDetail As Long
Dim arrGLDetail As Variant
Dim iArrGLDetail As Long
Dim sGLDetailType As String
Dim dGLDetailDate As Date

Dim arrMatchGL As Variant
Dim iArrMatchGL As Long

Dim wsBankDetail As Worksheet
Dim iMaxRowBankDetail As Long
Dim arrBankDetail As Variant
Dim iArrBankDetail As Long
Dim sBankDetailType As String
Dim dBankDetailDate As Date

Dim arrMatchBank As Variant
Dim iArrMatchBank As Long



Set wsPT = Worksheets(SheetNamePivotTableGLBank)
wsPT.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxColPT = lRealLastCol

' 1 - find range index of GL
iMaxRowPT = Range("A65536").End(xlUp).Row
Cells(iMaxRowPT, 1).Select
Set rngGL = ActiveCell.CurrentRegion
rngGL.Select
iColGLType = rngGL.Column
iColGLDate = iColGLType + 1
iColGLAmt = rngGL.Column + rngGL.Columns.Count - 1
iColGLMatch = iColGLAmt + 1
iRowStartGL = rngGL.Row
iRowEndGL = rngGL.Row + rngGL.Rows.Count - 1

'Debug.Print iColGLType
'Debug.Print iColGLDate
'Debug.Print iColGLAmt
'Debug.Print iRow1GL
'Debug.Print iRow2GL

' 2 - find range index of Bank
iMaxRowPT = Range(Cells(65536, iMaxColPT), Cells(65536, iMaxColPT)).End(xlUp).Row
Cells(iMaxRowPT, iMaxColPT).Select
Set rngBank = ActiveCell.CurrentRegion
rngBank.Select
iColBankType = rngBank.Column
iColBankDate = iColBankType + 1
iColBankAmt = rngBank.Column + rngBank.Columns.Count - 1
iColBankMatch = iColBankAmt + 1
iRowStartBank = rngBank.Row
iRowEndBank = rngBank.Row + rngBank.Rows.Count - 1

'Debug.Print iColBankType
'Debug.Print iColBankDate
'Debug.Print iColBankAmt
'Debug.Print iRow1Bank
'Debug.Print iRow2Bank



'Matching-1: to match GL-Pivot vs Bank-Pivot

iCountMatching = 1000
Cells(iRowStartGL + 1, iColGLMatch) = "Matching #)"
Cells(iRowStartBank + 1, iColBankMatch) = "Matching #)"

For iGL = iRowStartGL + 2 To iRowEndGL - 1
'For iGL = iRowStartGL + 2 To iRowStartGL + 3
    iFound = 0
    sGLType = Cells(iGL, iColGLType)
    dGLDate = Cells(iGL, iColGLDate)
    dGLAmt = Cells(iGL, iColGLAmt)
    'Debug.Print "GL"
    'Debug.Print sGLType
    'Debug.Print dGLDate
    'Debug.Print dGLAmt
    
    For jBank = iRowStartBank + 2 To iRowEndBank - 1
        sBankType = Cells(jBank, iColBankType)
        dBankDate = Cells(jBank, iColBankDate)
        dBankAmt = Cells(jBank, iColBankAmt)
        
        If sGLType = sBankType And Abs(dGLDate - dBankDate) < 4 And Abs(dGLAmt - dBankAmt) < 0.01 Then
            iFound = 1
            Cells(iGL, iColGLMatch) = iCountMatching
            Cells(jBank, iColBankMatch) = iCountMatching
            iCountMatching = iCountMatching + 1
            Exit For
        End If
        
    Next jBank

Next iGL

'Matching-2: GL-Pivot vs GL-detail

'Create array for GL detail data
Set wsGLDetail = Worksheets(SheetNameDataGL)
wsGLDetail.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowGLDetail = lRealLastRow
arrGLDetail = wsGLDetail.Range("W1:X" & iMaxRowGLDetail).Value

' create array for final GL matching result
arrMatchGL = wsGLDetail.Range("W1:W" & iMaxRowGLDetail).Value
arrMatchGL(1, 1) = "Matching GL-Bank"
For iArrMatchGL = 2 To UBound(arrMatchGL, 1)
    arrMatchGL(iArrMatchGL, 1) = ""
Next iArrMatchGL

wsPT.Select
For iGL = iRowStartGL + 2 To iRowEndGL - 1
'For iGL = iRowStartGL + 3 To iRowStartGL + 3
    sGLType = Cells(iGL, iColGLType)
    dGLDate = Cells(iGL, iColGLDate)
    sMatchingNumber = Cells(iGL, iColGLMatch)
    Debug.Print "Hello"
    Debug.Print sGLType
    Debug.Print dGLDate
    Debug.Print sMatchingNumber
    
    For iArrGLDetail = 2 To UBound(arrGLDetail, 1)
        sGLDetailType = arrGLDetail(iArrGLDetail, 1)
        dGLDetailDate = arrGLDetail(iArrGLDetail, 2)
        If sGLType = sGLDetailType And dGLDate = dGLDetailDate Then
            arrMatchGL(iArrGLDetail, 1) = sMatchingNumber
        End If
        
    Next iArrGLDetail
Next iGL

' write matching array into column
wsGLDetail.Select
Range(Cells(1, ColDataGLMatchingGLBank), Cells(1, ColDataGLMatchingGLBank).Resize(UBound(arrMatchGL, 1), 1)).Value = arrMatchGL


'Matching-3: Bank-Pivot vs Bank-detail

'Create array for GL detail data
Set wsBankDetail = Worksheets(SheetNameDataBank)
wsBankDetail.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBankDetail = lRealLastRow
arrBankDetail = wsBankDetail.Range("P1:Q" & iMaxRowBankDetail).Value

' create array for final GL matching result
arrMatchBank = wsBankDetail.Range("P1:P" & iMaxRowBankDetail).Value
arrMatchBank(1, 1) = "Matching GL-Bank"
For iArrMatchBank = 2 To UBound(arrMatchBank, 1)
    arrMatchBank(iArrMatchBank, 1) = ""
Next iArrMatchBank

wsPT.Select
For jBank = iRowStartBank + 2 To iRowEndBank - 1
'For jbank = iRowStartbank + 3 To iRowStartbank + 3
    sBankType = Cells(jBank, iColBankType)
    dBankDate = Cells(jBank, iColBankDate)
    sMatchingNumber = Cells(jBank, iColBankMatch)
    'Debug.Print sBankType
    'Debug.Print dBankDate
    
    For iArrBankDetail = 2 To UBound(arrBankDetail, 1)
        sBankDetailType = arrBankDetail(iArrBankDetail, 1)
        dBankDetailDate = arrBankDetail(iArrBankDetail, 2)
        If sBankType = sBankDetailType And dBankDate = dBankDetailDate Then
            arrMatchBank(iArrBankDetail, 1) = sMatchingNumber
        End If
        
    Next iArrBankDetail
Next jBank

' write matching array into column
wsBankDetail.Select
Range(Cells(1, ColDataBankMatchingGLBank), Cells(1, ColDataBankMatchingGLBank).Resize(UBound(arrMatchBank, 1), 1)).Value = arrMatchBank

wsPT.Select




'Debug.Print UBound(arrGLDetail, 1)
'For iArrGLDetail = 1 To UBound(arrGLDetail, 1)
'For iArrGLDetail = 1 To 1
'    Debug.Print "Type=" & arrGLDetail(iArrGLDetail, 1)
'    Debug.Print "Recon Date=" & arrGLDetail(iArrGLDetail, 2)
'Next iArrGLDetail

'Debug.Print UBound(arrMatchGL, 1)
'For iArrMatchGL = 1 To 1
'    Debug.Print "Type=" & arrMatchGL(iArrMatchGL, 1)
'Next iArrMatchGL


End Sub


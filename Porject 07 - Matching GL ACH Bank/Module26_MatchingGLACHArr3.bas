Attribute VB_Name = "Module26_MatchingGLACHArr3"
Option Explicit
Option Base 1

'use array for GL data,
'Use range for ACH1115 recipient ID for range search
Sub Match_GL_ACH1115_by_Array_v3()

Call Process_ACH1115_Data

Dim wsDataGL As Worksheet
Dim iMaxRowDataGL As Long
Dim iRowDataGL As Long
Dim sGLACHNumber As String
Dim dGLReconDate As Date
Dim dGLDocAmount As Double
Dim arrDataGL As Variant
Dim arrMatchGL As Variant
Dim iArrGL As Long

Dim wsACH1115 As Worksheet
Dim iMaxRowACH1115 As Long
Dim iRowACH1115 As Long
Dim rngACH1115RecID As Range
Dim rngFound As Range
Dim dACH1115DebitAmount As String
Dim sACH1115EffectiveDate As String
Dim dACH1115ReconDate As Date
Dim sACH1115RecID As String
Dim arrDataACH1115 As Variant
Dim arrMatchACH1115 As Variant
Dim iArrACH1115 As Long


Dim lRealLastRow As Long
Dim lRealLastCol As Long
Dim iCountMatching As Long

Dim rngCopy As Range
Dim rngPaste As Range

Dim wsTemp As Worksheet
Dim iMaxRowTemp As Long
Set wsTemp = Worksheets(sheetNameTemp)


Set wsACH1115 = Worksheets(sheetNameDataACH1115)
wsACH1115.Select
Set rngACH1115RecID = Columns(ColACH1115RecipientID)


'Process GL data
Set wsDataGL = Worksheets(SheetNameDataGL)
wsDataGL.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataGL = lRealLastRow
Columns(ColDataGLMatchingGLACH1115).ClearContents

'to build array for GL data
wsTemp.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select
wsDataGL.Select
Set rngCopy = Columns(ColDataGLReconDate)
Set rngPaste = wsTemp.Cells(1, 1)
rngCopy.Copy Destination:=rngPaste
Set rngCopy = Columns(ColDataGLACHNumber)
Set rngPaste = wsTemp.Cells(1, 2)
rngCopy.Copy Destination:=rngPaste
Set rngCopy = Columns(ColDataGLDocAmount)
Set rngPaste = wsTemp.Cells(1, 3)
rngCopy.Copy Destination:=rngPaste

wsTemp.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowTemp = lRealLastRow
arrDataGL = wsTemp.Range("A1:C" & iMaxRowTemp).Value

'TO build array for match result for GL Data, and initialize value as empty
arrMatchGL = wsTemp.Range("B1:B" & iMaxRowTemp).Value
arrMatchGL(1, 1) = "Matching GL-ACH1115)"
For iArrGL = 2 To UBound(arrMatchGL, 1)
    arrMatchGL(iArrGL, 1) = ""
Next iArrGL

iCountMatching = 20000


For iArrGL = 2 To UBound(arrDataGL, 1)
'For iArrGL = 2 To 5000
    sGLACHNumber = arrDataGL(iArrGL, 2)
    sGLACHNumber = Replace(sGLACHNumber, " ", "")
    dGLReconDate = arrDataGL(iArrGL, 1)
    dGLDocAmount = arrDataGL(iArrGL, 3)
    
    If sGLACHNumber = "" Then GoTo CONTINUENEXTARRAYGLLINE
    Set rngFound = rngACH1115RecID.Find(sGLACHNumber, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        dACH1115DebitAmount = rngFound(1, ColACH1115DebitAmount - ColACH1115RecipientID + 1)
        dACH1115ReconDate = rngFound(1, ColACH1115ReconDate - ColACH1115RecipientID + 1)
        
        If dACH1115ReconDate = dGLReconDate And Abs(dACH1115DebitAmount - dGLDocAmount) < 0.01 Then
            rngFound.Cells(1, ColACH1115MatchingGLACH1115 - ColACH1115RecipientID + 1) = iCountMatching

            arrMatchGL(iArrGL, 1) = iCountMatching
            
            iCountMatching = iCountMatching + 1
        End If

    End If

CONTINUENEXTARRAYGLLINE:
Next iArrGL

'write array back to GL data sheet
wsDataGL.Select
Range(Cells(1, ColDataGLMatchingGLACH1115), Cells(1, ColDataGLMatchingGLACH1115).Resize(UBound(arrMatchGL, 1), 1)).Value = arrMatchGL

'wsACH1115.Select
'Range(Cells(1, ColACH1115MatchingGLACH1115), Cells(1, ColACH1115MatchingGLACH1115).Resize(UBound(arrMatchACH1115, 1), 1)).Value = arrMatchACH1115




End Sub


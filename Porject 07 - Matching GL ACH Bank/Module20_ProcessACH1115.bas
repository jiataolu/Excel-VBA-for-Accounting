Attribute VB_Name = "Module20_ProcessACH1115"
Option Explicit

Sub Process_ACH1115_Data()

Dim wsACH1115 As Worksheet
Dim iMaxRowACH1115 As Long
Dim iRowACH1115 As Long
Dim sACH1115EffectiveDate As String
Dim dACH1115ReconDate As Date


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsACH1115 = Worksheets(sheetNameDataACH1115)
wsACH1115.Select
Call DeleteUnusedFormats
Columns(ColACH1115MatchingGLACH1115).ClearContents
Columns(ColACH1115ReconDate).ClearContents

Cells(1, ColACH1115ReconDate) = "Recon Date"
Cells(1, ColACH1115MatchingGLACH1115) = "Matching GL-ACH1115"
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowACH1115 = lRealLastRow


For iRowACH1115 = 2 To iMaxRowACH1115
    sACH1115EffectiveDate = Cells(iRowACH1115, ColACH1115EffectiveDate)
    dACH1115ReconDate = DateSerial(Left(sACH1115EffectiveDate, 4), Mid(sACH1115EffectiveDate, 5, 2), Right(sACH1115EffectiveDate, 2))
    Cells(iRowACH1115, ColACH1115ReconDate) = dACH1115ReconDate
Next iRowACH1115

End Sub

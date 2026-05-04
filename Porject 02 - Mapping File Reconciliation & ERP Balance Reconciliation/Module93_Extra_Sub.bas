Attribute VB_Name = "Module93_Extra_Sub"
Option Explicit

Sub Find_Col_Index()
Dim SheetNameStep As String
SheetNameStep = "Step_a"

Dim wsStep As Worksheet
Dim iLastRowStep As Integer
Dim iStepColIndexRecon As Integer
Dim iStepColIndexSource As Integer
Dim iRowStep As Integer


Dim sSheetNameRecon As String
Dim sHeaderNameRecon As String
Dim sFileNameSource As String
Dim sSheetNameSource As String
Dim sHeaderNameSource As String


Dim wsRecon As Worksheet
Dim iLastRowRecon As Integer
Dim iLastColRecon As Integer
Dim iColHeaderinRecon As Integer

Dim wkbSource As Workbook
Dim wsSource As Worksheet
Dim iLastRowSource As Integer
Dim iLastColSource As Integer
Dim iColHeaderinSource As Integer


Dim lRealLastRow As Long
Dim lRealLastColumn As Long
Dim iRow As Integer
Dim jCol As Integer
Dim sContent As String


Set wsStep = Worksheets(SheetNameStep)
wsStep.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastColumn = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iLastRowStep = lRealLastRow
If iLastRowStep < 2 Then Exit Sub
'iStepColIndexRecon = lRealLastColumn + 1
'iStepColIndexSource = lRealLastColumn + 2
wsStep.Columns(StepColIndexHeaderNameinSource).Delete
wsStep.Columns(StepColIndexHeaderNameinRecon).Delete

wsStep.Cells(1, StepColIndexHeaderNameinRecon) = "Col-Recon"
wsStep.Cells(1, StepColIndexHeaderNameinSource) = "Col-Source"



' Initialization for:
' Recon Sheet, should be in this macro
' Source File and Source Sheet
sSheetNameRecon = wsStep.Cells(2, StepColReconSheetName)
sFileNameSource = wsStep.Cells(2, StepColSourceWorkBookName)
sSheetNameSource = wsStep.Cells(2, StepColSourceSheetName)


Set wsRecon = Worksheets(sSheetNameRecon)
wsRecon.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastColumn = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iLastRowRecon = lRealLastRow
iLastColRecon = lRealLastColumn

Set wkbSource = Workbooks.Open(GetWorkPath & "\" & sFileNameSource)
wkbSource.Activate
Set wsSource = Worksheets(sSheetNameSource)
wsSource.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastColumn = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iLastRowSource = lRealLastRow
iLastColSource = lRealLastColumn

'Debug.Print iLastColRecon
'Debug.Print iLastColSource

For iRowStep = 2 To iLastRowStep
'For iRowStep = 2 To 2

    'Header name in both recon sheet and source sheet
    sHeaderNameRecon = UCase(Replace(wsStep.Cells(iRowStep, StepColReconHeaderName), " ", ""))
    sHeaderNameSource = UCase(Replace(wsStep.Cells(iRowStep, StepColSourceHeaderName), " ", ""))
    
    If sHeaderNameSource = "" Then GoTo CONTINUENEXTSTEPLINE
    
    
    'Find Col index of searched header name in Recon Sheet
    ThisWorkbook.Activate
    wsRecon.Select

    For iRow = 1 To iLastRowRecon
        For jCol = 1 To iLastColRecon
            sContent = wsRecon.Cells(iRow, jCol)
            If UCase(Replace(sContent, " ", "")) = sHeaderNameRecon Then
                ThisWorkbook.Activate
                wsStep.Select
                wsStep.Cells(iRowStep, StepColIndexHeaderNameinRecon) = jCol
                GoTo CONTINUEWITHSOURCESHEET
            End If
        
        Next jCol
    Next iRow
    
CONTINUEWITHSOURCESHEET:
    
    'Find Col index of searched header name in Source Sheet
    wkbSource.Activate
    wsSource.Select
    For iRow = 1 To iLastRowSource
        For jCol = 1 To iLastColSource
            sContent = wsSource.Cells(iRow, jCol)
            If UCase(Replace(sContent, " ", "")) = sHeaderNameSource Then
                ThisWorkbook.Activate
                wsStep.Select
                wsStep.Cells(iRowStep, StepColIndexHeaderNameinSource) = jCol
                GoTo CONTINUENEXTSTEPLINE
            End If
        Next jCol
    Next iRow
    
    
CONTINUENEXTSTEPLINE:
    
Next iRowStep


wkbSource.Close savechanges:=False

wsStep.Select

End Sub


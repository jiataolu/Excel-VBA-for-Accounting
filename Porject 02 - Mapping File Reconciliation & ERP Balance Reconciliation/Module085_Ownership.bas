Attribute VB_Name = "Module085_Ownership"
Option Explicit

Sub Mapping_085_Ownership()

Dim sFileNameRecon As String
Dim wkbRecon As Workbook
Dim wsRecon As Worksheet
Dim iMaxRowRecon As Long
Dim iRowRecon As Long
Dim iMaxColRecon As Integer
Dim sReconBizUnit As String
Dim sReconAccount As String
Dim sReconBU As String
Dim sReconGL As String
Dim sReconComboBUGL As String
Dim rngRecon As Range
Dim sReconColumnToCheck As String
Dim rngReconComboBUGL As Range
Dim rngFound As Range
Dim iFound As Integer

Dim sReconTeam As String
Dim sReconReviewer As String
Dim sReconApprover As String
Dim sReconPreparer As String
Dim sOwnershipInfo As String

Dim wsMap As Worksheet
Dim iMaxRowMap As Integer
Dim iRowMap As Integer
Dim sMapSAPBU As String
Dim sMapSapGL As String
Dim sMapLocalBU As String
Dim sMapLocalGL As String
Dim sMapComboSAPBUGL As String
Dim sMapComboLocalBUGL As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long


sFileNameRecon = GetWorkPath & "\" & FileNameRecon
'Debug.Print sFileNameTreasury
Set wkbRecon = Workbooks.Open(sFileNameRecon)
wkbRecon.Activate
Set wsRecon = Worksheets(1)
wsRecon.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowRecon = lRealLastRow
iMaxColRecon = lRealLastCol
If iMaxRowRecon < 2 Then Exit Sub

'sColumnToCheck = "3,5"
Set rngRecon = Range(Cells(1, 1), Cells(iMaxRowRecon, iMaxColRecon))
rngRecon.RemoveDuplicates Columns:=Array(3, 5), Header:=xlYes
Cells(1, 1).Select
Call DeleteUnusedFormats

'In recon file, key is combo of BU + GL
For iRowRecon = 2 To iMaxRowRecon
'For iRowRecon = 19289 To 19289
    sReconBizUnit = wsRecon.Cells(iRowRecon, ColReconBizUnit)
    sReconAccount = wsRecon.Cells(iRowRecon, ColReconAccount)
    
    sReconBU = Read_BUGL(sReconBizUnit)
    sReconGL = Read_BUGL(sReconAccount)
    sReconComboBUGL = sReconBU & "-" & sReconGL
    wsRecon.Cells(iRowRecon, ColReconGL) = sReconGL
    wsRecon.Cells(iRowRecon, ColReconBU) = sReconBU
    wsRecon.Cells(iRowRecon, ColReconComboBUGL) = sReconComboBUGL
    
    'Debug.Print sBizUnit
    'Debug.Print sAccount
    'Debug.Print sBU
    'Debug.Print sGL
    
Next iRowRecon

Set rngReconComboBUGL = Columns(ColReconComboBUGL)



' check Mapping sheet one by one.
ThisWorkbook.Activate
Set wsMap = Worksheets(SheetNameMapping)
wsMap.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMap = lRealLastRow
If iMaxRowMap < 2 Then Exit Sub


For iRowMap = 2 To iMaxRowMap
'For iRowMap = 6 To 8
    iFound = 0
    sOwnershipInfo = ""
    
    
    sMapSAPBU = wsMap.Cells(iRowMap, ColMapFISBUCode)
    sMapSapGL = wsMap.Cells(iRowMap, ColMapFISSapGL)
    sMapComboSAPBUGL = sMapSAPBU & "-" & sMapSapGL
    
    sMapLocalBU = wsMap.Cells(iRowMap, ColMapLocalBU)
    sMapLocalGL = wsMap.Cells(iRowMap, ColMapLocalGL)
    sMapComboLocalBUGL = sMapLocalBU & "-" & sMapLocalGL
    
    'search with SAP BU & GL
    Set rngFound = rngReconComboBUGL.Find(sMapComboSAPBUGL, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not rngFound Is Nothing Then
        iFound = 1
        GoTo RECONINFO
    End If
    
    'serach by local BU & GL
    Set rngFound = rngReconComboBUGL.Find(sMapComboLocalBUGL, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not rngFound Is Nothing Then iFound = 1
    
RECONINFO:
    
    'if both SAP & Local BU & GL can't be found in recon file, then finish this line
    If iFound = 0 Then GoTo WRITINGOWNERSHIP
    
    sReconTeam = rngFound.Cells(1, ColReconTEAM - ColReconComboBUGL + 1)
    sReconReviewer = rngFound.Cells(1, ColReconReviewer - ColReconComboBUGL + 1)
    sReconApprover = rngFound.Cells(1, ColReconApprover - ColReconComboBUGL + 1)
    sReconPreparer = rngFound.Cells(1, ColReconPreparer - ColReconComboBUGL + 1)
    
    If InStr(sReconTeam, "Bank & Cash Accounting") > 0 Then
        sOwnershipInfo = "Bank & Cash Accounting"
        GoTo WRITINGOWNERSHIP
    End If
    
    If InStr(sReconReviewer, "Not Required") = 0 And Replace(sReconReviewer, " ", "") <> "" Then
        sOwnershipInfo = sReconReviewer
        GoTo WRITINGOWNERSHIP
    End If
    
    If InStr(sReconApprover, "Approver, BL") = 0 And Replace(sReconReviewer, " ", "") <> "" Then
        sOwnershipInfo = sReconApprover
        GoTo WRITINGOWNERSHIP
    End If

    
    'Debug.Print sOwnershipInfo
WRITINGOWNERSHIP:
    wsMap.Cells(iRowMap, ColMapOwnership) = sOwnershipInfo

CONTINUEWITHNEXTMAPPINGLINE:
Next iRowMap

wkbRecon.Close savechanges:=False
End Sub


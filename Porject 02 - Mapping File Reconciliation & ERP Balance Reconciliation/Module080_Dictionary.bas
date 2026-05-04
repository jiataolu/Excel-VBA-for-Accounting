Attribute VB_Name = "Module080_Dictionary"
Option Explicit

Sub Mapping_080_Read_Company_Code()

Dim wsMap As Worksheet
Dim iMaxRowMap As Integer
Dim iRowMap As Integer
Dim sMapCompanyCode As String
Dim sMapERP As String
Dim sMapBUName As String
Dim sMapVendorCode As String
Dim sMapParentCode As String
Dim sMapSAPBUCode As String
Dim sMapRemark As String

Dim sFileNameCompanyCode As String
Dim wkbCompanyCode As Workbook
Dim wsCompanyCode As Worksheet
Dim rngCompanyCode As Range
Dim rngFound As Range
Dim sCompanyCodeERP As String
Dim sCompanyCodeBUName As String
Dim sCompanyCodeVendorCode As String
Dim sCompanyCodeParentCode As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long

sFileNameCompanyCode = GetWorkPath & "\" & FileNameCompanyCode
'Debug.Print sFileNameTreasury
Set wkbCompanyCode = Workbooks.Open(sFileNameCompanyCode)
wkbCompanyCode.Activate
Set wsCompanyCode = Worksheets(1)
Set rngCompanyCode = Columns(ColCompanyCodeCompanyCode)


ThisWorkbook.Activate
Set wsMap = Worksheets(SheetNameMapping)
wsMap.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMap = lRealLastRow
If iMaxRowMap < 2 Then Exit Sub

For iRowMap = 2 To iMaxRowMap
'For iRowMap = 2 To 3
    sMapCompanyCode = wsMap.Cells(iRowMap, ColMapFISBUCode)
    
    sMapERP = wsMap.Cells(iRowMap, ColMapERPSystem)
    sMapBUName = wsMap.Cells(iRowMap, ColMapBUName)
    sMapVendorCode = wsMap.Cells(iRowMap, ColMapVendorCode)
    sMapParentCode = wsMap.Cells(iRowMap, ColMapParentCode)
    sMapRemark = UCase(Replace(wsMap.Cells(iRowMap, ColMapRemark), " ", ""))
    'sMapSAPBUCode = wsMap.Cells(iRowMap, ColMapSAPBUCode)
    
    
    Debug.Print "BUCode-" & sMapCompanyCode
    Debug.Print "BU Name-" & sMapBUName
    Debug.Print "ERP-" & sMapERP
    Debug.Print "Vendor Code-" & sMapVendorCode
    Debug.Print "ParentC COde-" & sMapParentCode
    
    
    Set rngFound = rngCompanyCode.Find(sMapCompanyCode, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        Debug.Print "Found"
        
        sCompanyCodeERP = rngFound.Cells(1, ColCompanyCodeERP - ColCompanyCodeCompanyCode + 1)
        sCompanyCodeBUName = rngFound.Cells(1, ColCompanyCodeBUName - ColCompanyCodeCompanyCode + 1)
        sCompanyCodeVendorCode = rngFound.Cells(1, ColCompanyCodeVendorCode - ColCompanyCodeCompanyCode + 1)
        sCompanyCodeParentCode = rngFound.Cells(1, ColCompanyCodeParentCode - ColCompanyCodeCompanyCode + 1)
        
        'check ERP system
        If sCompanyCodeERP <> sMapERP Then
            wsMap.Cells(iRowMap, ColMapERPSystem) = sCompanyCodeERP
            If sMapRemark <> "NEW" Then wsMap.Cells(iRowMap, ColMapERPSystem).Interior.Color = RGB(258, 258, 153)

            'wsMap.Cells(iRowMap, ColMapERPSystem).Interior.Color = RGB(258, 258, 153)
        End If
        
        'check BU Name
        If sCompanyCodeBUName <> sMapBUName Then
            wsMap.Cells(iRowMap, ColMapBUName) = sCompanyCodeBUName
            If sMapRemark <> "NEW" Then wsMap.Cells(iRowMap, ColMapBUName).Interior.Color = RGB(258, 258, 153)

            'wsMap.Cells(iRowMap, ColMapBUName).Interior.Color = RGB(258, 258, 153)
        End If
        
        'check vendor code
        If sCompanyCodeVendorCode <> sMapVendorCode Then
            wsMap.Cells(iRowMap, ColMapVendorCode) = sCompanyCodeVendorCode
            If sMapRemark <> "NEW" Then wsMap.Cells(iRowMap, ColMapVendorCode).Interior.Color = RGB(258, 258, 153)
 
            'wsMap.Cells(iRowMap, ColMapVendorCode).Interior.Color = RGB(258, 258, 153)
        End If
        
        'check parent code
        If sCompanyCodeParentCode <> sMapParentCode Then
            wsMap.Cells(iRowMap, ColMapParentCode) = sCompanyCodeParentCode
            If sMapRemark <> "NEW" Then wsMap.Cells(iRowMap, ColMapParentCode).Interior.Color = RGB(258, 258, 153)

            'wsMap.Cells(iRowMap, ColMapParentCode).Interior.Color = RGB(258, 258, 153)
        End If
        
        
        'check if SAP BU is equal to FIS BU
        'If sMapSAPBUCode <> sMapCompanyCode Then
        '    wsMap.Cells(iRowMap, ColMapSAPBUCode) = sMapCompanyCode
            'If it is not new line
        '    If sMapRemark <> "NEW" Then wsMap.Cells(iRowMap, ColMapSAPBUCode).Interior.Color = RGB(258, 258, 153)
        'End If
        
        Debug.Print sCompanyCodeERP
        Debug.Print sCompanyCodeBUName
        Debug.Print sCompanyCodeVendorCode
        Debug.Print sCompanyCodeParentCode
    
    End If
Next iRowMap

wkbCompanyCode.Close savechanges:=False

End Sub


Attribute VB_Name = "Module060_Add_New"
Option Explicit

Sub Mapping_060_Add_New_Lines()

Dim wsMap As Worksheet
Dim iMaxRowMap As Integer
Dim iRowMap As Integer
Dim iCurrentRowMap As Integer


Dim wsFIS As Worksheet
Dim iMaxRowFIS As Integer
Dim iRowFIS As Integer
Dim sFISRemark As String

Dim sFISFISCode As String
Dim sFISKyribaCode As String
Dim sFISBUCode As String

Dim varCellValue As Variant
Dim sFISSapGL As String

Dim sFISBankAcctFull As String
Dim sFISCry As String
Dim sFISProductCode As String
Dim sFISBankAcctKey As String
Dim sFISDataSource As String
Dim sFISCompanyName As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long


Set wsMap = Worksheets(SheetNameMapping)
wsMap.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMap = lRealLastRow

iCurrentRowMap = iMaxRowMap

Set wsFIS = Worksheets(SheetNameFIS)
wsFIS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
If iMaxRowFIS < 2 Then Exit Sub

For iRowFIS = 2 To iMaxRowFIS
    sFISRemark = Replace(wsFIS.Cells(iRowFIS, ColFISRemark), " ", "")
    
    'Debug.Print iRowFIS
    
    If UCase(sFISRemark) = "NEW" Then
        

        sFISFISCode = wsFIS.Cells(iRowFIS, ColFISFISCode)
        sFISKyribaCode = wsFIS.Cells(iRowFIS, ColFISKyribaCode)
        sFISBUCode = wsFIS.Cells(iRowFIS, ColFISBUCode)
        'Debug.Print wsFIS.Cells(iRowFIS, ColFISSapGL).Value
        
        varCellValue = wsFIS.Cells(iRowFIS, ColFISSapGL)
        If IsError(varCellValue) Then
            sFISSapGL = "NA"
        Else
            sFISSapGL = varCellValue
        End If
        
        'sFISSapGL = wsFIS.Cells(iRowFIS, ColFISSapGL).Value
        sFISBankAcctFull = wsFIS.Cells(iRowFIS, ColFISBankAcct)
        sFISBankAcctFull = Long_Bank_Account(sFISBankAcctFull)
        
        
        sFISCry = wsFIS.Cells(iRowFIS, ColFISCurrency)
        sFISProductCode = wsFIS.Cells(iRowFIS, ColFISProductCode)
        sFISBankAcctKey = wsFIS.Cells(iRowFIS, ColFISKeyNumber)
        sFISCompanyName = wsFIS.Cells(iRowFIS, ColFISCompanyName)
        
        iCurrentRowMap = iCurrentRowMap + 1
        wsMap.Cells(iCurrentRowMap, ColMapFISCode) = sFISFISCode
        wsMap.Cells(iCurrentRowMap, ColMapKyribaCode) = sFISKyribaCode
        wsMap.Cells(iCurrentRowMap, ColMapFISBUCode) = sFISBUCode
        wsMap.Cells(iCurrentRowMap, ColMapFISSapGL) = sFISSapGL
        wsMap.Cells(iCurrentRowMap, ColMapBankAcctFull) = sFISBankAcctFull
        wsMap.Cells(iCurrentRowMap, ColMapCry) = sFISCry
        wsMap.Cells(iCurrentRowMap, ColMapProductCode) = sFISProductCode
        wsMap.Cells(iCurrentRowMap, ColMapBankAcctKey) = sFISBankAcctKey
        wsMap.Cells(iCurrentRowMap, ColMapRemark) = sFISRemark
        wsMap.Cells(iCurrentRowMap, ColMapCompanyName) = sFISCompanyName
        
        
        If Replace(wsFIS.Cells(iRowFIS, ColFISIsinFIS), " ", "") <> "" Then
            wsMap.Cells(iCurrentRowMap, ColMapDataSource) = "Treasury"
        Else
            wsMap.Cells(iCurrentRowMap, ColMapDataSource) = "PeopleSoft"
            
        End If
    End If
    
Next iRowFIS

wsMap.Select
Call DeleteUnusedFormats
Cells(1, 1).Select

End Sub

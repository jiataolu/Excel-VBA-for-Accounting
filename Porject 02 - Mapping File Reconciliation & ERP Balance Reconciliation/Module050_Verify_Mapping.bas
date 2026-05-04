Attribute VB_Name = "Module050_Verify_Mapping"
Option Explicit

Sub Mapping_050_Verfiy_Mapping_with_IFS_and_PeopleSoft()

' Verify Sheet "Mapping Consolidation" & "FIS & PeopleSoft"
' Add in column "Remark" Found in Sheet "Mappintg Consolidation" if found
' Add in column "Remark" New in Sheet "FIS & PeopleSoft" if not found

Dim wsFIS As Worksheet
Dim iMaxRowFIS As Integer
Dim iRowFIS As Integer
Dim sFISBankAcctKey As String
Dim sFISBankAcctFull As String
Dim sFISFISCode As String
Dim sFISKyribaCode As String
Dim sFISCry As String
Dim sFISBUCode As String
Dim sFISBankAcctKeyNoZero As String

Dim varCellValue As Variant
Dim sFISSapGL As String

Dim sFISProductCode As String
Dim sFISCompanyName As String


Dim wsMap As Worksheet
Dim iMaxRowMap As Integer
Dim iMaxColMap As Integer
Dim iRowMap As Integer
Dim sMapBankAcctFull As String
Dim sMapBankAcctKey As String
Dim iLen As Integer
Dim rngMapBankAcctKey As Range
Dim rngFound As Range
Dim sMapFISCode As String
Dim sMapKyribaCode As String
Dim sMapCry As String
Dim sMapBUCode As String
Dim sMapSapGL As String
Dim sMapDataSource As String
Dim sMapProductCode As String
Dim sMapCompanyName As String
Dim rngMapKyribaCode As Range


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsBUError As Worksheet
Dim iMaxRowBUError As Integer
Dim iCurrentRowBUError As Integer
Dim rngCopy As Range
Dim rngPaste As Range


Set wsBUError = Worksheets(SheetNameBUError)
wsBUError.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBUError = lRealLastRow
If iMaxRowBUError < 1 Then
    Rows("2:" & iMaxRowBUError).Delete
End If
iCurrentRowBUError = 1

'Mapping sheet
Set wsMap = Worksheets(SheetNameMapping)
wsMap.Select
Call DeleteUnusedFormats
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMap = lRealLastRow
iMaxColMap = lRealLastCol
If iMaxRowMap > 1 Then Range(Cells(2, 1), Cells(iMaxRowMap, iMaxColMap)).Interior.Pattern = xlNone


'Generate Key-123456789 bank code for Sheet Mapping Consolidation"
wsMap.Cells(1, ColMapBankAcctKey) = "Key Acct #"
wsMap.Cells(1, ColMapRemark) = "Remark"
For iRowMap = 2 To iMaxRowMap
    sMapBankAcctKey = "Key-"
    sMapBankAcctFull = wsMap.Cells(iRowMap, ColMapBankAcctFull)
    sMapFISCode = wsMap.Cells(iRowMap, ColMapFISCode)
    
    'All accounts compare with bank account number, the last 9 digits or defind by variable LenKeyBankAcctNo
    'except, Deal with 3 lines whose bank account is only "x". Add FIS bank code
    If UCase(Replace(sMapBankAcctFull, " ", "")) = "X" Then sMapBankAcctFull = Replace(sMapBankAcctFull & sMapFISCode, " ", "")
    
    iLen = Len(sMapBankAcctFull)
    
    If iLen < LenKeyBankAcctNo Then
        sMapBankAcctKey = sMapBankAcctKey & sMapBankAcctFull
    Else
        sMapBankAcctKey = sMapBankAcctKey & Right(sMapBankAcctFull, LenKeyBankAcctNo)
    End If
    
    wsMap.Cells(iRowMap, ColMapBankAcctKey) = sMapBankAcctKey
    
Next iRowMap

Set rngMapBankAcctKey = Columns(ColMapBankAcctKey)
Set rngMapKyribaCode = Columns(ColMapKyribaCode)

' scan every line in Sheet FIS & PeopleSoft, to verify with Sheet Mapping
Set wsFIS = Worksheets(SheetNameFIS)
wsFIS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
If iMaxRowFIS < 2 Then Exit Sub

For iRowFIS = 2 To iMaxRowFIS
'For iRowFIS = 180 To 180
    sFISBankAcctKey = wsFIS.Cells(iRowFIS, ColFISKeyNumber)
    
    'Account 752-82605 is exceptional, it has 3 lines, do not touch it.
    If sFISBankAcctKey = "Key-752-82605" Then GoTo CONTINUEMAPPINGROW
    
    'To check if Bank account number is valid
    'If not valid, like 0000, then just mark it as "Found", do not check it.
    sFISBankAcctKeyNoZero = Replace(sFISBankAcctKey, "0", "")
    If sFISBankAcctKeyNoZero = "Key-" Then
        'Debug.Print sFISBankAcctKeyNoZero
        
        sMapKyribaCode = wsFIS.Cells(iRowFIS, ColFISKyribaCode)
        'Debug.Print sMapKyribaCode
        Set rngFound = rngMapKyribaCode.Find(sMapKyribaCode, LookIn:=xlValues, lookat:=xlWhole)
        If Not rngFound Is Nothing Then
            rngFound.Cells(1, ColMapRemark - ColMapKyribaCode + 1) = "Found"
        Else
            wsFIS.Cells(iRowFIS, ColFISRemark) = "New"
        End If
        GoTo CONTINUEMAPPINGROW
    End If
    
    'if bank account is valid, then verify by key number
    'If valid, then check all information.
    Set rngFound = rngMapBankAcctKey.Find(sFISBankAcctKey, LookIn:=xlValues, lookat:=xlWhole)
    If rngFound Is Nothing Then
        wsFIS.Cells(iRowFIS, ColFISRemark) = "New"
    Else
        
        sFISBankAcctFull = wsFIS.Cells(iRowFIS, ColFISBankAcct)
        'sFISBankAcctFull = Replace(sFISBankAcctFull, "'", "")
        'sFISBankAcctFull = "'" & CStr(sFISBankAcctFull)
        
        
        sFISFISCode = wsFIS.Cells(iRowFIS, ColFISFISCode)
        sFISKyribaCode = wsFIS.Cells(iRowFIS, ColFISKyribaCode)
        sFISCry = wsFIS.Cells(iRowFIS, ColFISCurrency)
        sFISBUCode = wsFIS.Cells(iRowFIS, ColFISBUCode)
        
        varCellValue = wsFIS.Cells(iRowFIS, ColFISSapGL).Value
        If IsError(varCellValue) Then
            sFISSapGL = "NA"
        Else
            sFISSapGL = varCellValue
        End If
        'sFISSapGL = wsFIS.Cells(iRowFIS, ColFISSapGL).Value
        
        
        sFISProductCode = Replace(wsFIS.Cells(iRowFIS, ColFISProductCode), " ", "")
        sFISCompanyName = wsFIS.Cells(iRowFIS, ColFISCompanyName)
        
        'Debug.Print iRowFIS
        'Debug.Print sFISFISCode
        'Debug.Print sFISSapGL
        
        sMapBankAcctFull = rngFound.Cells(1, ColMapBankAcctFull - ColMapBankAcctKey + 1)
        'sMapBankAcctFull = Replace(sMapBankAcctFull, "'", "")
        'sMapBankAcctFull = "'" & CStr(sMapBankAcctFull)
        
        sMapFISCode = rngFound.Cells(1, ColMapFISCode - ColMapBankAcctKey + 1)
        sMapKyribaCode = rngFound.Cells(1, ColMapKyribaCode - ColMapBankAcctKey + 1)
        sMapCry = rngFound.Cells(1, ColMapCry - ColMapBankAcctKey + 1)
        sMapBUCode = rngFound.Cells(1, ColMapFISBUCode - ColMapBankAcctKey + 1)
        sMapSapGL = rngFound.Cells(1, ColMapFISSapGL - ColMapBankAcctKey + 1)
        sMapProductCode = rngFound.Cells(1, ColMapProductCode - ColMapBankAcctKey + 1)
        sMapCompanyName = rngFound.Cells(1, ColMapCompanyName - ColMapBankAcctKey + 1)
        
        'To check full bank account number
        If sFISBankAcctFull <> sMapBankAcctFull Then
            rngFound.Cells(1, ColMapBankAcctFull - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            sFISBankAcctFull = Long_Bank_Account(sFISBankAcctFull)
            rngFound.Cells(1, ColMapBankAcctFull - ColMapBankAcctKey + 1) = sFISBankAcctFull
        End If
        
        
        'To check Bank code
        If sFISFISCode <> sMapFISCode Then
            rngFound.Cells(1, ColMapFISCode - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            rngFound.Cells(1, ColMapFISCode - ColMapBankAcctKey + 1) = sFISFISCode
        End If
        
        'To check Kyriba Code
        If sFISKyribaCode <> sMapKyribaCode Then
            rngFound.Cells(1, ColMapKyribaCode - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            rngFound.Cells(1, ColMapKyribaCode - ColMapBankAcctKey + 1) = sFISKyribaCode
        End If
        
        'To check currency
        If sFISCry <> sMapCry Then
            rngFound.Cells(1, ColMapCry - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            rngFound.Cells(1, ColMapCry - ColMapBankAcctKey + 1) = sFISCry
        End If
        
        'To check FIS BU Code
        If sFISBUCode <> sMapBUCode Then
                                
            'Update Sheet-BU Error
            iCurrentRowBUError = iCurrentRowBUError + 1
            'wsBUError.Cells(iCurrentRowBUError, ColMapCompanyName) = rngFound.Cells(1, ColMapCompanyName - ColMapBankAcctKey + 1)
            Set rngPaste = wsBUError.Cells(iCurrentRowBUError, 1)
            Set rngCopy = Range(wsMap.Cells(rngFound.Row, 1), wsMap.Cells(rngFound.Row, ColMapComment))
            rngCopy.Copy Destination:=rngPaste
            
            rngFound.Cells(1, ColMapFISBUCode - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            rngFound.Cells(1, ColMapFISBUCode - ColMapBankAcctKey + 1) = sFISBUCode
            

        End If
        
        'To check FIS SAP GL
        If sFISSapGL <> sMapSapGL Then
                                
            'Update Sheet-BU Error
            iCurrentRowBUError = iCurrentRowBUError + 1
            'wsBUError.Cells(iCurrentRowBUError, ColMapCompanyName) = rngFound.Cells(1, ColMapCompanyName - ColMapBankAcctKey + 1)
            Set rngPaste = wsBUError.Cells(iCurrentRowBUError, 1)
            Set rngCopy = Range(wsMap.Cells(rngFound.Row, 1), wsMap.Cells(rngFound.Row, ColMapComment))
            rngCopy.Copy Destination:=rngPaste
            
            rngFound.Cells(1, ColMapFISSapGL - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            rngFound.Cells(1, ColMapFISSapGL - ColMapBankAcctKey + 1) = sFISSapGL
            

        End If
        
        
        
        
        
        'To check Product Code
        If sFISProductCode <> sMapProductCode Then
            rngFound.Cells(1, ColMapProductCode - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            rngFound.Cells(1, ColMapProductCode - ColMapBankAcctKey + 1) = sFISProductCode
        End If
        
        'To check data source
        If Replace(wsFIS.Cells(iRowFIS, ColFISIsinFIS), " ", "") <> "" Then
            rngFound.Cells(1, ColMapDataSource - ColMapBankAcctKey + 1) = "Treasury"
        Else
            rngFound.Cells(1, ColMapDataSource - ColMapBankAcctKey + 1) = "PeopleSoft"

        End If
        
        'To check Company Name
        If Replace(sFISCompanyName, " ", "") <> Replace(sMapCompanyName, " ", "") Then
            rngFound.Cells(1, ColMapCompanyName - ColMapBankAcctKey + 1).Interior.Color = RGB(255, 255, 153)
            rngFound.Cells(1, ColMapCompanyName - ColMapBankAcctKey + 1) = sFISCompanyName
        End If
        
        
        rngFound.Cells(1, ColMapRemark - ColMapBankAcctKey + 1) = "Found"
        
        
    End If
    
CONTINUEMAPPINGROW:
Next iRowFIS


wsMap.Select
End Sub



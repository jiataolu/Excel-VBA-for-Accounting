Attribute VB_Name = "Module10_ReconcileBSAmount"
Option Explicit

Sub Reconcile_Amount_PAP_with_Bank_Statement(CompanyName As String)

'Dim CompanyName As String
'CompanyName = "MSD"

Dim wsPAP As Worksheet
Dim iMaxRowPAP As Integer
Dim iPAP As Integer
Dim dPAPNetAmt As Double
Dim sPAPTradingPart As String
Dim sPAPCustomer As String
Dim sPAPBranch As String
Dim iFound As Integer

Dim wsBS As Worksheet
Dim iMaxRowBS As Integer
Dim iBS As Integer
Dim sBSEntity As String
Dim dBSAmt As Double
Dim sBankCode As String
Dim sBU As String
Dim sGL As String


Dim wsMapping As Worksheet
Dim rngBankCode As Range
Dim rngFound As Range


Dim lRealLastRow As Long
Dim lRealLastCol As Long


Set wsMapping = Worksheets("Mapping")
wsMapping.Select
Set rngBankCode = Columns(ColMappingBankCode)


Set wsBS = Worksheets("Bank Statement")
wsBS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBS = lRealLastRow

Set wsPAP = Worksheets("PAP Invoices")
wsPAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowPAP = lRealLastRow

For iPAP = 2 To iMaxRowPAP
'For iPAP = 42 To 42
    If UCase(wsPAP.Cells(iPAP, 1)) = "TOTAL" Then
        dPAPNetAmt = Round(wsPAP.Cells(iPAP, ColSAPNetAmt), 2)
        sPAPTradingPart = wsPAP.Cells(iPAP - 1, ColSAPTradingPart)
        sPAPCustomer = wsPAP.Cells(iPAP - 1, ColSAPAccount)
        sPAPBranch = wsPAP.Cells(iPAP - 1, ColSAPBranch)
        
        iFound = 0
        
        'For iBS = 8 To 8
        For iBS = 2 To iMaxRowBS
            sBSEntity = wsBS.Cells(iBS, ColBSEntity)
            dBSAmt = Round(wsBS.Cells(iBS, ColBSAMTOrig), 2)
            'Debug.Print sBSEntity
            'Debug.Print dBSAmt
            
            'Debug.Print dBSAmt - dPAPNetAmt
            If sBSEntity = CompanyName And dBSAmt = dPAPNetAmt Then
            'If dBSAmt = dPAPNetAmt Then
                'Debug.Print iBS
                wsBS.Cells(iBS, ColBSAMTPAP) = dPAPNetAmt
                wsBS.Cells(iBS, ColBSTradingPart) = sPAPTradingPart
                wsBS.Cells(iBS, ColBSCustomer) = sPAPCustomer
                wsBS.Cells(iBS, ColBSBranch) = sPAPBranch
                
                sBankCode = CStr(wsBS.Cells(iBS, ColBSAccount))
                If Len(sBankCode) > 4 Then sBankCode = Right(sBankCode, 4)
                
                iFound = 1
            End If
        Next iBS
        
        If iFound = 0 Then
            wsPAP.Rows(iPAP).Font.Color = RGB(255, 0, 0)
        Else
            'find coding info
            sBankCode = "TDB-" & sBankCode
            
            Set rngFound = rngBankCode.Find(sBankCode, LookIn:=xlValues, lookat:=xlWhole)
            
            If Not rngFound Is Nothing Then
                sBU = rngFound.Cells(1, ColMappingBU - ColMappingBankCode + 1)
                sGL = rngFound.Cells(1, ColMappingGL - ColMappingBankCode + 1)
            End If
            
            wsPAP.Cells(iPAP, ColSAPCodingInfo) = sBankCode & ", BU-" & sBU & ", GL-" & sGL
        End If
    End If
Next iPAP

wsBS.Select
Cells.Select
Cells.EntireColumn.AutoFit
Cells(1, 1).Select
End Sub

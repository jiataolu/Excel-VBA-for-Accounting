Attribute VB_Name = "Module09_Write_Missing"
Option Explicit

Sub Write_Missing_ERP_FIS()

Dim wsCP As Worksheet
Dim iMaxRowCP As Integer
Dim iRowCP As Integer
Dim iRowNewCP As Integer

Dim wsERP As Worksheet
Dim iMaxRowERP As Integer
Dim iRowERP As Integer
Dim sERPCheck As String
Dim sERPBU As String
Dim sERPGL As String
Dim sERPAMT As String

Dim wsFIS As Worksheet
Dim iMaxRowFIS As Integer
Dim iRowFIS As Integer
Dim sFISCheck As String
Dim sFISAMT As String
Dim sFISBankCode As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsCP = Worksheets("Cash Project")
wsCP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCP = lRealLastRow
If iMaxRowCP < 2 Then Exit Sub
'Debug.Print iMaxRowCP
iRowNewCP = iMaxRowCP

Set wsERP = Worksheets("ERP")
wsERP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowERP = lRealLastRow
If iMaxRowERP < 2 Then GoTo PROCESSFIS

For iRowERP = 2 To iMaxRowERP
    sERPCheck = wsERP.Cells(iRowERP, iERPCheck)
    If UCase(Replace(sERPCheck, " ", "")) = "MISSING" Then
        'Debug.Print iRowERP
        sERPBU = wsERP.Cells(iRowERP, iERPCompanyCode)
        sERPGL = wsERP.Cells(iRowERP, iERPSAPAcct)
        sERPAMT = wsERP.Cells(iRowERP, iERPAmt)
        'Debug.Print sERPBU
        'Debug.Print sERPGL
        'Debug.Print sERPAMT
        
        iRowNewCP = iRowNewCP + 1
        'Debug.Print iRowNewCP
        wsCP.Cells(iRowNewCP, iCPCat) = "Missing-ERP"
        wsCP.Cells(iRowNewCP, iCPBU) = sERPBU
        wsCP.Cells(iRowNewCP, iCPGL) = sERPGL
        wsCP.Cells(iRowNewCP, iCPAmtERP) = sERPAMT
        
        wsCP.Cells(iRowNewCP, iCPKey) = sERPBU & "-" & sERPGL
        
        
    End If
Next iRowERP

PROCESSFIS:

Set wsFIS = Worksheets("FIS")
wsFIS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
If iMaxRowFIS < 2 Then GoTo FINISH

For iRowFIS = 2 To iMaxRowFIS
    sFISCheck = wsFIS.Cells(iRowFIS, iFISCheck)
    If UCase(Replace(sFISCheck, " ", "")) = "MISSING" Then
        'Debug.Print iRowERP
        sFISBankCode = wsFIS.Cells(iRowFIS, iFISBankCode)
        sFISAMT = wsFIS.Cells(iRowFIS, iFISAmt)
        Debug.Print sFISBankCode
        Debug.Print sFISAMT
        
        iRowNewCP = iRowNewCP + 1
        Debug.Print iRowNewCP
        wsCP.Cells(iRowNewCP, iCPCat) = "Missing-FIS"
        wsCP.Cells(iRowNewCP, iCPBankCode) = sFISBankCode
        wsCP.Cells(iRowNewCP, iCPAmtBank) = sFISAMT
        
    End If
Next iRowFIS

FINISH:

End Sub

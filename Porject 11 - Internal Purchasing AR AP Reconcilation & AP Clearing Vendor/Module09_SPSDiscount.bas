Attribute VB_Name = "Module09_SPSDiscount"
Option Explicit

Sub SPS_Discount_Info(CompanyName As String)

If CompanyName <> "SPS" Then Exit Sub

Dim wsPAP As Worksheet
Dim iMaxPAP As Integer
Dim iRowPAP As Integer
Dim strKeyPAP As String

Dim wsDis As Worksheet
Dim iMaxDis As Integer
Dim iRowDis As Integer
Dim strKeyDis As String
Dim strProfitCenter As String
Dim strProduct As String
Dim strCustomer As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsDis = Worksheets("DISCOUNT INFO")
wsDis.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxDis = lRealLastRow

Set wsPAP = Worksheets("PAP Invoices")
wsPAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxPAP = lRealLastRow

For iRowPAP = 2 To iMaxPAP
    If wsPAP.Cells(iRowPAP, 1) = "Total" Then
        If wsPAP.Cells(iRowPAP, ColSAPDis) <> 0 Then
            If wsPAP.Cells(iRowPAP - 1, ColSAPBranch) <> "" Then
                strKeyPAP = wsPAP.Cells(iRowPAP - 1, ColSAPAccount) & "-" & wsPAP.Cells(iRowPAP - 1, ColSAPBranch)
            Else
                strKeyPAP = wsPAP.Cells(iRowPAP - 1, ColSAPAccount) & "-" & wsPAP.Cells(iRowPAP - 1, ColSAPAccount)
            End If
            
            'wsDis.Select
            For iRowDis = 2 To iMaxDis
                strKeyDis = wsDis.Cells(iRowDis, ColDisAccount) & "-" & wsDis.Cells(iRowDis, ColDisBranch)
                If strKeyDis = strKeyPAP Then
                    strProfitCenter = wsDis.Cells(iRowDis, ColDisProfitCenter)
                    strProduct = wsDis.Cells(iRowDis, ColDisProduct)
                    strCustomer = wsDis.Cells(iRowDis, ColDisCustomer)
                    wsPAP.Cells(iRowPAP, ColSAPProfitCenter) = strProfitCenter
                    wsPAP.Cells(iRowPAP, ColSAPProduct) = strProduct
                    wsPAP.Cells(iRowPAP, ColSAPCustomer) = strCustomer
                    Exit For
                End If
            Next iRowDis
            
            'Debug.Print iRowPAP & "   " & strKeyPAP
        End If
    End If
Next iRowPAP

End Sub

Sub Test()
Call SPS_Discount_Info("SPS")

End Sub

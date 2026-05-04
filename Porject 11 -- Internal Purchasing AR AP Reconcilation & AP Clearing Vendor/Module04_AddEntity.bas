Attribute VB_Name = "Module04_AddEntity"
Option Explicit

Sub Add_Entity_in_Bank_Statement(CompanyName As String)

'Dim CompanyName As String
'CompanyName = "MSD"

Dim wsCode As Worksheet
Dim sCodeEntity As String
Dim sCodeAccount As String
Dim sCodeDescription As String
Dim iMaxRowCode As Integer
Dim iCode As Integer

Dim wsBS As Worksheet
Dim iMaxRowBS As Integer
Dim iBS As Integer
Dim sBSAccount As String
Dim sBSDescription As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsBS = Worksheets("Bank Statement")
wsBS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBS = lRealLastRow

Set wsCode = Worksheets("Bank Code")
wsCode.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCode = lRealLastRow

'wsbs.Select
wsBS.Cells(1, ColBSEntity) = "Entity"
wsBS.Cells(1, ColBSAMTPAP) = "Amount PAP"
wsBS.Cells(1, ColBSTradingPart) = "Trading Part"
wsBS.Cells(1, ColBSCustomer) = "Customer ID"
wsBS.Cells(1, ColBSBranch) = "Branch"

For iCode = 2 To iMaxRowCode
    sCodeEntity = wsCode.Cells(iCode, ColBankCodeEntity)
    sCodeAccount = wsCode.Cells(iCode, ColBankCodeAccount)
    sCodeDescription = UCase(Replace(wsCode.Cells(iCode, ColBankCodeDescription), " ", ""))
    
    For iBS = 2 To iMaxRowBS
        sBSAccount = wsBS.Cells(iBS, ColBSAccount)
        sBSDescription = UCase(Replace(wsBS.Cells(iBS, ColBSDescription), " ", ""))
        'If sCodeAccount = sBSAccount And sCodeDescription = sBSDescription Then
        If sCodeAccount = sBSAccount And InStr(sBSDescription, sCodeDescription) > 0 Then
            wsBS.Cells(iBS, ColBSEntity) = sCodeEntity
            If sCodeEntity = CompanyName Then
                'Debug.Print iBS
                wsBS.Rows(iBS).Font.Color = -16776961
            End If
        End If
    Next iBS
Next iCode


wsBS.Select
End Sub

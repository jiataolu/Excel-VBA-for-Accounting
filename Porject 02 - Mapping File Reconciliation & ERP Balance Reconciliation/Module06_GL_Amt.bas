Attribute VB_Name = "Module06_GL_Amt"
Option Explicit

Sub Cash_Reconciliation_GL_Bank()
Call Read_Bank_GL_Template
Call Read_ERP_Data_into_Cash_Project
Call Read_FIS_Data_into_Cash_Porject
Call Write_Missing_ERP_FIS
Call Write_Difference
End Sub


Sub Read_Bank_GL_Template()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

'CP: Cash Project
Dim wsCP As Worksheet
Dim iMaxRowCP As Integer
Dim iNextRowCP As Integer

Dim wsGLBank As Worksheet
Dim iMaxRowGLBank As Long
Dim iMaxColGLBank As Integer
Dim i As Long
'Dim j As Integer

'clear Sheet "Cash Project"
Set wsCP = Worksheets("Cash Project")
wsCP.Select
Call DeleteUnusedFormats

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCP = lRealLastRow
If iMaxRowCP > 1 Then
    wsCP.Rows("2" & ":" & CStr(iMaxRowCP)).Delete
End If

iNextRowCP = 2

Set wsGLBank = Worksheets("GL-Bank")
wsGLBank.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowGLBank = lRealLastRow
iMaxColGLBank = lRealLastCol
If iMaxRowGLBank < 2 Then Exit Sub

For i = 2 To iMaxRowGLBank
        wsCP.Cells(iNextRowCP, iCPCat) = wsGLBank.Cells(i, iGLBankCat)
        wsCP.Cells(iNextRowCP, iCPBU) = wsGLBank.Cells(i, iGLBankBU)
        wsCP.Cells(iNextRowCP, iCPGL) = wsGLBank.Cells(i, iGLBankGL)
        wsCP.Cells(iNextRowCP, iCPBankCode) = wsGLBank.Cells(i, iGLBankBankCode)
        'wsCP.Cells(iNextRowCP, iCPKey) = wsCP.Cells(iNextRowCP, iCPBU) & "-" & wsCP.Cells(iNextRowCP, iCPGL)
        If Is_BU(wsCP.Cells(iNextRowCP, iCPBU)) Then
            wsCP.Cells(iNextRowCP, iCPKey) = wsCP.Cells(iNextRowCP, iCPBU) & "-" & wsCP.Cells(iNextRowCP, iCPGL)
            'There are some duplicate BU-Gl in sheet "GL-Bank" bottom, pleaes check with QiQI why such duplication
            If UCase(Left(wsCP.Cells(iNextRowCP, iCPCat), 2)) = "D-" Then
                wsCP.Cells(iNextRowCP, iCPKey) = wsCP.Cells(iNextRowCP, iCPKey) & "-D"
            End If
        End If
        iNextRowCP = iNextRowCP + 1
Next i

wsCP.Select
End Sub


Sub Read_ERP_Data_into_Cash_Project()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsERP As Worksheet
Dim iMaxRowERP As Integer
Dim iERP As Integer
Dim rngERP As Range
Dim rngFound As Range

Dim wsCP As Worksheet
Dim iMaxRowCP As Integer
Dim strKey As String
Dim rngCP As Range
Dim iCP As Integer

Dim iFound As Integer

Set wsERP = Worksheets("ERP")
wsERP.Select
Set rngERP = Columns(iERPKey)

Set wsCP = Worksheets("Cash Project")
wsCP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCP = lRealLastRow
'Debug.Print iMaxRowCP
If iMaxRowCP < 2 Then Exit Sub

For iCP = 2 To iMaxRowCP
    strKey = wsCP.Cells(iCP, iCPKey)
    Set rngFound = rngERP.Find(strKey, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        wsCP.Cells(iCP, iCPAmtERP) = rngFound.Cells(1, 2)
    End If
            
Next iCP

'to check if all data in ERP are read
'Set rngCP = Columns(iCPKey)
Set wsERP = Worksheets("ERP")
wsERP.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowERP = lRealLastRow
Columns(iERPCheck).Select
Selection.ClearContents
wsERP.Cells(1, iERPCheck) = "Is Read?"

'For iERP = 180 To 180
For iERP = 2 To iMaxRowERP
    'If wsERP.Cells(i, iERPAmt) <> 0 Then
        strKey = wsERP.Cells(iERP, iERPKey)
        'Debug.Print strKey
        iFound = 0
        
        For iCP = 2 To iMaxRowCP
            If strKey = wsCP.Cells(iCP, iCPKey) Then
                iFound = 1
                wsERP.Cells(iERP, iERPCheck) = wsERP.Cells(iERP, iERPCheck) + 1
            End If
        Next iCP
        If iFound = 0 Then wsERP.Cells(iERP, iERPCheck) = "Missing"
        
        'Set rngFound = rngCP.Find(strKey, LookIn:=xlValues, lookat:=xlWhole)
        'If rngFound Is Nothing Then wsERP.Cells(i, iERPCheck) = "Missing"
        'If Not rngFound Is Nothing Then wsERP.Cells(i, iERPCheck) = wsERP.Cells(i, iERPCheck) + 1
    'End If
Next iERP
End Sub


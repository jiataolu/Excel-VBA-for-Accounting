Attribute VB_Name = "Module02_Recon_Mapping"
Option Explicit

Sub Recon_Mapping()
Call Recon_Mapping_By_BUGL
Call Recon_Mapping_By_Bank_Code
End Sub
Sub Recon_Mapping_By_BUGL()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsGLBank As Worksheet
Dim iMaxRowGLBank As Integer
Dim iNextRowGLBank As Integer
Dim iRowGLBank As Integer
Dim strKeyGLBank As String

Dim wsMapping As Worksheet
Dim iMaxRowMapping As Integer
Dim iRowMapping As Integer
Dim strkeyMapping As String
Dim iKeyFound As Integer

Set wsGLBank = Worksheets("GL-Bank")
wsGLBank.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowGLBank = lRealLastRow
If iMaxRowGLBank < 2 Then Exit Sub
iNextRowGLBank = iMaxRowGLBank + 1

Set wsMapping = Worksheets("Mapping")
wsMapping.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMapping = lRealLastRow
If iMaxRowMapping < 2 Then Exit Sub

'Debug.Print iMaxRowMapping
Columns(iMappingCheckBUGL).Delete
wsMapping.Cells(1, iMappingCheckBUGL) = "Is in GL-Bank (by BU-GL)"
'For iRowMapping = 423 To 423
For iRowMapping = 2 To iMaxRowMapping
    If UCase(Cells(iRowMapping, iMappingGL)) = "MISSING" Then GoTo ContinueMapping
    strkeyMapping = Replace(Cells(iRowMapping, iMappingBU), " ", "") & "-" & Replace(Cells(iRowMapping, iMappingGL), " ", "")
    'Debug.Print strkeyMapping
    iKeyFound = 0
    
    For iRowGLBank = 2 To iNextRowGLBank - 1
        strKeyGLBank = Replace(wsGLBank.Cells(iRowGLBank, iGLBankBU), " ", "") & "-" & Replace(wsGLBank.Cells(iRowGLBank, iGLBankGL), " ", "")
        If strkeyMapping = strKeyGLBank Then
            iKeyFound = 1
            'Debug.Print iRowGLBank
            GoTo ContinueMapping
        End If
    Next iRowGLBank
    
    If iKeyFound = 0 Then
        'Debug.Print strkeyMapping
        Cells(iRowMapping, iMappingCheckBUGL) = "Missing"
    End If
ContinueMapping:
Next iRowMapping




End Sub

Sub Recon_Mapping_By_Bank_Code()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsGLBank As Worksheet
Dim iMaxRowGLBank As Integer
Dim iNextRowGLBank As Integer
Dim iRowGLBank As Integer
Dim strKeyGLBank As String

Dim wsMapping As Worksheet
Dim iMaxRowMapping As Integer
Dim iRowMapping As Integer
Dim strkeyMapping As String
Dim iKeyFound As Integer

Set wsGLBank = Worksheets("GL-Bank")
wsGLBank.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowGLBank = lRealLastRow
If iMaxRowGLBank < 2 Then Exit Sub
iNextRowGLBank = iMaxRowGLBank + 1

Set wsMapping = Worksheets("Mapping")
wsMapping.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMapping = lRealLastRow
If iMaxRowMapping < 2 Then Exit Sub

Columns(iMappingCheckBankCode).Delete
Cells(1, iMappingCheckBankCode) = "Is in GL-Bank (by bank code)"
'Debug.Print iMaxRowMapping
'For iRowMapping = 423 To 423
For iRowMapping = 2 To iMaxRowMapping
    strkeyMapping = Replace(Cells(iRowMapping, iMappingBankCode), " ", "")
    'Debug.Print strkeyMapping
    iKeyFound = 0
    
    For iRowGLBank = 2 To iNextRowGLBank - 1
        strKeyGLBank = Replace(wsGLBank.Cells(iRowGLBank, iGLBankBankCode), " ", "")
        If InStr(strKeyGLBank, strkeyMapping) > 0 Then
            iKeyFound = 1
            'Debug.Print iRowGLBank
            GoTo ContinueMapping
        End If
    Next iRowGLBank
    
    If iKeyFound = 0 Then
        Cells(iRowMapping, iMappingCheckBankCode) = "Missing"
    End If
ContinueMapping:
Next iRowMapping




End Sub


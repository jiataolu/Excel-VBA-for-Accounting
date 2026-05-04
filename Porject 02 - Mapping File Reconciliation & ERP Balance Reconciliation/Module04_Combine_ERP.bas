Attribute VB_Name = "Module04_Combine_ERP"
Option Explicit

Sub ERP_Data_Reading_3ERP()
Call ERP_Data_Reading_01_Initialization
Call ERP_NetSuite_04
Call ERP_JDE_03
Call ERP_FCCS_02
End Sub

Sub ERP_Add_or_Change_columns()



End Sub
Sub ERP_Data_Reading_01_Initialization()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsERP As Worksheet
Dim lMaxRowERP As Long
Dim lMaxColERP As Long
Dim rngERP As Range

Set wsERP = Worksheets("ERP")
wsERP.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lMaxRowERP = lRealLastRow
lMaxColERP = lRealLastCol
'Debug.Print lMaxRowERP

If lMaxRowERP < 2 Then Exit Sub
Set rngERP = Range(Cells(2, 1), Cells(lMaxRowERP, lMaxColERP))
rngERP.ClearContents

End Sub


Sub ERP_FCCS_02()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsERP As Worksheet
Dim lERPNextRow As Long

Dim wsFCCS As Worksheet
Dim lMaxRowFCCS As Long
Dim lMaxColFCCS As Long
Dim i As Long

Set wsERP = Worksheets("ERP")
wsERP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lERPNextRow = lRealLastRow + 1

Set wsFCCS = Worksheets("FCCS")
wsFCCS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lMaxRowFCCS = lRealLastRow
If lMaxRowFCCS < 2 Then Exit Sub

For i = 2 To lMaxRowFCCS
    wsERP.Cells(lERPNextRow, 1) = "FCCS"    'SAP Account
    wsERP.Cells(lERPNextRow, 2) = wsFCCS.Cells(i, 1)    'SAP Account
    wsERP.Cells(lERPNextRow, 3) = wsFCCS.Cells(i, 3)    'Company code
    wsERP.Cells(lERPNextRow, 4) = wsFCCS.Cells(i, 6)    'biz area
    'wsERP.Cells(lERPNextRow, 5) = "=C" & CStr(lERPNextRow) & " &  ""-"" &  B" & CStr(lERPNextRow)
    wsERP.Cells(lERPNextRow, 5) = wsERP.Cells(lERPNextRow, 3) & "-" & wsERP.Cells(lERPNextRow, 2)
    wsERP.Cells(lERPNextRow, 6) = wsFCCS.Cells(i, 4)    'biz area
    lERPNextRow = lERPNextRow + 1
Next i

wsERP.Select
End Sub


Sub ERP_JDE_03()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsERP As Worksheet
Dim lERPNextRow As Long

Dim wsJDE As Worksheet
Dim lMaxRowJDE As Long
Dim lMaxColJDE As Long
Dim lRowDataStart As Long
Dim strJDE As String
Dim i As Long
Dim j As Integer

Set wsERP = Worksheets("ERP")
wsERP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lERPNextRow = lRealLastRow + 1

Set wsJDE = Worksheets("JDE")
wsJDE.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lMaxRowJDE = lRealLastRow
lMaxColJDE = lRealLastCol

If lMaxRowJDE < 2 Then Exit Sub

lRowDataStart = 0
For i = 1 To lMaxRowJDE
    strJDE = ""
    For j = 1 To lMaxColJDE
        strJDE = strJDE & Cells(i, j)
    Next j
    strJDE = Replace(strJDE, " ", "")
    'Debug.Print strJDE
    If InStr(strJDE, "ObjectCompanySub") > 0 Then
        lRowDataStart = i + 1
        Exit For
    End If
Next i

'Debug.Print lRowDataStart

If lRowDataStart > 0 Then
    For i = lRowDataStart To lMaxRowJDE
        wsERP.Cells(lERPNextRow, 1) = "JDE"    'SAP Account
        wsERP.Cells(lERPNextRow, 2) = wsJDE.Cells(i, 1)    'SAP Account
        wsERP.Cells(lERPNextRow, 3) = wsJDE.Cells(i, 2)    'Company code
        wsERP.Cells(lERPNextRow, 4) = "MMS"                 'biz area
        'wsERP.Cells(lERPNextRow, 5) = "=C" & CStr(lERPNextRow) & " & ""-"" & B" & CStr(lERPNextRow)  'biz area
        wsERP.Cells(lERPNextRow, 5) = wsERP.Cells(lERPNextRow, 3) & "-" & wsERP.Cells(lERPNextRow, 2)
        wsERP.Cells(lERPNextRow, 6) = wsJDE.Cells(i, 7)    'biz area
        lERPNextRow = lERPNextRow + 1
    Next i
End If

wsERP.Select
End Sub


Sub ERP_NetSuite_04()

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsERP As Worksheet
Dim lERPNextRow As Long

Dim wsNetS As Worksheet
Dim lMaxRowNetS As Long
Dim lMaxColNetS As Long
Dim lRowDataStart As Long
Dim lRowDataEnd As Long
Dim strNetS As String
Dim i As Long
Dim j As Integer
Dim strGL As String
Dim strGLCurrent As String
Dim dAMTGL As Double
Dim dAMTGLCurrent As Double
Dim strText As String


Set wsERP = Worksheets("ERP")
wsERP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lERPNextRow = lRealLastRow + 1

Set wsNetS = Worksheets("NetSuite")
wsNetS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
lMaxRowNetS = lRealLastRow
lMaxColNetS = lRealLastCol



If lMaxRowNetS < 2 Then Exit Sub

lRowDataStart = 0
lRowDataEnd = 0
For i = 1 To lMaxRowNetS
    For j = 1 To lMaxColNetS
        If Cells(i, 1) = "Bank" Then lRowDataStart = i + 1
        If Cells(i, 1) = "Total Bank" Then lRowDataEnd = i - 1
    Next j
    If lRowDataStart * lRowDataEnd > 1 Then Exit For
Next i

'Debug.Print lRowDataStart
'Debug.Print lRowDataEnd

strGL = ""
dAMTGL = 0
For i = lRowDataStart To lRowDataEnd
    strText = Cells(i, 1)
    dAMTGLCurrent = Cells(i, 2)
    strGLCurrent = JDE_GL(strText)
    'Debug.Print i & " : " & strText & "-" & strGLCurrent
    If strGL <> strGLCurrent Then
        If strGL <> "" Then
            'Debug.Print lERPNextRow
            wsERP.Cells(lERPNextRow, 1) = "NetSuite"            'SAP Account
            wsERP.Cells(lERPNextRow, 2) = strGL                 'SAP Account
            wsERP.Cells(lERPNextRow, 3) = "7600"                  'Company code
            wsERP.Cells(lERPNextRow, 4) = "CMM"                 'biz area
            'wsERP.Cells(lERPNextRow, 5) = "=C" & CStr(lERPNextRow) & " & ""-"" & B" & CStr(lERPNextRow)  'biz area
            wsERP.Cells(lERPNextRow, 5) = wsERP.Cells(lERPNextRow, 3) & "-" & wsERP.Cells(lERPNextRow, 2)
            wsERP.Cells(lERPNextRow, 6) = dAMTGL                'Amount
            lERPNextRow = lERPNextRow + 1
        End If
        
        strGL = strGLCurrent
        dAMTGL = dAMTGLCurrent
    End If
Next i
wsERP.Cells(lERPNextRow, 1) = "NetSuite"            'SAP Account
wsERP.Cells(lERPNextRow, 2) = strGLCurrent                 'SAP Account
wsERP.Cells(lERPNextRow, 3) = "7600"                  'Company code
wsERP.Cells(lERPNextRow, 4) = "CMM"                 'biz area
'wsERP.Cells(lERPNextRow, 5) = "=C" & CStr(lERPNextRow) & " & B" & CStr(lERPNextRow)  'biz area
wsERP.Cells(lERPNextRow, 5) = wsERP.Cells(lERPNextRow, 3) & "-" & wsERP.Cells(lERPNextRow, 2)
wsERP.Cells(lERPNextRow, 6) = dAMTGLCurrent                'Amount

wsERP.Select

End Sub


Attribute VB_Name = "Module08_ReconcileSAP"
Option Explicit


Sub Reconcile_PAP_invoices(CompanyName As String)

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Long
Dim jMaxColSAP As Long
Dim iRowSAP As Integer
Dim jColSAP As Integer

Dim wsPAP As Worksheet
Dim iMaxRowPAP As Integer
Dim jMaxColPAP As Integer
Dim iRowPAP As Long
Dim rngPAP As Range
Dim sKey1 As String
Dim sKey2 As String
Dim dTotalAmt As Double
Dim dTotalDis As Double
Dim sAss As String
Dim sAssDash As String
Dim sCustomerID As String
Dim sTradingPartner As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsToRunFBL5N As Worksheet
Dim rngCustID As Range
Dim rngFound As Range
Dim sExtraTrading As String

Set wsToRunFBL5N = Worksheets("To Run FBL5N")
wsToRunFBL5N.Select
Set rngCustID = Columns(ColToRunCustID)

Set wsPAP = Worksheets("PAP Invoices")
wsPAP.Select
Cells.Select
Selection.Delete

Set wsSAP = Worksheets("FBL5N")
wsSAP.Select
Call DeleteUnusedFormats
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow
jMaxColSAP = lRealLastCol

If iMaxRowSAP < 2 Then Exit Sub

iRowPAP = 1
For iRowSAP = 1 To 1
    For jColSAP = 1 To jMaxColSAP
        wsPAP.Cells(iRowPAP, jColSAP) = wsSAP.Cells(iRowSAP, jColSAP)
    Next jColSAP
Next iRowSAP

iRowPAP = 2
For iRowSAP = 2 To iMaxRowSAP
    If wsSAP.Cells(iRowSAP, ColSAPAssignment) <> "" Then
        For jColSAP = 1 To jMaxColSAP
            wsPAP.Cells(iRowPAP, jColSAP) = wsSAP.Cells(iRowSAP, jColSAP)
        Next jColSAP
        If wsSAP.Cells(iRowSAP, ColSAPAccount) = "7006153" Then
            wsPAP.Cells(iRowPAP, ColSAPTradingPart) = "'8232"
        End If
        'If wsSAP.Cells(iRowSAP, ColSAPAccount) = "7006153" Then Debug.Print iRowSAP
        iRowPAP = iRowPAP + 1
    End If
Next iRowSAP


'Sort PAP Sheet
wsPAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowPAP = lRealLastRow
jMaxColPAP = lRealLastCol
If iMaxRowPAP < 2 Then Exit Sub

Set rngPAP = Range(Cells(1, 1), Cells(iMaxRowPAP, jMaxColPAP))
rngPAP.Select
With wsPAP.Sort
    .SortFields.Clear
    .SortFields.Add Key:=rngPAP.Columns(ColSAPTradingPart), SortOn:=xlSortOnValues, Order:=xlAscending
    .SortFields.Add Key:=rngPAP.Columns(ColSAPAccount), SortOn:=xlSortOnValues, Order:=xlDescending
    .SortFields.Add Key:=rngPAP.Columns(ColSAPBranch), SortOn:=xlSortOnValues, Order:=xlDescending
    .SortFields.Add Key:=rngPAP.Columns(ColSAPAssignment), SortOn:=xlSortOnValues, Order:=xlDescending
    .SetRange rngPAP
    .Header = xlYes
    .Apply
End With
Cells.Select
Cells.EntireColumn.AutoFit

'Insert empty lines to seperate by trading partner and customer ID
wsPAP.Cells(iMaxRowPAP + 1, 1) = "Total"

sKey1 = wsPAP.Cells(iMaxRowPAP, ColSAPAccount) & wsPAP.Cells(iMaxRowPAP, ColSAPBranch) & wsPAP.Cells(iMaxRowPAP, ColSAPTradingPart)
For iRowPAP = iMaxRowPAP - 1 To 2 Step -1
    sKey2 = wsPAP.Cells(iRowPAP, ColSAPAccount) & wsPAP.Cells(iRowPAP, ColSAPBranch) & wsPAP.Cells(iRowPAP, ColSAPTradingPart)

    If sKey1 <> sKey2 Then
        sKey1 = sKey2
        'Rows(iRowPAP).Select
        Rows(iRowPAP + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        wsPAP.Cells(iRowPAP + 1, 1) = "Total"
    End If
Next iRowPAP

'Find total amount
wsPAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowPAP = lRealLastRow
jMaxColPAP = lRealLastCol

dTotalAmt = 0
dTotalDis = 0
For iRowPAP = 2 To iMaxRowPAP
    If wsPAP.Cells(iRowPAP, 1) = "Total" Then
        wsPAP.Cells(iRowPAP, ColSAPAmt) = dTotalAmt
        wsPAP.Cells(iRowPAP, ColSAPDis) = dTotalDis
        wsPAP.Cells(iRowPAP, ColSAPNetAmt) = dTotalAmt - dTotalDis
        
        dTotalAmt = 0
        dTotalDis = 0
        Range(Cells(iRowPAP, 1), Cells(iRowPAP, jMaxColPAP)).Select
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
    Else
        dTotalAmt = dTotalAmt + wsPAP.Cells(iRowPAP, ColSAPAmt)
        dTotalDis = dTotalDis + wsPAP.Cells(iRowPAP, ColSAPDis)
        
        Range(Cells(iRowPAP, 1), Cells(iRowPAP, jMaxColPAP)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End If
    
    
Next iRowPAP

'To check and add extra trading partner
For iRowPAP = 2 To iMaxRowPAP
    sCustomerID = Cells(iRowPAP, ColSAPCustID)
    
    Set rngFound = rngCustID.Find(sCustomerID, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        sExtraTrading = rngFound.Cells(1, 2)
        If sExtraTrading <> "" Then
            sTradingPartner = Cells(iRowPAP, ColSAPTradingPart)
            sTradingPartner = sTradingPartner & " (" & sExtraTrading & ")"
            Cells(iRowPAP, ColSAPTradingPart) = sTradingPartner
        End If
    End If
Next iRowPAP

If CompanyName = "Well.ca" Then
    wsPAP.Select
    lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
    lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
    iMaxRowPAP = lRealLastRow
    
    If iMaxRowPAP > 1 Then
        For iRowPAP = 2 To iMaxRowPAP
            If UCase(wsPAP.Cells(iRowPAP, 1) <> "TOTAL") Then
                sAss = wsPAP.Cells(iRowPAP, ColSAPAssignment)
                If Len(sAss) > 2 Then
                    sAssDash = Left(sAss, 3) & "-" & Right(sAss, Len(sAss) - 3)
                Else
                    sAssDash = sAss
                End If
                wsPAP.Cells(iRowPAP, ColSAPAssignment) = sAssDash
            End If
        Next iRowPAP
    End If
End If

wsPAP.Select
Cells(1, 1).Select

End Sub

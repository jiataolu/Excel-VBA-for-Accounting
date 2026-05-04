Attribute VB_Name = "Module13_Mapping_FX"
Option Explicit


Sub Find_Mapping_Info_FX_Step2_Process_FX_Coding()

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sCRY As String

Dim sFXBU As String
Dim sFXGL As String
Dim sFXVendor As String
Dim sFXProfitC As String
Dim sFXKeyCode As String
Dim sFXAssInfo As String
Dim sFXCostCenter As String

Dim sBU As String
Dim sGL As String
Dim sVendor As String
Dim sProfitC As String
Dim sKeyCode As String
Dim sAssInfo As String
Dim sCostCenter As String
Dim iFoundFX As Integer


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsItems = Worksheets("2-Items to post")
wsItems.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub


'To check if there is FX transactions
iFoundFX = 0
For iRowItems = 2 To iMaxRowItems
    
    sCRY = wsItems.Cells(iRowItems, iColItemsPostCurrency)
    sCRY = UCase(Replace(sCRY, " ", ""))
    
    If sCRY <> "" Then
        iFoundFX = 1
        'Call Initialize_Items_Sheet_FX
        Exit For
    End If

Next iRowItems

If iFoundFX = 0 Then GoTo FINISHFX

'if there is fx, then process FX coding
For iRowItems = 2 To iMaxRowItems
'For iRowItems = 2 To 2
    
    sCRY = wsItems.Cells(iRowItems, iColItemsPostCurrency)
    sCRY = UCase(Replace(sCRY, " ", ""))
    
    
    If sCRY = "" Then GoTo CONTINUEWITHNEXTITEM
    
    'FX-BU value
    sBU = wsItems.Cells(iRowItems, iColItemsPostBU)
    sFXBU = sBU
    wsItems.Cells(iRowItems, iColItemsFXBU) = sFXBU
    
    'FX-GL value
    sGL = wsItems.Cells(iRowItems, iColItemsPostGL)
    sFXGL = sGL
    wsItems.Cells(iRowItems, iColItemsFXGL) = sFXGL
    
    'FX-Vendor value
    sVendor = wsItems.Cells(iRowItems, iColItemsPostVendor)
    sFXVendor = sVendor
    wsItems.Cells(iRowItems, iColItemsFXVendorCode) = sFXVendor
    
    'FX-Profit Center value
    sProfitC = wsItems.Cells(iRowItems, iColItemsPostProfitC)
    sFXProfitC = sProfitC
    wsItems.Cells(iRowItems, iColItemsFXProfitC) = sFXProfitC
    
    'FX-KeyCode value
    sKeyCode = wsItems.Cells(iRowItems, iColItemsPostKeyCode)
    sFXKeyCode = sKeyCode
    wsItems.Cells(iRowItems, iColItemsFXKeyCode) = sFXKeyCode
    
    'FX-Assignment value
    sAssInfo = wsItems.Cells(iRowItems, iColItemsPostAssInfo)
    sFXAssInfo = sAssInfo
    wsItems.Cells(iRowItems, iColItemsFXAssInfo) = sFXAssInfo
    
    'FX-CostCenter value
    sCostCenter = wsItems.Cells(iRowItems, iColItemsPostCostCenter)
    sFXCostCenter = sCostCenter
    wsItems.Cells(iRowItems, iColItemsFXCostCenter) = sFXCostCenter
    
    'Switch value of BU and GL
    wsItems.Cells(iRowItems, iColItemsPostBU) = MainCompanyCode
    wsItems.Cells(iRowItems, iColItemsPostGL) = MainGLFX

    'set vendor code for main JE as zero
    wsItems.Cells(iRowItems, iColItemsPostVendor) = ""
    
    'highlight FX-Amount cell
    wsItems.Cells(iRowItems, iColItemsFXAmt).Interior.Color = vbYellow
    
    
CONTINUEWITHNEXTITEM:
Next iRowItems

Cells.EntireColumn.AutoFit

FINISHFX:
End Sub

Sub Find_Mapping_Info_FX_Step1_Initialize_Items_Sheet_FX()
Dim wsItems As Worksheet
Dim lRealLastRow As Long
Dim lRealLastCol As Long
Dim iMaxRowItems As Integer

Set wsItems = Worksheets("2-Items to post")
wsItems.Select

wsItems.Cells(1, iColItemsPostCurrency) = "Currency"
wsItems.Cells(1, iColItemsFXAmt) = "FX-Amt"
wsItems.Cells(1, iColItemsFXBU) = "FX-Bu"
wsItems.Cells(1, iColItemsFXGL) = "FX-Gl"
wsItems.Cells(1, iColItemsFXVendorCode) = "FX-Vendor"
wsItems.Cells(1, iColItemsFXProfitC) = "FX-ProfitC"
wsItems.Cells(1, iColItemsFXKeyCode) = "FX-KeyCode"
wsItems.Cells(1, iColItemsFXAssInfo) = "FX-Assignment"
wsItems.Cells(1, iColItemsFXCostCenter) = "FX-CostCenter"


Columns(iColItemsPostCurrency).HorizontalAlignment = xlCenter
Columns(iColItemsFXAmt).HorizontalAlignment = xlCenter
Columns(iColItemsFXBU).HorizontalAlignment = xlCenter
Columns(iColItemsFXGL).HorizontalAlignment = xlCenter
Columns(iColItemsFXVendorCode).HorizontalAlignment = xlCenter
Columns(iColItemsFXProfitC).HorizontalAlignment = xlCenter
Columns(iColItemsFXKeyCode).HorizontalAlignment = xlCenter
Columns(iColItemsFXAssInfo).HorizontalAlignment = xlCenter
Columns(iColItemsFXCostCenter).HorizontalAlignment = xlCenter






lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow

With Range(Cells(1, iColItemsPostCurrency), Cells(iMaxRowItems, iColItemsFXCostCenter)).Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = 0.799981688894314
    .PatternTintAndShade = 0
End With


Set wsItems = Nothing




End Sub





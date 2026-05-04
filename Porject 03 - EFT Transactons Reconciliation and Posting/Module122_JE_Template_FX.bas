Attribute VB_Name = "Module122_JE_Template_FX"
Option Explicit

Sub Fill_JE_Template_FX()

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sCRY As String

Dim sFXAmount As String
Dim sFXBU As String
Dim sFXGL As String
Dim sFXVendor As String
Dim sFXProfitC As String
Dim sFXKeyCode As String
Dim sFXAssInfo As String
Dim sFXCostCenter As String
Dim sKeyCode As String

Dim iFoundFX As Integer

Dim sDate As String
Dim sLineText As String

Dim sJESheetName As String
Dim wsJE As Worksheet
Dim iMaxJERow As Integer


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

If iFoundFX = 0 Then
    GoTo FINISHFX
Else
    Debug.Print "FX"
    'GoTo FINISHFX
End If


sJESheetName = "3 - C-SAP Standard Template"
sDate = Format(wsItems.Cells(2, iColItemsPostingDate), "MM/DD/YYYY")
sLineText = "EFT " & sDate

Set wsJE = Worksheets(sJESheetName)

'if there is fx, then process FX coding
For iRowItems = 2 To iMaxRowItems
'For iRowItems = 2 To 2
    
    sCRY = wsItems.Cells(iRowItems, iColItemsPostCurrency)
    sCRY = UCase(Replace(sCRY, " ", ""))
    
    
    If sCRY = "" Then GoTo CONTINUEWITHNEXTITEM
    
    'FX-Amount value
    sFXAmount = CStr(Abs(wsItems.Cells(iRowItems, iColItemsFXAmt)))
    
    'FX-BU value
    sFXBU = wsItems.Cells(iRowItems, iColItemsFXBU)
    
    'FX-GL value
    sFXGL = wsItems.Cells(iRowItems, iColItemsFXGL)
    
    'FX-Vendor value
    sFXVendor = wsItems.Cells(iRowItems, iColItemsFXVendorCode)
    
    'FX-Profit Center value
    sFXProfitC = wsItems.Cells(iRowItems, iColItemsFXProfitC)
    
    'FX-KeyCode value
    sFXKeyCode = wsItems.Cells(iRowItems, iColItemsFXKeyCode)
    
    'FX-Assignment value
    sFXAssInfo = wsItems.Cells(iRowItems, iColItemsFXAssInfo)
    
    'FX-CostCenter value
    sFXCostCenter = wsItems.Cells(iRowItems, iColItemsFXCostCenter)
    
    
    'Write Header of JE, based on currency
    wsJE.Select
    lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
    lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
    iMaxJERow = lRealLastRow

    Call JE_Header_Info(JESheetName:=sJESheetName, _
                        JECompany:="9000", _
                        PostingDate:=sDate, _
                        DocDate:=sDate, _
                        DocType:="SA", _
                        JECur:=sCRY, _
                        DocHeaderText:=sLineText, _
                        SheetRow:=iMaxJERow + 2)
    
    If sFXKeyCode = "40" Or sFXKeyCode = "21" Then sKeyCode = "50"
    
    If sFXKeyCode = "50" Or sFXKeyCode = "31" Then sKeyCode = "40"
    
    'Write each line for FX account - 67023
    Call JE_L_Line(JESheetName:=sJESheetName, _
                    PostingKey:=sKeyCode, _
                    GLAccount:=MainGLFX, _
                    NewCompanyCode:="9000", _
                    DocumentCurrencyAmount:=sFXAmount, _
                    ProfitCenter:="", _
                    Assignment:=sFXAssInfo, _
                    LineText:=sLineText, _
                    FontBold:=False)
                    
    'Write each line for real offset account
    Call JE_L_Line(JESheetName:=sJESheetName, _
                    PostingKey:=sFXKeyCode, _
                    GLAccount:=sFXGL, _
                    Vendor:=sFXVendor, _
                    NewCompanyCode:=sFXBU, _
                    DocumentCurrencyAmount:=sFXAmount, _
                    ProfitCenter:="", _
                    Assignment:=sFXAssInfo, _
                    LineText:=sLineText, _
                    FontBold:=False)
                    
    
    wsItems.Select
                    
wsItems.Select
    
    
    
CONTINUEWITHNEXTITEM:
Next iRowItems

Cells.EntireColumn.AutoFit


Worksheets(sJESheetName).Select
Worksheets(sJESheetName).Columns(12).NumberFormat = "General"
Worksheets(sJESheetName).Columns(13).NumberFormat = "General"
Worksheets(sJESheetName).Columns(14).NumberFormat = "General"
Worksheets(sJESheetName).Columns(16).NumberFormat = "General"
Worksheets(sJESheetName).Columns(21).NumberFormat = "General"
Worksheets(sJESheetName).Columns(19).Style = "Comma"


FINISHFX:
End Sub

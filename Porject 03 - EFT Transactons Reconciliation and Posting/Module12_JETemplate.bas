Attribute VB_Name = "Module12_JETemplate"
Option Explicit

Sub Fill_JE_Template()
Application.ScreenUpdating = False
Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer

Dim sPostBU As String
Dim sPostGL As String
Dim sPostVendor As String
Dim sPostProfitCenter As String
Dim sPostingKey As String
Dim sAMT As String
Dim sABSAMT As String
Dim sOriGL As String
Dim sDate As String
Dim sLineText As String
Dim sPostAssInfo As String
Dim sPostCostCenter As String

Dim wsTempSheet As Worksheet
Dim iMaxRowTemp As Integer
Dim iRowTemp As Integer

Dim sJESheetName As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

sJESheetName = "3 - C-SAP Standard Template"

Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow

If iMaxRowItems < 2 Then Exit Sub


Call JE_Clean(sJESheetName)
'Call Check_Create_Temp_Sheet
wsItems.Select

sDate = Format(wsItems.Cells(2, iColItemsPostingDate), "MM/DD/YYYY")
sLineText = "EFT " & sDate

Call JE_Header_Info(JESheetName:=sJESheetName, _
                    JECompany:="9000", _
                    PostingDate:=sDate, _
                    DocDate:=sDate, _
                    DocType:="SA", _
                    JECur:="USD", _
                    DocHeaderText:=sLineText)

wsItems.Select

'Write line for Concentration account
For iRowItems = 2 To iMaxRowItems
'For iRowItems = 10 To 10

    'sOriGL = wsItems.Cells(iRowItems, iColItemsGL)
    
    sPostBU = wsItems.Cells(iRowItems, iColItemsPostBU)
    'If this row has been coded by previous line  in same sheet, then don't do anything
    If InStr(UCase(sPostBU), "SEE") > 0 Then GoTo DONEXTLINE_0
    
    
    sPostBU = "9000"
    sPostGL = wsItems.Cells(iRowItems, iColItemsGL)
    sPostProfitCenter = "9510"
    sPostAssInfo = "WIRE TRANSFER"
    sAMT = CStr(wsItems.Cells(iRowItems, iColItemsAMT))
    sABSAMT = CStr(Abs(wsItems.Cells(iRowItems, iColItemsAMT)))

    
    
    'if bank amount is negative, then debit concentration account
    'if bank amount is positive, then credit concentration account
    If sAMT > 0 Then
        sPostingKey = "50"
    Else
        sPostingKey = "40"
    End If
    
    'Write each line for concentration account
    Call JE_L_Line(JESheetName:=sJESheetName, _
                    PostingKey:=sPostingKey, _
                    GLAccount:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sPostProfitCenter, _
                    Assignment:=sPostAssInfo, _
                    LineText:=sLineText, _
                    FontBold:=True)
                    
    
    wsItems.Select
    
    
DONEXTLINE_0:
Next iRowItems




'write line for each offset account
For iRowItems = 2 To iMaxRowItems
'For iRowItems = 2 To 2

    sOriGL = wsItems.Cells(iRowItems, iColItemsGL)
    
    sPostBU = wsItems.Cells(iRowItems, iColItemsPostBU)
    sPostGL = wsItems.Cells(iRowItems, iColItemsPostGL)
    sPostVendor = wsItems.Cells(iRowItems, iColItemsPostVendor)
    sPostProfitCenter = wsItems.Cells(iRowItems, iColItemsPostProfitC)
    sPostingKey = wsItems.Cells(iRowItems, iColItemsPostKeyCode)
    sPostAssInfo = wsItems.Cells(iRowItems, iColItemsPostAssInfo)
    sAMT = CStr(wsItems.Cells(iRowItems, iColItemsAMT))
    sABSAMT = CStr(Abs(wsItems.Cells(iRowItems, iColItemsAMT)))
    sPostCostCenter = wsItems.Cells(iRowItems, iColItemsPostCostCenter)
    
    'If this row has been coded by previous line  in same sheet, then don't do anything
    If InStr(UCase(sPostBU), "SEE") > 0 Then GoTo DONEXTLINE
    
    
    'if this line has coding inforamtion, then write into JE template
    If Replace(sPostBU, " ", "") <> "" And InStr(sPostBU, WaitToConfirmInfo) = 0 Then
        Call JE_L_Line(JESheetName:=sJESheetName, _
                    PostingKey:=sPostingKey, _
                    GLAccount:=sPostGL, _
                    Vendor:=sPostVendor, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sPostProfitCenter, _
                    Assignment:=sPostAssInfo, _
                    LineText:=sLineText, _
                    CostCenter:=sPostCostCenter)
                    
    'if this line has no coding informaion, then use BU-9000 GL-10890 to code
    Else
        sPostBU = "9000"
        sPostGL = SuspenseAccount
        sPostProfitCenter = "9510"
        sPostVendor = ""
        
        If CDbl(sAMT) < 0 Then
            sPostingKey = "50"
        Else
            sPostingKey = "40"
        End If
        
        Call JE_L_Line(JESheetName:=sJESheetName, _
                        PostingKey:=sPostingKey, _
                        GLAccount:=sPostGL, _
                        Vendor:=sPostVendor, _
                        NewCompanyCode:=sPostBU, _
                        DocumentCurrencyAmount:=sABSAMT, _
                        ProfitCenter:=sPostProfitCenter, _
                        Assignment:=sPostAssInfo, _
                        LineText:=sLineText, _
                        CostCenter:=sPostCostCenter, _
                        FontColorIndex:=3)
                                
    End If
    

    
    wsItems.Select
    
    
DONEXTLINE:
Next iRowItems



Worksheets(sJESheetName).Select
Worksheets(sJESheetName).Columns(19).Style = "Comma"

Set wsItems = Nothing

Application.ScreenUpdating = True
End Sub







Sub JE_Clean(JESheetName As String)

'Row 1 to Row 4 are Header information. cleaning is from Row 5
Dim wsJE As Worksheet
Dim iMaxRowJE As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsJE = Worksheets(JESheetName)
wsJE.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowJE = lRealLastRow
If iMaxRowJE < 5 Then Exit Sub

wsJE.Rows("5" & ":" & iMaxRowJE).Delete

Set wsJE = Nothing

End Sub


Sub JE_Header_Info(JESheetName As String, _
                    Optional ByVal JECompany As String = "9000", _
                    Optional ByVal PostingDate As String = "01/01/1900", _
                    Optional ByVal DocDate As String = "01/01/1900", _
                    Optional ByVal DocType As String = "SA", _
                    Optional ByVal JECur As String = "USD", _
                    Optional ByVal DocHeaderText As String = "", _
                    Optional ByVal JEType As String = "Bank OPEX transactions (fees & interest, credit cards)", _
                    Optional ByVal SheetRow As Integer = 4)

'Header information is on Row 4
Dim wsJE As Worksheet

Set wsJE = Worksheets(JESheetName)
wsJE.Cells(SheetRow, 1) = "H"
wsJE.Cells(SheetRow, 2) = JECompany
wsJE.Cells(SheetRow, 3) = PostingDate
wsJE.Cells(SheetRow, 4) = DocDate
wsJE.Cells(SheetRow, 5) = DocType
wsJE.Cells(SheetRow, 6) = JECur
wsJE.Cells(SheetRow, 7) = DocHeaderText
wsJE.Cells(SheetRow, 8) = JEType

wsJE.Rows(19).Style = "Comma"

Set wsJE = Nothing

End Sub

'This funciton is up to AA-"Tax Jurisdiction"
Sub JE_L_Line(Optional ByVal JESheetName As String = "", _
                Optional ByVal PostingKey As String = "", _
                Optional ByVal GLAccount As String = "", _
                Optional ByVal Vendor As String = "", _
                Optional ByVal Customer As String = "", _
                Optional ByVal NewCompanyCode As String = "", _
                Optional ByVal AssetNumber As String = "", _
                Optional ByVal AssetTransactionType As String = "", _
                Optional ByVal DocumentCurrencyAmount As String = "", _
                Optional ByVal SpecialGLIndicator As String = "", _
                Optional ByVal ProfitCenter As String = "", _
                Optional ByVal CostCenter As String = "", _
                Optional ByVal Assignment As String = "", _
                Optional ByVal WBSElement As String = "", _
                Optional ByVal LineText As String = "", _
                Optional ByVal TaxCode As String = "", _
                Optional ByVal TaxJurisdiction As String = "", _
                Optional ByVal FontColorIndex As Integer = "1", _
                Optional ByVal FontBold As Boolean = False)


Dim wsJE As Worksheet
Dim iMaxRowJE As Integer
Dim iRowJE As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsJE = Worksheets(JESheetName)
wsJE.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowJE = lRealLastRow
If iMaxRowJE < 4 Then
    MsgBox "Please check JE Template."
    Exit Sub
End If

iRowJE = iMaxRowJE + 1

wsJE.Cells(iRowJE, 1) = "L"
wsJE.Cells(iRowJE, 12) = PostingKey
wsJE.Cells(iRowJE, 13) = GLAccount
wsJE.Cells(iRowJE, 14) = Vendor
wsJE.Cells(iRowJE, 15) = Customer
wsJE.Cells(iRowJE, 16) = NewCompanyCode
wsJE.Cells(iRowJE, 17) = AssetNumber
wsJE.Cells(iRowJE, 18) = AssetTransactionType
wsJE.Cells(iRowJE, 19) = DocumentCurrencyAmount
wsJE.Cells(iRowJE, 20) = SpecialGLIndicator
wsJE.Cells(iRowJE, 21) = ProfitCenter
wsJE.Cells(iRowJE, 22) = CostCenter
wsJE.Cells(iRowJE, 23) = Assignment
wsJE.Cells(iRowJE, 24) = WBSElement
wsJE.Cells(iRowJE, 25) = LineText
wsJE.Cells(iRowJE, 26) = TaxCode
wsJE.Cells(iRowJE, 27) = TaxJurisdiction

wsJE.Rows(iRowJE).font.ColorIndex = FontColorIndex
wsJE.Rows(iRowJE).font.Bold = FontBold
wsJE.Rows(iRowJE).NumberFormat = "General"

Set wsJE = Nothing
End Sub

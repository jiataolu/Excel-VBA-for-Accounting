Attribute VB_Name = "Module06_JETemplate"
Option Explicit

Sub Fill_JE_Template()
Application.ScreenUpdating = False
Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer

Dim sPostBU As String

Dim sPostGL As String

'use two variable to differentiate debit / credit side of one transfer
Dim sPostGL1 As String  'add afterward to seperate regular GL and vendor code
Dim sPostGL2 As String  'add afterward to seperate regular GL and vendor code
Dim sCurrency As String


Dim sPostVendor As String
Dim sPostProfitCenter As String
Dim sPostingKey1 As String
Dim sPostingKey2 As String
Dim dAMT As Double
Dim sABSAMT As String
Dim sOriGL As String
Dim sDate As String
Dim sLineText As String
Dim sPostAssInfo As String
Dim sPostCostCenter As String
Dim sBankvsBank As String
Dim sAss As String
Dim sProfitC As String
Dim sBankCode As String


Dim wsTempSheet As Worksheet
Dim iMaxRowTemp As Integer
Dim iRowTemp As Integer

Dim sJESheetName As String

Dim sCompanyCodeHeaderCADGL As String
Dim iFoundCADGL As Integer
Dim sCompanyCodeHeaderCADVENDOR As String
Dim iFoundCADVendor As Integer
Dim sCompanyCodeHeaderUSDGL As String
Dim iFoundUSDGL As Integer
Dim sCompanyCodeHeaderUSDVENDOR As String
Dim iFoundUSDVendor As Integer

Dim iFound As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim sSheetNametoUse As String

Dim wkbMap As Workbook
Dim wsMap As Worksheet
Dim rngMapBankCode As Range
Dim rngFound As Range

'open Mapping file
Application.DisplayAlerts = False
Set wkbMap = Workbooks.Open(Map_File_Full_Name, UpdateLinks:=False)
wkbMap.Activate
Set wsMap = Worksheets("Mapping Consolidated")
wsMap.Select
Set rngMapBankCode = Columns(SheetMapColBankCode)


'sJESheetName = "3 - C-SAP Standard Template"
ThisWorkbook.Activate
Set wsItems = Worksheets(Sheet04Name_Pivot)
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow

If iMaxRowItems < 2 Then Exit Sub


Call JE_Clean(Sheet05Name_JEUploadCAD)
Call JE_Clean(Sheet05Name_JEUploadUSD)

'Call Check_Create_Temp_Sheet
wsItems.Select

sDate = Format(Month_End_Date(month(Worksheets(Sheet02Name_JEData).Cells(2, 1)), year(Worksheets(Sheet02Name_JEData).Cells(2, 1))), "MM/DD/YYYY")
'sDate = Format(wsItems.Cells(2, iColItemsPostingDate), "MM/DD/YYYY")
sLineText = "ZBA " & Format(sDate, "MMM YYYY")

'sCompanyCodeHeader = Worksheets(Sheet04Name_Pivot).Cells(2, 1)
'sCompanyCodeHeaderforVendor = ""
'iFound = 0

wsItems.Select

'To find CAD-GL, CAD-Vendor, USD-GL, USD-Vendor
sCompanyCodeHeaderCADGL = ""
iFoundCADGL = 0
sCompanyCodeHeaderCADVENDOR = ""
iFoundCADVendor = 0
sCompanyCodeHeaderUSDGL = ""
iFoundUSDGL = 0
sCompanyCodeHeaderUSDVENDOR = ""
iFoundUSDVendor = 0


For iRowItems = 2 To iMaxRowItems - 1
    sPostGL1 = Replace(wsItems.Cells(iRowItems, 3), " ", "")
    sPostGL2 = Replace(wsItems.Cells(iRowItems, 6), " ", "")
    sCurrency = UCase(Replace(wsItems.Cells(iRowItems, 7), " ", ""))
    
    If sCurrency = "CAD" Then
        If Validate_GL(sPostGL1) And Validate_GL(sPostGL2) And iFoundCADGL = 0 Then
            sCompanyCodeHeaderCADGL = wsItems.Cells(iRowItems, 1)
            iFoundCADGL = 1
        End If
        
        If Not (Validate_GL(sPostGL1) And Validate_GL(sPostGL2)) And iFoundCADVendor = 0 Then
            sCompanyCodeHeaderCADVENDOR = wsItems.Cells(iRowItems, 1)
            iFoundCADVendor = 1
        End If
        
    ElseIf sCurrency = "USD" Then
        If Validate_GL(sPostGL1) And Validate_GL(sPostGL2) And iFoundUSDGL = 0 Then
            sCompanyCodeHeaderUSDGL = wsItems.Cells(iRowItems, 1)
            iFoundUSDGL = 1
        End If
        
        If Not (Validate_GL(sPostGL1) And Validate_GL(sPostGL2)) And iFoundUSDVendor = 0 Then
            sCompanyCodeHeaderUSDVENDOR = wsItems.Cells(iRowItems, 1)
            iFoundUSDVendor = 1
        End If
    
    End If

Next iRowItems


'Process regular GL posting
If iFoundCADGL Then Call JE_Header_Info(Sheet05Name_JEUploadCAD, _
                                        sCompanyCodeHeaderCADGL, _
                                        sDate, _
                                        sDate, _
                                        sLineText, _
                                        JEUpLoadDocType, _
                                        "CAD", _
                                        JEUpLoadJEType)
If iFoundUSDGL Then Call JE_Header_Info(Sheet05Name_JEUploadUSD, _
                                        sCompanyCodeHeaderUSDGL, _
                                        sDate, _
                                        sDate, _
                                        sLineText, _
                                        JEUpLoadDocType, _
                                        "USD", _
                                        JEUpLoadJEType)

For iRowItems = 2 To iMaxRowItems - 1
'For iRowItems = 2 To 2

    sPostGL1 = Replace(wsItems.Cells(iRowItems, 3), " ", "")
    sPostGL2 = Replace(wsItems.Cells(iRowItems, 6), " ", "")
    
    'If one bank account use vendor code, not regular GL, then skip it.
    If Not (Validate_GL(sPostGL1) And Validate_GL(sPostGL2)) Then GoTo NEXTITEMLOOP1
    
    sSheetNametoUse = ""
    sCurrency = UCase(Replace(wsItems.Cells(iRowItems, 7), " ", ""))
    If UCase(sCurrency) = "CAD" Then sSheetNametoUse = Sheet05Name_JEUploadCAD
    If UCase(sCurrency) = "USD" Then sSheetNametoUse = Sheet05Name_JEUploadUSD
    
    If UCase(sCurrency) <> "CAD" And UCase(sCurrency) <> "USD" Then
        MsgBox "Something wrong with currency, please check currency"
        Exit Sub
    End If
    sBankvsBank = ": " & wsItems.Cells(iRowItems, 2) & " " & wsItems.Cells(iRowItems, 4)
    
    dAMT = wsItems.Cells(iRowItems, 8)
    sABSAMT = CStr(Abs(dAMT))
    If dAMT > 0 Then
        sPostingKey1 = "40"
        sPostingKey2 = "50"
    Else
        sPostingKey1 = "50"
        sPostingKey2 = "40"
    End If
    
    
    sPostBU = wsItems.Cells(iRowItems, 1)
    
    sPostGL = Replace(wsItems.Cells(iRowItems, 3), " ", "")
    
    'find info for assignment and profit center
    sAss = ""
    sProfitC = ""
    sBankCode = wsItems.Cells(iRowItems, 2)
    Set rngFound = rngMapBankCode.Find(sBankCode, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        sAss = rngFound.Cells(1, SheetMapColAss - SheetMapColBankCode + 1)
        sProfitC = rngFound.Cells(1, SheetMapColProfitC - SheetMapColBankCode + 1)
    End If
    
    
    If Not Validate_GL(sPostGL) Then
        sPostingKey1 = CStr(CInt(sPostingKey1) - 20 + 1)
    End If
    
    If Validate_GL(sPostGL) Then
        Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey1, _
                    GLAccount:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)
    Else
            Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey1, _
                    Vendor:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)

    End If
    
    
    sPostBU = wsItems.Cells(iRowItems, 5)
    sPostGL = Replace(wsItems.Cells(iRowItems, 6), " ", "")
    
    'find info for assignment and profit center
    sAss = ""
    sProfitC = ""
    sBankCode = wsItems.Cells(iRowItems, 4)
    Set rngFound = rngMapBankCode.Find(sBankCode, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        sAss = rngFound.Cells(1, SheetMapColAss - SheetMapColBankCode + 1)
        sProfitC = rngFound.Cells(1, SheetMapColProfitC - SheetMapColBankCode + 1)
    End If
    
    If Not Validate_GL(sPostGL) Then
        sPostingKey2 = CStr(CInt(sPostingKey2) - 20 + 1)
    End If
    
    If Validate_GL(sPostGL) Then
        Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey2, _
                    GLAccount:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)
    Else
        Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey2, _
                    Vendor:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)
    
    End If
    
NEXTITEMLOOP1:
Next iRowItems

'Process vendor code posting

If iFoundCADVendor Then Call JE_Header_Info(Sheet05Name_JEUploadCAD, _
                                            sCompanyCodeHeaderCADVENDOR, _
                                            sDate, _
                                            sDate, _
                                            sLineText, _
                                            JEUpLoadDocType, _
                                            "CAD", _
                                            JEUpLoadJEType)
If iFoundUSDVendor Then Call JE_Header_Info(Sheet05Name_JEUploadUSD, _
                                            sCompanyCodeHeaderUSDVENDOR, _
                                            sDate, _
                                            sDate, _
                                            sLineText, _
                                            JEUpLoadDocType, _
                                            "USD", _
                                            JEUpLoadJEType)

For iRowItems = 2 To iMaxRowItems - 1
'For iRowItems = 2 To 2

    sPostGL1 = Replace(wsItems.Cells(iRowItems, 3), " ", "")
    sPostGL2 = Replace(wsItems.Cells(iRowItems, 6), " ", "")
    
    'If one bank account use vendor code, not regular GL, then skip it.
    If Validate_GL(sPostGL1) And Validate_GL(sPostGL2) Then GoTo NEXTITEMLOOP2
    
    sSheetNametoUse = ""
    sCurrency = UCase(Replace(wsItems.Cells(iRowItems, 7), " ", ""))
    If UCase(sCurrency) = "CAD" Then sSheetNametoUse = Sheet05Name_JEUploadCAD
    If UCase(sCurrency) = "USD" Then sSheetNametoUse = Sheet05Name_JEUploadUSD
    
    If UCase(sCurrency) <> "CAD" And UCase(sCurrency) <> "USD" Then
        MsgBox "Something wrong with currency, please check currency"
        Exit Sub
    End If
    sBankvsBank = ": " & wsItems.Cells(iRowItems, 2) & " " & wsItems.Cells(iRowItems, 4)
    
    dAMT = wsItems.Cells(iRowItems, 8)
    sABSAMT = CStr(Abs(dAMT))
    If dAMT > 0 Then
        sPostingKey1 = "40"
        sPostingKey2 = "50"
    Else
        sPostingKey1 = "50"
        sPostingKey2 = "40"
    End If
    
    
    sPostBU = wsItems.Cells(iRowItems, 1)
    
    sPostGL = Replace(wsItems.Cells(iRowItems, 3), " ", "")
    
    'find info for assignment and profit center
    sAss = ""
    sProfitC = ""
    sBankCode = wsItems.Cells(iRowItems, 2)
    Set rngFound = rngMapBankCode.Find(sBankCode, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        sAss = rngFound.Cells(1, SheetMapColAss - SheetMapColBankCode + 1)
        sProfitC = rngFound.Cells(1, SheetMapColProfitC - SheetMapColBankCode + 1)
    End If


    If Not Validate_GL(sPostGL) Then
        sPostingKey1 = CStr(CInt(sPostingKey1) - 20 + 1)
    End If
    
    If Validate_GL(sPostGL) Then
        Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey1, _
                    GLAccount:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)
    Else
            Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey1, _
                    Vendor:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)

    End If
    
    
    sPostBU = wsItems.Cells(iRowItems, 5)
    sPostGL = Replace(wsItems.Cells(iRowItems, 6), " ", "")
    
    If Not Validate_GL(sPostGL) Then
        sPostingKey2 = CStr(CInt(sPostingKey2) - 20 + 1)
    End If
    
    'find info for assignment and profit center
    sAss = ""
    sProfitC = ""
    sBankCode = wsItems.Cells(iRowItems, 4)
    Set rngFound = rngMapBankCode.Find(sBankCode, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        sAss = rngFound.Cells(1, SheetMapColAss - SheetMapColBankCode + 1)
        sProfitC = rngFound.Cells(1, SheetMapColProfitC - SheetMapColBankCode + 1)
    End If

    If Validate_GL(sPostGL) Then
        Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey2, _
                    GLAccount:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)
    Else
        Call JE_L_Line(JESheetName:=sSheetNametoUse, _
                    PostingKey:=sPostingKey2, _
                    Vendor:=sPostGL, _
                    NewCompanyCode:=sPostBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitC, _
                    Assignment:=sAss, _
                    LineText:=sLineText & sBankvsBank)
    
    End If
    
NEXTITEMLOOP2:
Next iRowItems


Worksheets(Sheet05Name_JEUploadUSD).Select
Worksheets(Sheet05Name_JEUploadUSD).Columns(19).Style = "Comma"

Worksheets(Sheet05Name_JEUploadCAD).Select
Worksheets(Sheet05Name_JEUploadCAD).Columns(19).Style = "Comma"


wkbMap.Close savechanges:=False
Application.DisplayAlerts = True

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
If iMaxRowJE < 4 Then Exit Sub

wsJE.Rows("4" & ":" & iMaxRowJE).Delete

Set wsJE = Nothing

'Debug.Print iMaxRowJE
End Sub


Sub JE_Header_Info(JESheetName As String, _
                    CompanyCodeinHLine As String, _
                    PostingDate As String, _
                    DocDate As String, _
                    DocHeaderText As String, _
                    DocType As String, _
                    CurrencyPost As String, _
                    JEType As String)

'Header information is on Row 4
Dim wsJE As Worksheet
Dim iRowMaxJE As Long
Dim iRowJE As Long
Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsJE = Worksheets(JESheetName)
wsJE.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iRowMaxJE = lRealLastRow
If iRowMaxJE < 3 Then
    MsgBox "Please check JE Template."
    Exit Sub
End If
iRowJE = iRowMaxJE + 1

wsJE.Cells(iRowJE, 1) = "H"
wsJE.Cells(iRowJE, 2) = CompanyCodeinHLine
wsJE.Cells(iRowJE, 3) = PostingDate
wsJE.Cells(iRowJE, 4) = DocDate
wsJE.Cells(iRowJE, 5) = DocType
wsJE.Cells(iRowJE, 6) = CurrencyPost
wsJE.Cells(iRowJE, 7) = DocHeaderText
wsJE.Cells(iRowJE, 8) = JEType
'Debug.Print iRowJE

'wsJE.Rows(19).Style = "Comma"

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



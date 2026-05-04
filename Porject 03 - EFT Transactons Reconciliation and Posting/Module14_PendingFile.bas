Attribute VB_Name = "Module14_PendingFile"
Option Explicit

Sub Move_to_Pending_File()

Dim sResponse As VbMsgBoxResult
sResponse = MsgBox("Pending items can only be saved once. Do you want to continue?", vbYesNo + vbQuestion, "Confirmation")

If sResponse = vbNo Then Exit Sub

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sPostBU As String

Dim wkbPending As Workbook
Dim sPendingFileFullName As String
Dim wsPending As Worksheet
Dim iMaxRowPending As Integer
Dim iRowPending As Integer


Dim lRealLastRow As Long
Dim lRealLastCol As Long

sPendingFileFullName = GetWorkPath & "\" & FileNamePending
Set wkbPending = Workbooks.Open(sPendingFileFullName)
Set wsPending = wkbPending.Worksheets("Pending")
wsPending.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowPending = lRealLastRow
iRowPending = iMaxRowPending


ThisWorkbook.Activate
Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub

For iRowItems = 2 To iMaxRowItems
    sPostBU = Cells(iRowItems, iColItemsPostBU)
    If Replace(sPostBU, " ", "") = "" Or InStr(sPostBU, WaitToConfirmInfo) > 0 Then
        iRowPending = iRowPending + 1
        wsPending.Cells(iRowPending, iColPendingPostingDate) = wsItems.Cells(iRowItems, iColItemsPostingDate)
        wsPending.Cells(iRowPending, iColPendingDocNumber) = wsItems.Cells(iRowItems, iColItemsDocNumber)
        wsPending.Cells(iRowPending, iColPendingGL) = wsItems.Cells(iRowItems, iColItemsGL)
        wsPending.Cells(iRowPending, iColPendingAMT) = wsItems.Cells(iRowItems, iColItemsAMT)
        wsPending.Cells(iRowPending, iColPendingBankInfo) = wsItems.Cells(iRowItems, iColItemsBankInfo)
        wsPending.Cells(iRowPending, iColPendingKeyBankAcct) = wsItems.Cells(iRowItems, iColItemsKeyBankAccount)
        wsPending.Cells(iRowPending, iColPendingPostBU) = wsItems.Cells(iRowItems, iColItemsPostBU)
        wsPending.Cells(iRowPending, iColPendingPostGL) = wsItems.Cells(iRowItems, iColItemsPostGL)
        wsPending.Cells(iRowPending, iColPendingPostVendor) = wsItems.Cells(iRowItems, iColItemsPostVendor)
        wsPending.Cells(iRowPending, iColPendingPostProfitCenter) = wsItems.Cells(iRowItems, iColItemsPostProfitC)
        wsPending.Cells(iRowPending, iColPendingPostCostCenter) = wsItems.Cells(iRowItems, iColItemsPostCostCenter)

    End If
Next iRowItems


wkbPending.Close SaveChanges:=True

Set wkbPending = Nothing
Set wsPending = Nothing
Set wsItems = Nothing

End Sub

Sub JE_Pending_Ready_to_Post()
Application.ScreenUpdating = False

Dim wkbPending As Workbook
Dim sPendingFileFullName As String
Dim wsPending As Worksheet
Dim iMaxRowPending As Integer
Dim iRowPending As Integer

Dim sBU As String
Dim sGL As String
Dim dAMT As Double
Dim sAMT As String
Dim sABSAMT As String
Dim sVendor As String
Dim sProfitCenter As String
Dim sKeyCode As String
Dim sAssInfo As String
Dim sJEPosted As String
Dim sOriGL As String
Dim dPendingTotalAMT As Double
Dim sPendingTotalAMT As String
Dim sDate As String
Dim sLineText As String
Dim sCostCenter As String

Dim sJESheetName As String

'Dim wsTempSheet As Worksheet
'Dim iMaxRowTemp As Integer
'Dim iRowTemp As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

sPendingFileFullName = GetWorkPath & "\" & FileNamePending
Set wkbPending = Workbooks.Open(sPendingFileFullName)
Set wsPending = wkbPending.Worksheets("Pending")
wsPending.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowPending = lRealLastRow
If iMaxRowPending < 2 Then Exit Sub

sJESheetName = "3 - C-SAP Standard Template"
ThisWorkbook.Activate
Call JE_Clean(sJESheetName)

wkbPending.Activate
'Call Check_Create_Temp_Sheet

dPendingTotalAMT = 0

sDate = wsPending.Cells(2, iColPendingPostingDate)
sLineText = "EFT " & sDate

For iRowPending = 2 To iMaxRowPending
    wkbPending.Activate
    wsPending.Select
    If Replace(wsPending.Cells(iRowPending, iColPendingPostBU), " ", "") = "" Then GoTo DONEXTLINE
    
    sOriGL = wsPending.Cells(iRowPending, iColPendingGL)
    
    dAMT = wsPending.Cells(iRowPending, iColPendingAMT)
    sAMT = CStr(dAMT)
    sABSAMT = CStr(Abs(dAMT))
    
    dPendingTotalAMT = dPendingTotalAMT + dAMT
    
    sBU = wsPending.Cells(iRowPending, iColPendingPostBU)
    sGL = wsPending.Cells(iRowPending, iColPendingPostGL)
    sVendor = wsPending.Cells(iRowPending, iColPendingPostVendor)
    sProfitCenter = wsPending.Cells(iRowPending, iColPendingPostProfitCenter)
    sAssInfo = wsPending.Cells(iRowPending, iColPendingPostAssInfo)
    sCostCenter = wsPending.Cells(iRowPending, iColPendingPostCostCenter)
    
    sKeyCode = ""
    If Replace(sGL, " ", "") <> "" Then
        If dAMT < 0 Then
            sKeyCode = "50"
        Else
            sKeyCode = "40"
        End If
    End If

    If Replace(sVendor, " ", "") <> "" Then
        If dAMT < 0 Then
            sKeyCode = "31"
        Else
            sKeyCode = "21"
        End If
    End If
    
    wsPending.Cells(iRowPending, iColPendingPostKeyCode) = sKeyCode
    wsPending.Cells(iRowPending, iColPendingJEPosted) = "Posted"
    
    
    ThisWorkbook.Activate
    Call JE_L_Line(JESheetName:=sJESheetName, _
                    PostingKey:=sKeyCode, _
                    GLAccount:=sGL, _
                    Vendor:=sVendor, _
                    NewCompanyCode:=sBU, _
                    DocumentCurrencyAmount:=sABSAMT, _
                    ProfitCenter:=sProfitCenter, _
                    Assignment:=sAssInfo, _
                    LineText:=sLineText, _
                    CostCenter:=sCostCenter)
  
                    



DONEXTLINE:
Next iRowPending




If dPendingTotalAMT < 0 Then
    sKeyCode = "40"
Else
    sKeyCode = "50"
End If

sPendingTotalAMT = CStr(Abs(dPendingTotalAMT))

ThisWorkbook.Activate
Call JE_L_Line(JESheetName:=sJESheetName, _
                PostingKey:=sKeyCode, _
                GLAccount:=SuspenseAccount, _
                Vendor:="", _
                NewCompanyCode:="9000", _
                DocumentCurrencyAmount:=sPendingTotalAMT, _
                ProfitCenter:="9510", _
                CostCenter:=sCostCenter, _
                FontBold:=True)

wkbPending.Close SaveChanges:=True

Set wkbPending = Nothing
Set wsPending = Nothing

Application.ScreenUpdating = True
End Sub


Sub Remove_Posted_Pending_items()

Dim sResponse As VbMsgBoxResult
sResponse = MsgBox("All posted items in Pening Transaction file will be deleted. Do you want to continue?", vbYesNo + vbQuestion, "Confirmation")

If sResponse = vbNo Then Exit Sub

Dim wkbPending As Workbook
Dim sPendingFileFullName As String
Dim wsPending As Worksheet
Dim iMaxRowPending As Integer
Dim iRowPending As Integer
Dim sJEPosted As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

sPendingFileFullName = GetWorkPath & "\" & FileNamePending
Set wkbPending = Workbooks.Open(sPendingFileFullName)
Set wsPending = wkbPending.Worksheets("Pending")
wsPending.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowPending = lRealLastRow
If iMaxRowPending < 2 Then Exit Sub


For iRowPending = iMaxRowPending To 2 Step -1
    sJEPosted = wsPending.Cells(iRowPending, iColPendingJEPosted)
    
    If UCase(Replace(sJEPosted, " ", "")) = "POSTED" Then
        Rows(iRowPending).Delete
    End If
Next iRowPending


wkbPending.Close SaveChanges:=True
End Sub

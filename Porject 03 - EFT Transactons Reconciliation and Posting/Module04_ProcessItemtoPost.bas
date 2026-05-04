Attribute VB_Name = "Module04_ProcessItemtoPost"
Option Explicit


Sub Filter_Items_to_Post()

Call Initialize_Sheet_Items_to_Process

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer

Dim wsItems As Worksheet
Dim iRowItems As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsItems = Worksheets("2-Items to post")

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow

If iMaxRowSAP < 2 Then Exit Sub
iRowItems = 1
For iRowSAP = 2 To iMaxRowSAP
    If Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = xlNone And InStr(UCase(Cells(iRowSAP, iColSAPClear)), "OFFSET") = 0 Then
        Debug.Print iRowSAP
        iRowItems = iRowItems + 1
        wsItems.Cells(iRowItems, iColItemsPostingDate) = wsSAP.Cells(iRowSAP, iColSAPPostingDate)
        wsItems.Cells(iRowItems, iColItemsDocNumber) = wsSAP.Cells(iRowSAP, iColSAPDocNumber)
        wsItems.Cells(iRowItems, iColItemsGL) = wsSAP.Cells(iRowSAP, iColSAPGL)
        wsItems.Cells(iRowItems, iColItemsAMT) = wsSAP.Cells(iRowSAP, iColSAPAMT)
    End If
Next iRowSAP

wsItems.Select
Cells.EntireColumn.AutoFit
Cells(1, 1).Select

End Sub

Sub Initialize_Sheet_Items_to_Process()

Dim wsItems As Worksheet

Set wsItems = Worksheets("2-Items to Post")
wsItems.Select
Cells.Select
Selection.Delete

Cells(1, iColItemsPostingDate) = "Posting Date"
Cells(1, iColItemsDocNumber) = "Document Number"
Cells(1, iColItemsGL) = "GL"
Cells(1, iColItemsAMT) = "Amount"
Cells(1, iColItemsBankInfo) = "Bank Info"
Cells(1, iColItemsKeyBankAccount) = "Key Bank Acct"

Columns(iColItemsPostingDate).NumberFormat = "mm/dd/yyy"
Columns(iColItemsAMT).Style = "Currency"

Columns(iColItemsPostingDate).HorizontalAlignment = xlCenter
Columns(iColItemsDocNumber).HorizontalAlignment = xlCenter
Columns(iColItemsGL).HorizontalAlignment = xlCenter

Cells(1, 1).Select
End Sub

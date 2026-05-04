Attribute VB_Name = "Module06_ReadBankInfo"
Option Explicit

Sub Find_Bank_Description()

Dim strBSFileFullName As String
Dim wkbBS As Workbook
Dim wsBS As Worksheet
Dim iMaxRowBS As Integer
Dim iRowBS As Integer
Dim strBSBankCode As String
Dim dBSAMT As Double
Dim strBSComment As String
Dim sBSBooked As String


Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim iGL As Integer
Dim dAMT As Double
Dim strItemsComment As String

Dim wsCon As Worksheet
Dim rngCon As Range
Dim rngFound As Range
Dim strBankCode As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long


ThisWorkbook.Activate

Set wsCon = Worksheets("Concentration & Clearing GL")
wsCon.Select
Set rngCon = Columns(iColConcenClear)

Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub


strBSFileFullName = Bank_Statement_File_Full_Name
If strBSFileFullName = "" Then
    MsgBox "Bank statement is missing."
    Exit Sub
End If

Set wkbBS = Workbooks.Open(strBSFileFullName)
wkbBS.Activate
Set wsBS = Worksheets(1)
wsBS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowBS = lRealLastRow
If iMaxRowBS < 2 Then
    MsgBox "Please check bank statement"
    wkbBS.Close SaveChanges:=False
    Exit Sub
End If


ThisWorkbook.Activate
wsItems.Select
For iRowItems = 2 To iMaxRowItems
'For iRowItems = 2 To 2
    iGL = CInt(Cells(iRowItems, iColItemsGL))
    dAMT = CDbl(Cells(iRowItems, iColItemsAMT))
    Debug.Print iRowItems & "--" & iGL & "--" & dAMT
    
    
    Set rngFound = rngCon.Find(iGL, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        strBankCode = rngFound.Cells(1, 2)
        Debug.Print "Bank Code :" & strBankCode
        Debug.Print iRowItems & "--" & iGL & "--" & strBankCode & "--" & dAMT
        
        strItemsComment = ""
        For iRowBS = 2 To iMaxRowBS
            strBSBankCode = wsBS.Cells(iRowBS, iColBSBankCode)
            dBSAMT = wsBS.Cells(iRowBS, iColBSAMT)
            strBSComment = wsBS.Cells(iRowBS, iColBSComment)
            'Debug.Print iRowBS & "--" & strBSBankCode & "--" & dBSAMT
            sBSBooked = UCase(Replace(wsBS.Cells(iRowBS, iColBSBooked), " ", ""))
            
            
            If dBSAMT = dAMT * (-1) And InStr(strBankCode, strBSBankCode) > 0 And sBSBooked <> "FOUND" Then
            'If dBSAMT = dAMT * (-1) Then
                Debug.Print "Found"
                strItemsComment = strBSComment
                Debug.Print iRowItems & "--" & iGL & "--" & dAMT & "--" & strItemsComment
                wsItems.Cells(iRowItems, iColItemsBankInfo) = strItemsComment
                
                wsBS.Cells(iRowBS, iColBSBooked) = "Found"
                Exit For
            End If
        Next iRowBS
        
    End If

Next iRowItems

wkbBS.Close SaveChanges:=False

End Sub

Sub Format_Items_Sheet_By_Bank_Code()

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim iGL As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub


Cells.Select
Cells.VerticalAlignment = xlVAlignCenter
Cells.EntireColumn.AutoFit
Cells(1, 1).Select

Columns(iColItemsBankInfo).ColumnWidth = 65
Columns(iColItemsBankInfo).WrapText = True

Columns(iColItemsKeyBankAccount).HorizontalAlignment = xlCenter
'Columns(iColItemsKeyBankAccount).EntireColumn.AutoFit
Columns(iColItemsKeyBankAccount).ColumnWidth = 25
Columns(iColItemsKeyBankAccount).WrapText = True

'To freeze top row
wsItems.Activate
With ActiveWindow:
    .FreezePanes = False
    .SplitColumn = 0
    .SplitRow = 1
    .FreezePanes = True
End With

End Sub

Attribute VB_Name = "Module010_Intial_FIS"
Option Explicit

Sub Mapping_010_Clear_Initialize_FIS()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim wsFIS As Worksheet
Dim iMaxRowFIS As Integer
Dim str4HeaderCombined As String
Dim strCell As String
Dim iStartRowHeader As Integer
Dim iTotalRow As Integer
Dim iCountAcct As Integer
Dim iMaxColFIS As Integer
'Dim jColProductCode As Integer
'Dim jColSource As Integer
'Dim jColRemark As Integer
Dim sBankAcctFIS As String



Dim i As Integer
Dim j As Integer

Dim wkbCashPosition As Workbook
Dim wsFormatting As Worksheet
Dim iFoundwsFormatting As Integer
Dim sFileNameTreasury As String
Dim iFormattingTotalRow As Integer
Dim iMaxRowFormatting As Integer
Dim iMaxColFormatting As Integer


Dim rngCopy As Range
Dim rngPaste As Range

'Clear FIS Sheet
Set wsFIS = Worksheets("FIS & PeopleSoft")
wsFIS.Select
Cells.Select
Selection.Delete
Selection.ClearFormats
Cells(1, 1).Select
Set rngPaste = Cells(1, 1)

'Open cash position file ("To read 01 - Treasury Report")
sFileNameTreasury = GetWorkPath & "\" & FileNameCashPosition
'Debug.Print sFileNameTreasury
Set wkbCashPosition = Workbooks.Open(sFileNameTreasury)

'To search sheet "Formatting", if it exists
'Debug.Print Sheets.Count()
iFoundwsFormatting = 0
For i = 1 To Sheets.Count()
    If Sheets(i).Name = "Formatting" Then
        iFoundwsFormatting = 1
        Exit For
    End If
Next i
If iFoundwsFormatting = 0 Then
    MsgBox "The file from Treasury, Sheet-""Formatting"" is missing."
    Exit Sub
End If


Set wsFormatting = wkbCashPosition.Worksheets("Formatting")
wsFormatting.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFormatting = lRealLastRow
iMaxColFormatting = lRealLastCol
If iMaxRowFormatting < 2 Then
    MsgBox "There is some problem with the file from Treasury, Sheet-Formatting."
    Exit Sub
End If

'To find row with word "TOTAL", as last row of "Formatting" sheet
iFormattingTotalRow = 0
For i = 1 To iMaxRowFormatting
    strCell = Cells(i, 1)
    strCell = Replace(strCell, " ", "")
    strCell = UCase(strCell)
    If InStr(strCell, "TOTAL") > 0 Then
        iFormattingTotalRow = i
        Exit For
    End If
Next i
If iFormattingTotalRow = 0 Then
    MsgBox "The file from Treasury, Sheet-Formatting is missing ""Total"" Lines."
    Exit Sub
End If
'Debug.Print iFormattingTotalRow

'To copy & paste from Cash Position file's "Formatting" to "FIS & PeopleSoft" Sheet
Set rngCopy = Range(Cells(1, 1), Cells(iFormattingTotalRow, iMaxColFormatting))
'Set rngCopy = Range(Cells(1, 1), Cells(6, iMaxColFormatting))
rngCopy.Copy
rngPaste.PasteSpecial Paste:=xlPasteValues
'rngPaste.PasteSpecial xlPasteFormats

'Close Cash Postion file, and work on Sheet "FIS & PeopleSoft"
Application.CutCopyMode = False
wkbCashPosition.Close savechanges:=False
wsFIS.Select
Cells.Select
Selection.EntireColumn.AutoFit
Cells(1, 1).Select

Call DeleteUnusedFormats

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
iMaxColFIS = lRealLastCol

'find Header row
iStartRowHeader = 0
For i = 1 To iMaxRowFIS
    str4HeaderCombined = ""
    For j = 1 To 4
        str4HeaderCombined = str4HeaderCombined & Cells(i, j)
    Next j
    str4HeaderCombined = Replace(str4HeaderCombined, " ", "")
    'Debug.Print str4HeaderCombined
    If str4HeaderCombined = FIS4Header Then
        iStartRowHeader = i
        Exit For
    End If
Next i

If iStartRowHeader = 0 Then
    MsgBox "Please check FIS Sheet for columns"
    Exit Sub
End If


'find last row, which is Total
iTotalRow = 0
For i = iStartRowHeader To iMaxRowFIS
    Cells(i, 1) = Replace(Cells(i, 1), " ", "")
    If UCase(Cells(i, 1)) = "TOTAL" Then
        iTotalRow = i
        Exit For
    End If
Next i

If iTotalRow = 0 Then
    Cells(iMaxRowFIS + 2, 1) = "Total"
    iMaxRowFIS = iMaxRowFIS + 2
    iTotalRow = iMaxRowFIS
End If

'count the number of account
iCountAcct = 0
For i = iStartRowHeader + 1 To iTotalRow - 1
    strCell = Cells(i, 1)
    strCell = Replace(strCell, " ", "")
    If strCell <> "" Then iCountAcct = iCountAcct + 1
Next i

Cells(iTotalRow, 2) = CStr(iCountAcct)


'delete unncecccsary columns in FIS sheet
wsFIS.Select
For j = iMaxColFIS To 1 Step -1
    If Replace(Cells(iStartRowHeader, j), " ", "") <> "FISCode" And Replace(Cells(iStartRowHeader, j), " ", "") <> "KyribaCode" And Replace(Cells(iStartRowHeader, j), " ", "") <> "BUFIS" And Replace(Cells(iStartRowHeader, j), " ", "") <> "GLCode" And Replace(Cells(iStartRowHeader, j), " ", "") <> "A/cNumber" And Replace(Cells(iStartRowHeader, j), " ", "") <> "CRY" And Replace(Cells(iStartRowHeader, j), " ", "") <> "Company" Then
        Columns(j).Delete
    End If
Next j

If iTotalRow > 0 Then Rows(iTotalRow).Delete
If iStartRowHeader > 1 Then Range(Rows(1), Rows(iStartRowHeader - 1)).Delete
Call DeleteUnusedFormats


lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
iMaxColFIS = lRealLastCol
If iMaxColFIS <> 7 Then
    MsgBox "For Treasuary report-Sheet Formatting, there is problem with columns, please check and run macro again"
    Exit Sub
End If

'jColProductCode = iMaxColFIS + 1
'jColSource = iMaxColFIS + 2
'jColRemark = iMaxColFIS + 3

wsFIS.Cells(1, ColFISRemark) = "Remark"
wsFIS.Cells(1, ColFISIsinPS) = "In PS"
wsFIS.Cells(1, ColFISIsinFIS) = "In Treasury"
wsFIS.Cells(1, ColFISProductCode) = "Product Code"
wsFIS.Cells(1, ColFISKeyNumber) = "Key Number"


'Add "Y" information in column of if it is from Cash Position report
If iMaxRowFIS > 1 Then
    For i = 2 To iMaxRowFIS
        wsFIS.Cells(i, ColFISIsinFIS) = "Y"
        
        sBankAcctFIS = wsFIS.Cells(i, ColFISBankAcct)
        sBankAcctFIS = Long_Bank_Account(sBankAcctFIS)
        
        'sBankAcctFIS = Replace(wsFIS.Cells(i, ColFISBankAcct), " ", "")
        'sBankAcctFIS = Replace(wsFIS.Cells(i, ColFISBankAcct), "'", "")
        'sBankAcctFIS = Remove_Leading_Zero(sBankAcctFIS)
        'sBankAcctFIS = "'" & CStr(sBankAcctFIS)
        
        wsFIS.Cells(i, ColFISBankAcct) = sBankAcctFIS
    Next i
End If

Call DeleteUnusedFormats





'Call Find_Deleted_Bank

'Call Find_New_Bank_Acct_in_FIS(iStartRowHeader + 1, iTotalRow - 1)

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub


Function Long_Bank_Account(BankAcctNo As String)
Long_Bank_Account = ""

Dim strInfo As String

strInfo = BankAcctNo
strInfo = Replace(strInfo, " ", "")
strInfo = Replace(strInfo, "'", "")
strInfo = Remove_Leading_Zero(strInfo)
strInfo = "'" & CStr(strInfo)

Long_Bank_Account = strInfo
End Function


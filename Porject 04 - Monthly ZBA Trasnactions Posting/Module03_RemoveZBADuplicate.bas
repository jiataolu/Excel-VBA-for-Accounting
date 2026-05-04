Attribute VB_Name = "Module03_RemoveZBADuplicate"
Option Explicit

Sub Remove_Duplicate_Rows()
Call Read_Bank_Data_Step_1_Clear_JEData_Sheet
Call Read_Bank_Data_Step_2_Copy_Paste_Data
Call Read_Bank_Data_Step_3_Remove_Duplicaate_ZBA_Data_for_2_ZBA
Call Read_Bank_Data_Step_4_for_1_ZBA

End Sub

Sub Read_Bank_Data_Step_1_Clear_JEData_Sheet()

Dim ws03JEDataClean As Worksheet

Set ws03JEDataClean = Worksheets(Sheet03Name_JEDataClean)
ws03JEDataClean.Select
Cells.Select
Selection.Delete

Cells.Interior.Pattern = xlNone
With Selection.Interior
    .Pattern = xlNone
    '.TintAndShade = 0
    '.PatternTintAndShade = 0
End With
Cells(1, 1).Select

End Sub

Sub Read_Bank_Data_Step_2_Copy_Paste_Data()

Dim ws02Data As Worksheet
Dim ws03DataClean As Worksheet

Dim rngCopy As Range
Dim rngPaste As Range

Set ws02Data = Worksheets(Sheet02Name_JEData)
ws02Data.Select
Cells.Select
Set rngCopy = Selection

Set ws03DataClean = Worksheets(Sheet03Name_JEDataClean)
ws03DataClean.Select
Cells(1, 1).Select
Set rngPaste = Cells(1, 1)

rngCopy.Copy Destination:=rngPaste

End Sub

Sub Read_Bank_Data_Step_3_Remove_Duplicaate_ZBA_Data_for_2_ZBA()

Dim ws03DataClean As Worksheet
Dim iMaxRowDataClean As Long
Dim iRowDataClean As Long
Dim sDup As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set ws03DataClean = Worksheets(Sheet03Name_JEDataClean)
ws03DataClean.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowDataClean = lRealLastRow
If iMaxRowDataClean < 2 Then Exit Sub

For iRowDataClean = iMaxRowDataClean To 2 Step -1
    sDup = UCase(ws03DataClean.Cells(iRowDataClean, Sheet02ColZBADuplicate))
    sDup = Replace(sDup, " ", "")
    If Len(sDup) > 0 Then
        If Left(sDup, 1) = "D" Then Rows(iRowDataClean).Delete
    End If
Next iRowDataClean

Call DeleteUnusedFormats

End Sub

Sub Read_Bank_Data_Step_4_for_1_ZBA()

Dim wsCopy As Worksheet

Dim wsWork As Worksheet
Dim iMaxRowWork As Long
Dim iRowWork As Long

Dim sBankCode As String
Dim sBU As String
Dim sGL As String
Dim sZBABankCode As String
Dim sZBABU As String
Dim sZBAGL As String
Dim sAmount As String

Dim sBankCodeFrom As String
Dim sBUFrom As String
Dim sGLFrom As String
Dim sBankCodeTo As String
Dim sBUTo As String
Dim sGLTo As String
Dim sAmountAdj As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim rngCopy As Range
Dim rngPaste As Range

Set wsCopy = Worksheets(Sheet03Name_JEDataClean)
wsCopy.Select
Cells.Select
Set rngCopy = Selection

Set wsWork = Worksheets(Sheet03Name_JEDataClean1ZBA)
wsWork.Select
Cells.Select
Selection.Delete

Cells.Interior.Pattern = xlNone
With Selection.Interior
    .Pattern = xlNone
    '.TintAndShade = 0
    '.PatternTintAndShade = 0
End With
Cells(1, 1).Select
Set rngPaste = Cells(1, 1)

rngCopy.Copy Destination:=rngPaste

wsWork.Select
Columns(Sheet02ColZBADuplicate).Delete
Cells(1, Sheet02ColFromBankCode) = "Bank_Code_1"
Cells(1, Sheet02ColFromBU) = "BU_1"
Cells(1, Sheet02ColFromGL) = "GL_1"
Cells(1, Sheet02ColToBankCode) = "Bank_Code_2"
Cells(1, Sheet02ColToBU) = "BU_2"
Cells(1, Sheet02ColToGL) = "GL_2"
Cells(1, Sheet02ColAmountAdj) = "Amount_ADJ"

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowWork = lRealLastRow
If iMaxRowWork < 2 Then Exit Sub

       
For iRowWork = 2 To iMaxRowWork
    sBankCode = Cells(iRowWork, Sheet02ColAccountBankCode)
    sBU = Cells(iRowWork, Sheet02ColAccountBU)
    sGL = Cells(iRowWork, Sheet02ColAccountGL)
    
    sZBABankCode = Cells(iRowWork, Sheet02ColZBABankCode)
    sZBABU = Cells(iRowWork, Sheet02ColZBABU)
    sZBAGL = Cells(iRowWork, Sheet02ColZBAGL)
    
    sAmount = Cells(iRowWork, Sheet02ColAmount)
    
    If sBankCode < sZBABankCode Then
        sBankCodeFrom = sBankCode
        sBUFrom = sBU
        sGLFrom = sGL
        sBankCodeTo = sZBABankCode
        sBUTo = sZBABU
        sGLTo = sZBAGL
        sAmountAdj = sAmount
    Else
        sBankCodeFrom = sZBABankCode
        sBUFrom = sZBABU
        sGLFrom = sZBAGL
        sBankCodeTo = sBankCode
        sBUTo = sBU
        sGLTo = sGL
        sAmountAdj = CStr(CDbl(sAmount * -1))
    End If
    
    Cells(iRowWork, Sheet02ColFromBankCode) = sBankCodeFrom
    Cells(iRowWork, Sheet02ColFromBU) = sBUFrom
    Cells(iRowWork, Sheet02ColFromGL) = sGLFrom
    Cells(iRowWork, Sheet02ColToBankCode) = sBankCodeTo
    Cells(iRowWork, Sheet02ColToBU) = sBUTo
    Cells(iRowWork, Sheet02ColToGL) = sGLTo
    Cells(iRowWork, Sheet02ColAmountAdj) = sAmountAdj
    
Next iRowWork

Columns(Sheet02ColFromBankCode).HorizontalAlignment = xlCenter
Columns(Sheet02ColFromBU).HorizontalAlignment = xlCenter
Columns(Sheet02ColFromGL).HorizontalAlignment = xlCenter
Columns(Sheet02ColToBankCode).HorizontalAlignment = xlCenter
Columns(Sheet02ColToBU).HorizontalAlignment = xlCenter
Columns(Sheet02ColToGL).HorizontalAlignment = xlCenter
Columns(Sheet02ColAmountAdj).Style = "Comma"

Cells.Select
Cells.EntireColumn.AutoFit
Cells(1, 1).Select

End Sub

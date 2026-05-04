Attribute VB_Name = "Module21_S4Data"
Option Explicit

Sub Transfer_S4_Data()

Dim sFullFileNameS4 As String
Dim wkbS4 As Workbook
Dim wsS4 As Worksheet
Dim iMaxRowS4 As Long
Dim iMaxColS4 As Integer
Dim iRowS4 As Long
Dim iColS4 As Integer
Dim sHeaderS4 As String


Dim sFullFileNameMSD As String
Dim wkbMSD As Workbook
Dim wsMSD As Worksheet
Dim iMaxRowMSD As Long
Dim iMaxColMSD As Integer
Dim iRowMSD As Long
Dim iColMSD As Integer
Dim sHeaderMSD As String


Dim sCellValue As String
Dim sAcctNo As String
Dim rngCopy As Range
Dim rngPaste As Range


Dim lRealLastRow As Long
Dim lRealLastCol As Long

sFullFileNameMSD = GetWorkPath & "\" & SubFolder & "\" & FileSAPMSD
Set wkbMSD = Workbooks.Open(sFullFileNameMSD)
wkbMSD.Activate

Set wsMSD = Worksheets(1)
wsMSD.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMSD = lRealLastRow
iMaxColMSD = lRealLastCol
If iMaxRowMSD > 1 Then Rows("2:" & iMaxRowMSD).Delete


sFullFileNameS4 = GetWorkPath & "\" & SubFolder & "\" & FileS4MSD
Set wkbS4 = Workbooks.Open(sFullFileNameS4)
wkbS4.Activate

Set wsS4 = Worksheets(1)
wsS4.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowS4 = lRealLastRow
iMaxColS4 = lRealLastCol


' Copy Paste all columns with same header
For iColMSD = 1 To iMaxColMSD
    wkbMSD.Activate
    wsMSD.Select
    sHeaderMSD = wsMSD.Cells(1, iColMSD)
    Set rngPaste = Cells(2, iColMSD)
    
    wkbS4.Activate
    wsS4.Select
    For iColS4 = 1 To iMaxColS4
        sHeaderS4 = wsS4.Cells(1, iColS4)
        If UCase(sHeaderS4) = UCase(sHeaderMSD) Then
        
            Set rngCopy = Range(Cells(2, iColS4), Cells(iMaxRowS4, iColS4))
            rngCopy.Copy
            rngPaste.PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            
            GoTo CONTINUECOLMSD1
        End If
    Next iColS4
    
CONTINUECOLMSD1:
Next iColMSD


'Copy Paste Column Account
For iColMSD = 1 To iMaxColMSD
    wkbMSD.Activate
    wsMSD.Select
    sHeaderMSD = wsMSD.Cells(1, iColMSD)
    If sHeaderMSD <> "Account" Then GoTo CONTINUECOLMSD2
    
    Set rngPaste = Cells(2, iColMSD)
    Debug.Print sHeaderMSD
    
    wkbS4.Activate
    wsS4.Select
    For iColS4 = 1 To iMaxColS4
        sHeaderS4 = wsS4.Cells(1, iColS4)
        If sHeaderS4 = "Cleared/Open Items Symbol" Then
        
            Debug.Print iColS4
            Set rngCopy = Range(Cells(2, iColS4), Cells(iMaxRowS4, iColS4))
            rngCopy.Copy
            rngPaste.PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            
            GoTo CONTINUECOLMSD2
        End If
    Next iColS4
    
CONTINUECOLMSD2:
Next iColMSD

wkbS4.Close SaveChanges:=False

'Clear Column Account
wkbMSD.Activate
wsMSD.Select

For iRowS4 = iMaxRowS4 To 2 Step -1
    sCellValue = Replace(wsMSD.Cells(iRowS4, 1), " ", "")
    If sCellValue <> "" And sCellValue <> "@5B\QCleared@" Then
        sAcctNo = sCellValue
        sAcctNo = Replace(sAcctNo, "Account", "")
        wsMSD.Cells(iRowS4, 1) = sAcctNo
        
        Range(Cells(iRowS4, 1), Cells(iRowS4, 11)).Interior.Color = RGB(255, 255, 153)
    End If
    
    If sCellValue = "@5B\QCleared@" Then
        'wsMSD.Cells(iRowS4, 1).ClearContents
        'wsMSD.Cells(iRowS4, 1) = sAcctNo
        Range(Cells(iRowS4, 1), Cells(iRowS4, 11)).Interior.Color = RGB(255, 255, 153)
    End If
    
    If sCellValue = "" Then wsMSD.Cells(iRowS4, 1) = sAcctNo
Next iRowS4

'Column Assignment

For iRowS4 = 2 To iMaxRowS4
    sCellValue = wsMSD.Cells(iRowS4, 5)
    
    If Replace(Cells(iRowS4, 4), " ", "") <> "" Then wsMSD.Cells(iRowS4, 4) = sCellValue

Next iRowS4

wkbMSD.Close SaveChanges:=True

End Sub


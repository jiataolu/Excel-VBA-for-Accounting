Attribute VB_Name = "Module14_Cash_Position_report"
Option Explicit

Sub Cash_Position_Report()

Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
    
Dim strYearNumber As String
Dim strMonthNumber As String
Dim strMonthLetter As String
Dim strTodayDate As String

'Dim strFileNamexlsm As String
'Dim strFullFileNamexlsm As String
Dim strFileNamexlsx As String
Dim strFullFileNamexlsx As String

Dim wkbNewxlsxFile As Workbook
Dim iCountSheetNew As Integer
Dim strSheetNameTemp As String
Dim i As Integer
Dim wsNew As Worksheet
Dim wsCurrent As Worksheet

Dim wkbMacro As Workbook
Dim wsMacro As Worksheet
Dim strSheetNameMacro As String
Dim iCountSheetMacro As Integer
Dim iHidden As Integer
Dim lColor As Long
Dim wsCashProject As Worksheet
Dim iMaxRowCash As Integer
Dim iRowFormulaERPFIS As Integer

Dim rngCopy As Range
Dim rngPaste As Range

Dim sFormula1 As String
Dim sFormula2 As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

strTodayDate = Date

strYearNumber = CStr(Format(strTodayDate, "yyyy"))
'Debug.Print strYearNumber

strMonthNumber = CStr(Format(strTodayDate, "MM"))
'Debug.Print strMonthNumber

strMonthLetter = Format(strTodayDate, "mmm")
'Debug.Print strMonthLetter

'strFileNamexlsm = strYearNumber & "_" & strMonthNumber & " Cash Position Report " & strMonthLetter & " " & strYearNumber & ".xlsm"
'Debug.Print strFileNamexlsm

'strFullFileNamexlsm = GetWorkPath & "\" & strFileNamexlsm
'Debug.Print strFullFileNamexlsm

'ThisWorkbook.SaveCopyAs strFullFileNamexlsm


strFileNamexlsx = strYearNumber & "_" & strMonthNumber & " Cash Position Report " & strMonthLetter & " " & strYearNumber & ".xlsx"
'Debug.Print strFileNamexlsx

strFullFileNamexlsx = GetWorkPath & "\" & strFileNamexlsx
'Debug.Print strFullFileNamexlsx

strSheetNameTemp = "$$$Temp"
Set wkbNewxlsxFile = Workbooks.Add
wkbNewxlsxFile.Activate

iCountSheetNew = Worksheets.Count
'Debug.Print iCountSheetNew
For i = 1 To iCountSheetNew
    Worksheets(i).Name = strSheetNameTemp & CStr(i)
Next i

Set wkbMacro = ThisWorkbook
wkbMacro.Activate
Set wsCashProject = Worksheets("Cash Project")
wsCashProject.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowCash = lRealLastRow

sFormula1 = ""
sFormula2 = ""
iRowFormulaERPFIS = 0
For i = 1 To iMaxRowCash
    If Cells(i, 1) = "ERP & FIS" Then
        sFormula1 = Cells(i, iCPAmtERP).Formula
        sFormula2 = Cells(i, iCPAmtBank).Formula
        iRowFormulaERPFIS = i
        Exit For
    End If
Next i
'Debug.Print sFormula1
'Debug.Print sFormula2

iCountSheetMacro = Worksheets.Count



For i = 1 To iCountSheetMacro
    wkbMacro.Activate
    Worksheets(i).Activate
    strSheetNameMacro = Worksheets(i).Name
    iHidden = Worksheets(i).Visible
    lColor = Worksheets(i).Tab.Color
    'Debug.Print strSheetNameMacro
    'Debug.Print lColor
    Cells.Select
    Set rngCopy = Selection
    
    wkbNewxlsxFile.Activate
    Set wsCurrent = ActiveSheet
    Set wsNew = Worksheets.Add(after:=wsCurrent)
    wsNew.Name = strSheetNameMacro
    wsNew.Select
    Set rngPaste = Cells(1, 1)
    
    rngCopy.Copy Destination:=rngPaste
    Cells(1, 1).Select
    wsNew.Visible = iHidden
    'Debug.Print wsCurrent.Name
    If lColor <> False Then wsNew.Tab.Color = lColor

Next i

wkbNewxlsxFile.Activate
Set wsCashProject = Worksheets("Cash Project")
wsCashProject.Select
If iRowFormulaERPFIS > 0 Then
    Cells(iRowFormulaERPFIS, iCPAmtERP).Formula = sFormula1
    Cells(iRowFormulaERPFIS, iCPAmtBank).Formula = sFormula2
End If


For i = 1 To iCountSheetNew
    Worksheets(strSheetNameTemp & CStr(i)).Delete
Next i

'to protect sheet GL-Bank
Worksheets("GL-Bank").Select
Cells(1, 1).Select
Selection.AutoFilter

Range("A1").Select
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True
Worksheets("GL-Bank").Protect Password:="banking", Contents:=True, AllowFiltering:=True

'To protect sheet "Mapping"
Worksheets("Mapping").Select
Cells(1, 1).Select
Selection.AutoFilter

Range("A1").Select
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True
Worksheets("Mapping").Protect Password:="banking", Contents:=True, AllowFiltering:=True

Worksheets("Cash Project").Select


wkbNewxlsxFile.SaveCopyAs strFullFileNamexlsx
wkbNewxlsxFile.Close savechanges:=False

ThisWorkbook.Activate
Worksheets("Cash Project").Select
Cells(1, 1).Select

' Re-enable macro execution
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

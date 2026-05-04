Attribute VB_Name = "Module52_SubFunction1"
Option Explicit

Sub Sheet_Temp_Check_Init()

Dim iSheetCount As Integer
Dim i As Integer
Dim iSheetFound As Integer
Dim wsTemp As Worksheet

iSheetFound = 0
For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "$$$Temp" Then
        iSheetFound = 1
        Exit For
    End If
Next

If iSheetFound = 0 Then
    Set wsTemp = Worksheets.Add
    wsTemp.Name = "$$$Temp"
Else
    Set wsTemp = Worksheets("$$$Temp")
End If

wsTemp.Select
Cells.Select
Selection.ClearContents

Cells(1, 1) = "Group - PART"
Cells(2, 1) = "GL"
Cells(2, 2) = "Leading Number"
Cells(2, 3) = "Total Amount"

Cells(1, 6) = "Group - Empty Assigment Field, Non-Empty Text Field (58155KH90)"
Cells(2, 6) = "GL"
Cells(2, 7) = "Total Amount"
End Sub

Sub Sheet_Temp_Delete()
Application.DisplayAlerts = False

Dim iSheetCount As Integer
Dim i As Integer
Dim iSheetFound As Integer
Dim wsTemp As Worksheet

iSheetFound = 0
For i = 1 To Worksheets.Count()
    If Worksheets(i).Name = "$$$Temp" Then
        iSheetFound = 1
        Exit For
    End If
Next

If iSheetFound = 1 Then
    Set wsTemp = Worksheets("$$$Temp")
    wsTemp.Delete
End If

Application.DisplayAlerts = True
End Sub


Sub Sheet_Temp_002_Check_Init()

Dim iSheetCount As Integer
Dim i As Integer
Dim iSheetFound As Integer
Dim wsTemp As Worksheet

iSheetFound = 0
For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "$$$Temp002" Then
        iSheetFound = 1
        Exit For
    End If
Next

If iSheetFound = 0 Then
    Set wsTemp = Worksheets.Add
    wsTemp.Name = "$$$Temp002"
Else
    Set wsTemp = Worksheets("$$$Temp002")
End If

wsTemp.Select
Cells.Select
Selection.ClearContents
End Sub

Sub Sheet_Temp_002_Delete()
Application.DisplayAlerts = False

Dim iSheetCount As Integer
Dim i As Integer
Dim iSheetFound As Integer
Dim wsTemp As Worksheet

iSheetFound = 0
For i = 1 To Worksheets.Count()
    If Worksheets(i).Name = "$$$Temp002" Then
        iSheetFound = 1
        Exit For
    End If
Next

If iSheetFound = 1 Then
    Set wsTemp = Worksheets("$$$Temp002")
    wsTemp.Delete
End If

Application.DisplayAlerts = True
End Sub







Function Leading_Number(Description As String)
Leading_Number = ""

Dim iPosition As Integer
iPosition = InStr(Description, " ")
If iPosition > 1 Then
    Leading_Number = Left(Description, iPosition - 1)
    Leading_Number = Right(Leading_Number, 4)
End If
End Function

Sub DeleteUnusedFormats()
    Dim lLastRow As Long, lLastColumn As Long
    Dim lRealLastRow As Long, lRealLastColumn As Long
    
    With Range("a1").SpecialCells(xlCellTypeLastCell)
        lLastRow = .Row
        lLastColumn = .Column
    End With
    lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
    lRealLastColumn = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
    
    If lRealLastRow < lLastRow Then
        Range(Cells(lRealLastRow + 1, 1), Cells(lLastRow, 1)).EntireRow.Delete
    End If
    If lRealLastColumn < lLastColumn Then
        Range(Cells(1, lRealLastColumn + 1), Cells(1, lLastColumn)).EntireColumn.Delete
    End If
    ActiveSheet.UsedRange
    
End Sub


Attribute VB_Name = "Module92_PublicSub"
Option Explicit

Public Const NameHD As String = "jlu"
Public Const NameHTTP As String = "jiatao_lu"

Function GetWorkPath() As String

GetWorkPath = ""

Dim sTEAMPath As String
'sTEAMPath = "https://psav-my.sharepoint.com/personal/jlu_psav_com/Documents"
sTEAMPath = "https://mckessoncorpca-my.sharepoint.com/personal/" & NameHTTP & "_mckesson_ca/Documents"

Dim sNormPath As String
'sNormPath = "C:\Users\JiataoLu\OneDrive - Encore"
sNormPath = "C:\Users\" & NameHD & "\OneDrive - McKesson Corporation"

Dim sPath As String
sPath = ThisWorkbook.Path

If UCase(Left(sPath, 2)) = "C:" Then
    GetWorkPath = sPath
    Exit Function
End If

sPath = Replace(sPath, sTEAMPath, sNormPath)
'Debug.Print sPath
sPath = Replace(sPath, "/", "\")
'Debug.Print sPath

GetWorkPath = sPath

End Function



Function Map_File_Full_Name()
Map_File_Full_Name = ""

Dim objFSO As Object
Dim objFolder As Object
Dim objFiles As Object
Dim objFile As Object
Dim iCountFiles As Integer

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.Getfolder(GetWorkPath())
Set objFiles = objFolder.Files

iCountFiles = 0
For Each objFile In objFiles
    If InStr(UCase(objFile.Name), "MAPPING") > 0 And InStr(UCase(objFile.Name), ".XLS") > 0 Then
        Map_File_Full_Name = GetWorkPath() & "\" & objFile.Name
        iCountFiles = iCountFiles + 1
    End If
Next

If iCountFiles = 0 Then MsgBox "There is no mapping file, please check."

If iCountFiles > 1 Then MsgBox "There are more than one mapping file, please check."

End Function

Function Month_End_Date(month As Integer, year As Integer) As String
    ' Get the first day of the next month
    Dim firstDayNextMonth As Date
    Dim lastDayCurrentMonth As Integer
    firstDayNextMonth = DateSerial(year, month + 1, 1)
    
    ' Subtract one day to get the last day of the current month
    lastDayCurrentMonth = Day(firstDayNextMonth - 1)
    Month_End_Date = month & "/" & lastDayCurrentMonth & "/" & year
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


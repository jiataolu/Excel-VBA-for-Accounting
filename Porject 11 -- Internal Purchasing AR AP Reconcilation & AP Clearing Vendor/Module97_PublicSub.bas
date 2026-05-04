Attribute VB_Name = "Module97_PublicSub"
Option Explicit

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


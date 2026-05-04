Attribute VB_Name = "Module58_SubFunction4"
Option Explicit








Function Number_wt_Leading_Zero(NumberString As String) As String

Number_wt_Leading_Zero = NumberString

While Left(Number_wt_Leading_Zero, 1) = 0 And Len(Number_wt_Leading_Zero) > 1
    Number_wt_Leading_Zero = Right(Number_wt_Leading_Zero, Len(Number_wt_Leading_Zero) - 1)
Wend

End Function

Function Map_File_Full_Name()
Map_File_Full_Name = ""

Dim objFSO As Object
Dim objFolder As Object
Dim objFiles As Object
Dim objFile As Object
Dim iCountFiles As Integer

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.Getfolder(GetWorkPath() & "\" & SubFolderMapping)
Set objFiles = objFolder.Files

iCountFiles = 0
For Each objFile In objFiles
    If InStr(UCase(objFile.Name), "MAPPING") > 0 And InStr(UCase(objFile.Name), ".XLS") > 0 Then
        Map_File_Full_Name = GetWorkPath() & "\" & SubFolderMapping & "\" & objFile.Name
        iCountFiles = iCountFiles + 1
    End If
Next

If iCountFiles = 0 Then MsgBox "There is no mapping file, please check."

If iCountFiles > 1 Then MsgBox "There are more than one mapping file, please check."

End Function

Function Bank_Statement_File_Full_Name()
Bank_Statement_File_Full_Name = ""


Dim strBSFileFullName As String
Dim objFSO As Object
Dim objFolder As Object
Dim objFiles As Object
Dim objFile As Object


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.Getfolder(GetWorkPath() & "\" & SubFolderInput)
Set objFiles = objFolder.Files

strBSFileFullName = ""
For Each objFile In objFiles
    If InStr(UCase(objFile.Name), UCase(FileNameBS)) > 0 And InStr(UCase(objFile.Name), ".XLSX") > 0 Then
        strBSFileFullName = GetWorkPath() & "\" & SubFolderInput & "\" & objFile.Name
        Exit For
    End If
Next

If strBSFileFullName = "" Then
    MsgBox "Bank statement is missing."
End If


Bank_Statement_File_Full_Name = strBSFileFullName
'Debug.Print Bank_Statement_File_Full_Name
End Function

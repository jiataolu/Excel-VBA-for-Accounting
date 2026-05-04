Attribute VB_Name = "Module54_SubFunction2"
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



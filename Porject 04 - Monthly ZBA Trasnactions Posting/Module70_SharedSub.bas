Attribute VB_Name = "Module70_SharedSub"
Option Explicit

Function Validate_GL(InputGL As String)

Validate_GL = False

Dim i As Integer
Dim iLen As Integer
Dim sOneGL As String
Dim iAscOneGL As Integer

iLen = Len(InputGL)

If iLen > 8 Then Exit Function

If iLen = 0 Then Exit Function

For i = 1 To iLen
    sOneGL = Mid(InputGL, i, 1)
    iAscOneGL = Asc(sOneGL)
    
    If iAscOneGL < 48 Or iAscOneGL > 57 Then Exit Function
Next i


Validate_GL = True
End Function



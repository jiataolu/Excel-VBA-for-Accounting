Attribute VB_Name = "Module12_CP_Bank_Code"
Option Explicit

Sub Bank_Code_List_in_Cash_Project()

'set up for Regex
Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim strMyString As String

Set objRegex = CreateObject("VBScript.RegExp")
strPattern = "(\w)+(-)(\d)+"
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern


Dim wsCP As Worksheet
Dim i As Integer
Dim iNextLine As Integer
Dim strCell As String

Set wsCP = Worksheets("Cash Project")
wsCP.Select
iNextLine = 2

Columns(iCPBankCode + 2).Select
Selection.Delete

Cells(1, iCPBankCode + 2) = "Code"

'For i = 7 To 7
For i = 2 To 540
    strCell = wsCP.Cells(i, 8)
    'Debug.Print strCell
    strMyString = strCell
    
    bolTest = objRegex.Test(strMyString)
    Set objMatches = objRegex.Execute(strMyString)

    If bolTest = True Then
        For Each objMatch In objMatches
            'Debug.Print objMatch.Value
            wsCP.Cells(iNextLine, iCPBankCode + 2) = objMatch.Value
            iNextLine = iNextLine + 1
        Next objMatch
    End If
    
Next i


End Sub


Attribute VB_Name = "Module56_SubFunction3BankInfo"
Option Explicit

Function BOA_ORIG_Info(Comments As String)
BOA_ORIG_Info = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object


Dim strMyString As String
strMyString = Comments

strPattern = "ORIG:(.*?)ORG"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    BOA_ORIG_Info = objMatches(0).Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function

Function BOA_ORIG_Bank_Account(OrigInfo As String)
BOA_ORIG_Bank_Account = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

strPattern = "ID:(.*)ORG"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    BOA_ORIG_Bank_Account = objSubMatch.SubMatches(0)
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function

Function BOA_ORIG_Info_B(Comments As String)
BOA_ORIG_Info_B = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object


Dim strMyString As String
strMyString = Comments

strPattern = "ORIG:(.*?)BNF"    '(.*?) non-greedy mode

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    BOA_ORIG_Info_B = objMatches(0).Value
    BOA_ORIG_Info_B = Replace(BOA_ORIG_Info_B, "BNF", "")
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function


Function BOA_ORIG_Bank_Account_B(OrigInfo As String)
BOA_ORIG_Bank_Account_B = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

strPattern = "ID:(.*)"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    BOA_ORIG_Bank_Account_B = objSubMatch.SubMatches(0)
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function



Function BOA_BNF_Info(Comments As String)
BOA_BNF_Info = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object

Dim strMyString As String
strMyString = Comments

strPattern = "BNF:(.*?)BNF"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then BOA_BNF_Info = objMatches(0).Value

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function


Function BOA_BNF_Bank_Account(OrigInfo As String)
BOA_BNF_Bank_Account = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

strPattern = "ID:(.*)BNF"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    BOA_BNF_Bank_Account = objSubMatch.SubMatches(0)
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function


Function USB_Info(OrigInfo As String)
USB_Info = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

If InStr(UCase(OrigInfo), "FUNDS TRANSFER") = 0 Then Exit Function

strPattern = "act(\D*)(\d*)"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    'USB_Info = objSubMatch.SubMatches(1)
    USB_Info = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function

Function USB_Bank_Account(OrigInfo As String)
USB_Bank_Account = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

strPattern = "\d+"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    USB_Bank_Account = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function


Function WFB_Info(OrigInfo As String)
WFB_Info = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

If InStr(UCase(OrigInfo), "ZBA FUNDING ACCOUNT TRANSFER") = 0 Then Exit Function

strPattern = "(FROM|TO)(\D*)(\d*)"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    'USB_Info = objSubMatch.SubMatches(1)
    WFB_Info = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function

Function WFB_Bank_Account(OrigInfo As String)
WFB_Bank_Account = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

strPattern = "\d+"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    WFB_Bank_Account = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function

Function JPM_Info(OrigInfo As String)
JPM_Info = ""
'Debug.Print OrigInfo

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

'If InStr(UCase(OrigInfo), "ZBA FUNDING ACCOUNT TRANSFER") = 0 Then Exit Function

strPattern = "(FROM|TO)(\D*)(\d*)"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    'USB_Info = objSubMatch.SubMatches(1)
    JPM_Info = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function

Function JPM_Bank_Account(OrigInfo As String)
JPM_Bank_Account = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = OrigInfo

strPattern = "\d+"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    JPM_Bank_Account = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing

End Function


Function Number_String_Only_First(Info As String)
Number_String_Only_First = ""

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object
Dim objSubMatch As Object

Dim strMyString As String
strMyString = Info

strPattern = "\d+"

Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    Set objSubMatch = objMatches(0)
    Number_String_Only_First = objSubMatch.Value
End If

Set objRegex = Nothing
Set objMatches = Nothing
Set objMatch = Nothing


End Function


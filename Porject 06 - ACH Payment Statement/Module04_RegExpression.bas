Attribute VB_Name = "Module04_RegExpression"
Option Explicit

Function RegExp_CK(CKinfo As String)

'Dim CKinfo As String
'CKinfo = "4428**01*075911988*DA*4900826065*20250207\TRN*1*0000514005491386404\REF*CK*1386400047939482"
RegExp_CK = ""

Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")
regex.Pattern = "(TRN\*1\*000051400549|\*TN\*000051400549)(\d+)"
regex.Global = True


Dim matches As Object
Set matches = regex.Execute(CKinfo)


'Debug.Print CKinfo

'CKinfo = "4428**01*121000248*DA*4185452604*20250926\TRN*1*0000514005491466905\REF*CK*1466900046748226"
'Debug.Print CKinfo


RegExp_CK = matches(0).SubMatches(1)
'Debug.Print matches(0).SubMatches(1)

'Debug.Print matches.Count

'Dim match As Object
'For Each match In matches
    'Debug.Print match.Value
'Next match
End Function


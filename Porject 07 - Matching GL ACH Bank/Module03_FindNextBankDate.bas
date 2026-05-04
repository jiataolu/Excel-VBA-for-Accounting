Attribute VB_Name = "Module03_FindNextBankDate"
Option Explicit
Option Base 1

Function Find_Next_Date_in_Bank(TransDate As Date, BankDateList As Variant)
'Dim TransDate As Date
'TransDate = #2/2/2026#

Dim i As Integer
Dim Found As Integer

Find_Next_Date_in_Bank = TransDate

Found = 0
For i = 1 To UBound(BankDateList, 1)
    'Debug.Print "i=" & i & ":" & arrBankDate(i, 1)
    'Debug.Print TransDate
    If TransDate < BankDateList(i, 1) Then
        Found = 1
        Find_Next_Date_in_Bank = BankDateList(i, 1)
        Exit For
    End If
Next i
    If Found = 0 Then Find_Next_Date_in_Bank = Find_Next_Date_in_Bank + 1
End Function


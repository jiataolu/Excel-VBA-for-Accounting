Attribute VB_Name = "Module020_Add_PeopleSoft"
Option Explicit

Sub Mapping_020_Combine_PeopleSoft_FIS()
Application.ScreenUpdating = False
Application.DisplayAlerts = False



'Sheet FIS & PeopleSoft
Dim wsFIS As Worksheet
Dim iMaxRowFIS As Long
Dim iRowFIS As Long
Dim sBankAcctFIS As String
'Dim rngFISBankAcct As Range
'Dim rngFound As Range
Dim iRowCurrentFIS As Long
Dim iFound As Integer



'PeopleSoft file
Dim sFileNamePS As String
Dim wkbPS As Workbook
Dim wsPS As Worksheet
Dim iMaxRowPS As Long
Dim iRowPS As Long
Dim sBankAcctPS As String
Dim sBUCodePS As String
Dim sSapGL As String
Dim sProductCodePS As String
Dim sBankNamePS As String

Dim sBankShortName As String
Dim sBankAcctLast4 As String
Dim sBankCode As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

'In sheet "FIS & PeopleSoft"
Set wsFIS = Worksheets(SheetNameFIS)
wsFIS.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowFIS = lRealLastRow
If iMaxRowFIS < 2 Then Exit Sub

iRowCurrentFIS = iMaxRowFIS


' Process PeopleSoft file and data sheet
sFileNamePS = GetWorkPath & "\" & FileNamePS
'Debug.Print sFileNameTreasury
Set wkbPS = Workbooks.Open(sFileNamePS)
wkbPS.Activate
Set wsPS = Worksheets(1)
wsPS.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowPS = lRealLastRow
If iMaxRowPS < 2 Then Exit Sub

For iRowPS = 2 To iMaxRowPS
'For iRowPS = 293 To 293
    iFound = 0
    
    wkbPS.Activate
    wsPS.Select
    
    'Process bank account (long)
    sBankAcctPS = wsPS.Cells(iRowPS, ColPSBankAcct)
    'Debug.Print sBankAcctPS
    
    sBUCodePS = wsPS.Cells(iRowPS, ColPSBUCode)
    sSapGL = wsPS.Cells(iRowPS, ColPSSapGL)
    
    sProductCodePS = wsPS.Cells(iRowPS, ColPSProductCode)
    sBankNamePS = wsPS.Cells(iRowPS, ColPSBankName)
    
   
    'to check with FIS Sheet by Bank Account
    ThisWorkbook.Activate
    wsFIS.Select
    'Debug.Print "Max Row: - " & iRowCurrentFIS
    For iRowFIS = 2 To iRowCurrentFIS
        sBankAcctFIS = wsFIS.Cells(iRowFIS, ColFISBankAcct)
    
        'if PeopleSoft's bank account number is included in Cash position report's bank account
        'then this account exist already in Cash Position report
        'Otherwise, it is new account
        'If sBankAcctPS = sBankAcctFIS Then
        If InStr(sBankAcctFIS, sBankAcctPS) > 0 Then
            'Debug.Print "FIS Row-" & iRowFIS & "-" & sBankAcctFIS
            wsFIS.Cells(iRowFIS, ColFISProductCode) = sProductCodePS
            wsFIS.Cells(iRowFIS, ColFISIsinPS) = "Y"
            
            iFound = 1
            Exit For
        End If
    Next iRowFIS
    Debug.Print iRowPS & "-" & iFound
    If iFound = 0 Then
        wkbPS.Activate
        wsPS.Select
        
        iRowCurrentFIS = iRowCurrentFIS + 1
        
        sBankAcctPS = Long_Bank_Account(sBankAcctPS)
        wsFIS.Cells(iRowCurrentFIS, ColFISBankAcct) = sBankAcctPS
        wsFIS.Cells(iRowCurrentFIS, ColFISBUCode) = sBUCodePS
        wsFIS.Cells(iRowCurrentFIS, ColFISSapGL) = sSapGL
        wsFIS.Cells(iRowCurrentFIS, ColFISProductCode) = sProductCodePS
        wsFIS.Cells(iRowCurrentFIS, ColFISIsinPS) = "Y"
        wsFIS.Cells(iRowCurrentFIS, ColFISCompanyName) = sBankNamePS
                        
        wsFIS.Cells(iRowCurrentFIS, ColFISFISCode) = ""
        
    End If
    
Next iRowPS


wkbPS.Close savechanges:=False

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

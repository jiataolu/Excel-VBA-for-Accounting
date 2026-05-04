Attribute VB_Name = "Module08_ReadKeyBankAccount"
Option Explicit


Sub Find_Key_Bank_Info_and_Account()

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim iGL As Integer
Dim dAMT As Double
Dim sComments As String
Dim sCommentsCompress As String
Dim sBank As String
Dim sBOAInfoA As String
Dim sBOAinfoB As String
Dim sKeyBankAcct As String
Dim iPosStart As Integer
Dim iLenInfo As Integer
Dim SBankInfo As String

Dim wsConcenClear As Worksheet
Dim rngCon As Range
Dim rngFound As Range

Dim wsMappingEx As Worksheet
Dim iMaxRowMappingEx As Integer
Dim iRowMappingEx As Integer
Dim sMappingExtype As String
Dim sMappingExKeyWord As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsConcenClear = Worksheets("Concentration & Clearing GL")
wsConcenClear.Select
Set rngCon = Columns(iColConcenClear)

Set wsMappingEx = Worksheets("Mapping Exceptional")
wsMappingEx.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMappingEx = lRealLastRow


Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub


For iRowItems = 2 To iMaxRowItems
'For iRowItems = 3 To 3
    sBank = ""
    iGL = CInt(Cells(iRowItems, iColItemsGL))
    dAMT = Cells(iRowItems, iColItemsAMT)
    sComments = Cells(iRowItems, iColItemsBankInfo)
    sCommentsCompress = UCase(Replace(sComments, " ", ""))
    sKeyBankAcct = ""
    
    'To find bank name by GL for each line
    Set rngFound = rngCon.Find(iGL, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        'sBank = Left(Replace(rngFound.Cells(1, 2), " ", ""), 3)
        sBank = Mid(Replace(rngFound.Cells(1, 2), " ", ""), 5, 2)
        Debug.Print sBank
    End If
    
    'When bank is BOA, GL=10301
    'If UCase(sBank) = "BOA" Then
    If UCase(sBank) = "BA" Then
        'To decide ORIG or BNF info and key bank account
        If dAMT > 0 Then   'when dAMT<0 regualr BOA description ORIG:BNF, BNF is coding party
            'First information, which also decide key bank account for coding

            sBOAInfoA = BOA_BNF_Info(sComments)
            sKeyBankAcct = BOA_BNF_Bank_Account(sBOAInfoA)
            
            'Second information
            sBOAinfoB = BOA_ORIG_Info_B(sComments)
            
        Else    'when dAMT<0, regualr BOA description ORIG:ORG, orig is coding party
            'First information, which also decide key bank account for coding
            sBOAInfoA = BOA_ORIG_Info(sComments)
            sKeyBankAcct = BOA_ORIG_Bank_Account(sBOAInfoA)
            
            'Second information (only highlight information, without need to find key bank accout)
            sBOAinfoB = BOA_BNF_Info(sComments)
            
        End If
        
        
        iPosStart = InStr(sComments, sBOAInfoA)
        iLenInfo = Len(sBOAInfoA)
        If iLenInfo > 0 Then
            Cells(iRowItems, iColItemsBankInfo).Characters(Start:=iPosStart, Length:=iLenInfo).font.ColorIndex = 3
        End If
        
        iPosStart = InStr(sComments, sBOAinfoB)
        iLenInfo = Len(sBOAinfoB)
        'Debug.Print iPosStart
        If iLenInfo > 0 Then
            Cells(iRowItems, iColItemsBankInfo).Characters(Start:=iPosStart, Length:=iLenInfo).font.ColorIndex = 46
        End If
        
        Cells(iRowItems, iColItemsKeyBankAccount) = "'" & sKeyBankAcct
        
                
        
        'if bank info is simply as (TRSF FR 1291540794)
        If Replace(sKeyBankAcct, " ", "") = "" Then
            sKeyBankAcct = Number_String_Only_First(sComments)
            Cells(iRowItems, iColItemsKeyBankAccount) = sKeyBankAcct
        End If
    'The end of BOA case
    End If
    
    'when bank is JPM, GL=10320
    'If UCase(sBank) = "JPM" Then
    If UCase(sBank) = "JP" Then
        'Debug.Print iRowItems
        
        SBankInfo = JPM_Info(sComments)
        sKeyBankAcct = JPM_Bank_Account(SBankInfo)
        
        Debug.Print "SBankInfo" & "-" & SBankInfo
        'Debug.Print "sKeyBankAcct" & "-" & sKeyBankAcct
        
        iPosStart = InStr(sComments, SBankInfo)
        iLenInfo = Len(SBankInfo)
        If iLenInfo > 0 Then
            Cells(iRowItems, iColItemsBankInfo).Characters(Start:=iPosStart, Length:=iLenInfo).font.ColorIndex = 3
        End If
        
        Cells(iRowItems, iColItemsKeyBankAccount) = "'" & sKeyBankAcct

    'end of JPM
    End If
    
    'when bank is FTB, GL=10322
    'If UCase(sBank) = "FTB" Then
    If UCase(sBank) = "FT" Then
    
    'end of FTB
    End If

    'when bank is USB GL=10326
    'If UCase(sBank) = "USB" Then
    If UCase(sBank) = "UB" Then
        SBankInfo = USB_Info(sComments)

        sKeyBankAcct = USB_Bank_Account(SBankInfo)
        
        iPosStart = InStr(sComments, SBankInfo)
        iLenInfo = Len(SBankInfo)
        If iLenInfo > 0 Then
            Cells(iRowItems, iColItemsBankInfo).Characters(Start:=iPosStart, Length:=iLenInfo).font.ColorIndex = 3
        End If
        
        Cells(iRowItems, iColItemsKeyBankAccount) = "'" & sKeyBankAcct

    'end of USB
    End If

    'when bank is WFB GL=10327
    'If UCase(sBank) = "WFB" Then
    If UCase(sBank) = "WF" Then
        SBankInfo = WFB_Info(sComments)
        sKeyBankAcct = WFB_Bank_Account(SBankInfo)
        
        iPosStart = InStr(sComments, SBankInfo)
        iLenInfo = Len(SBankInfo)
        If iLenInfo > 0 Then
            Cells(iRowItems, iColItemsBankInfo).Characters(Start:=iPosStart, Length:=iLenInfo).font.ColorIndex = 3
        End If
        
        Cells(iRowItems, iColItemsKeyBankAccount) = "'" & sKeyBankAcct
    
    'end of WFB
    End If
    
Next iRowItems

'check key word
For iRowItems = 2 To iMaxRowItems
'For iRowItems = 4 To 4
    'sBank = ""
    'iGL = CInt(Cells(iRowItems, iColItemsGL))
    'dAMT = Cells(iRowItems, iColItemsAMT)
    sComments = Cells(iRowItems, iColItemsBankInfo)
    sCommentsCompress = UCase(Replace(sComments, " ", ""))
    'sKeyBankAcct = Cells(iRowItems, iColItemsKeyBankAccount)
    Debug.Print sComments
    
    For iRowMappingEx = 2 To iMaxRowMappingEx
    'For iRowMappingEx = 24 To 24
        sMappingExtype = wsMappingEx.Cells(iRowMappingEx, iColMapType)
        If UCase(Replace(sMappingExtype, " ", "")) = "KEYWORD" Then
            sMappingExKeyWord = wsMappingEx.Cells(iRowMappingEx, iColMapBankAcct)
            'Debug.Print sMappingExKeyWord
            
            If Bank_Info_Has_KeyWord(sComments, sMappingExKeyWord) Then
                'Debug.Print "Rresult:"
                'Debug.Print "To write Keyword:" & sMappingExKeyWord
                Cells(iRowItems, iColItemsKeyBankAccount) = sMappingExKeyWord
                Exit For
            End If
        
        End If
    Next iRowMappingEx
    
Next iRowItems



Set rngCon = Nothing
Set rngFound = Nothing
End Sub


Function Bank_Info_Has_KeyWord(BankInfo, KeyWord)
Bank_Info_Has_KeyWord = False
'Debug.Print "Start"
'Debug.Print KeyWord
'Keyword is combination, which as "[XXXX]"
If InStr(KeyWord, "[") > 0 And InStr(KeyWord, "]") > 0 Then
    'Debug.Print "Step 1"
    Bank_Info_Has_KeyWord = Bank_Info_Has_KeyWord_Regex(BankInfo, KeyWord)
    'Debug.Print "Sub Result:" & Bank_Info_Has_KeyWord
    
'Keyword is regular, not combination
Else
    If InStr(UCase(Replace(BankInfo, " ", "")), UCase(Replace(KeyWord, " ", ""))) > 0 Then
        'Debug.Print "Yes, regular keyword"
        Bank_Info_Has_KeyWord = True
    End If
End If

End Function




'Sub regex()
Function Bank_Info_Has_KeyWord_Regex(BankInfo, KeyWord)
Bank_Info_Has_KeyWord_Regex = False

'Debug.Print "Step 2"

Dim objRegex As Object
Dim strPattern As String
Dim bolTest As Boolean
Dim objMatches As Object
Dim objMatch As Object



Dim strMyString As String
Dim strSubKeyWord As String
Dim iFound As Integer


strMyString = KeyWord
'strMyString = "The adventure of Batwoman & The adventure of Batman & The adventure of BatWOwowowoman "
'strMyString = "[ALIGHT SOLUTIONS] [ID:0004217685]"

'Debug.Print "Lu:"
'strPattern = "Bat(wo)+man"
strPattern = "\[.*?\]"


Set objRegex = CreateObject("VBScript.RegExp")
objRegex.Global = True
objRegex.IgnoreCase = True
objRegex.Pattern = strPattern
bolTest = objRegex.test(strMyString)
Set objMatches = objRegex.Execute(strMyString)

If bolTest = True Then
    iFound = 1
    For Each objMatch In objMatches
        'Debug.Print objMatch.Value
        'Debug.Print objMatch.Value
        strSubKeyWord = UCase(Replace(objMatch.Value, " ", ""))
        strSubKeyWord = Replace(strSubKeyWord, "[", "")
        strSubKeyWord = Replace(strSubKeyWord, "]", "")
        
        If InStr(UCase(Replace(BankInfo, " ", "")), strSubKeyWord) > 0 Then
            'Debug.Print "BankInfo:" & UCase(Replace(BankInfo, " ", ""))
            'Debug.Print "Keyword:" & strSubKeyWord
            iFound = iFound * 1
        Else
            iFound = iFound * 0
        End If
    Next objMatch
End If
'Debug.Print "Ifound:" & iFound
If iFound = 1 Then Bank_Info_Has_KeyWord_Regex = True

'Debug.Print "Hello regex keyword"
End Function

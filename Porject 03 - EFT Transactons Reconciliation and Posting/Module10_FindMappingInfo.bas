Attribute VB_Name = "Module10_FindMappingInfo"
Option Explicit

Sub Find_Mapping_Info_Step1()

'this sub is to find BU, GL from "Mapping Exceptional" sheet and mapping file.
Call Initialize_Items_Sheet

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sKeyBankAcct As String
Dim sPostBU As String
Dim sPostGL As String
Dim sPostVendor As String
Dim sPostCurrency As String

Dim wsMapExc As Worksheet
Dim rngExcBankAcct As Range


Dim wkbMap As Workbook
Dim wsMap As Worksheet
Dim iMaxRowMap As Integer
Dim iRowMap As Integer
Dim rngBankAcct As Range
Dim rngFound As Range


Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wkbMap = Workbooks.Open(Map_File_Full_Name())
wkbMap.Activate
Set wsMap = Worksheets("Mapping EFT")
wsMap.Select
Set rngBankAcct = Columns(iColMapBankAcct)



ThisWorkbook.Activate
Set wsMapExc = Worksheets("Mapping Exceptional")
wsMapExc.Select
Set rngExcBankAcct = Columns(iColMapExcBankAcct)


Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub


For iRowItems = 2 To iMaxRowItems
'For iRowItems = 4 To 4
    sKeyBankAcct = Cells(iRowItems, iColItemsKeyBankAccount)
    sKeyBankAcct = Number_wt_Leading_Zero(sKeyBankAcct)
    sKeyBankAcct = Replace(sKeyBankAcct, "-", "")
    Debug.Print "BankAcct=" & sKeyBankAcct
    
    'To search in Mapping Exceptional sheet
    Set rngFound = rngExcBankAcct.Find(sKeyBankAcct, LookIn:=xlValues, lookat:=xlWhole)
        
    If Not rngFound Is Nothing Then
        
        sPostBU = rngFound.Cells(1, 2)
        sPostGL = rngFound.Cells(1, 3)
        sPostVendor = rngFound.Cells(1, 4)
        
        wsItems.Cells(iRowItems, iColItemsPostBU) = sPostBU
        wsItems.Cells(iRowItems, iColItemsPostGL) = sPostGL
        wsItems.Cells(iRowItems, iColItemsPostVendor) = sPostVendor
        
    'If bank account is not in Mapping Exceptional Sheet, then search mapping file
    Else
        'Debug.Print "hELLO"
        Set rngFound = rngBankAcct.Find(sKeyBankAcct, LookIn:=xlValues, lookat:=xlWhole)
        If Not rngFound Is Nothing Then
            sPostBU = rngFound.Cells(1, iColMapBU - iColMapBankAcct + 1)
            sPostGL = rngFound.Cells(1, iColMapGL - iColMapBankAcct + 1)
            sPostCurrency = rngFound.Cells(1, iColMapCurrency - iColMapBankAcct + 1)
        
            wsItems.Cells(iRowItems, iColItemsPostBU) = sPostBU
            wsItems.Cells(iRowItems, iColItemsPostGL) = sPostGL
            
            If UCase(Replace(sPostCurrency, " ", "")) <> "USD" Then wsItems.Cells(iRowItems, iColItemsPostCurrency) = sPostCurrency
            
            
            Debug.Print sKeyBankAcct
            Debug.Print "sPostBU=" & sPostBU
            Debug.Print "sPostGL=" & sPostGL
        End If
    End If
    
    'if KeyBank has "Vendor:"
    'If InStr(UCase(sKeyBankAcct), "VENDOR:") > 0 Then

     '   sKeyBankAcct = UCase(sKeyBankAcct)
      '  sKeyBankAcct = Replace(sKeyBankAcct, "VENDOR:", "")
       ' sKeyBankAcct = Replace(sKeyBankAcct, " ", "")
        
        'If InStr(sKeyBankAcct, "-") > 0 Then
         '   sPostBU = Left(sKeyBankAcct, InStr(sKeyBankAcct, "-") - 1)
          '  sPostGL = Right(sKeyBankAcct, Len(sKeyBankAcct) - InStr(sKeyBankAcct, "-"))
            
           ' wsItems.Cells(iRowItems, iColItemsPostBU) = sPostBU
            'wsItems.Cells(iRowItems, iColItemsPostVendor) = sPostGL
        'End If
        
    'End If
    
    
    
Next


wkbMap.Close SaveChanges:=False

Set wsItems = Nothing
Set wkbMap = Nothing
Set wsMap = Nothing
Set rngExcBankAcct = Nothing
Set rngBankAcct = Nothing
Set rngFound = Nothing
End Sub


Sub Find_Mapping_Info_Step2()

'This function is to find the coding party which is already in Items sheet
Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sPostBU As String
Dim sPostGL As String
Dim dAMT As Double
Dim iSecondRowItems As Integer


Dim wsConClear As Worksheet
Dim rngGL As Range
Dim rngFound As Range

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsConClear = Worksheets("Concentration & Clearing GL")
wsConClear.Select
Set rngGL = Columns(iColConcenClear)

Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub

For iRowItems = 2 To iMaxRowItems - 1
'For iRowItems = 2 To 2
    sPostBU = Cells(iRowItems, iColItemsPostBU)
    sPostGL = Cells(iRowItems, iColItemsPostGL)
    dAMT = Cells(iRowItems, iColItemsAMT)
    
    'to check only BU 9000
    If sPostBU <> "9000" Then GoTo GOTONextLine
    
    'If BU is 9000, then to check if it is posted to another concentration account
    Set rngFound = rngGL.Find(sPostGL, LookIn:=xlValues, lookat:=xlWhole)
    If rngFound Is Nothing Then GoTo GOTONextLine
    
    'Debug.Print sBU
    'Debug.Print sGL
    'Debug.Print dAMT
    For iSecondRowItems = iRowItems + 1 To iMaxRowItems
    'For iSecondRowItems = 18 To 18
     If sPostGL = Cells(iSecondRowItems, iColItemsGL) And dAMT = Cells(iSecondRowItems, iColItemsAMT) * (-1) Then
            'Debug.Print iSecondRowItems
            Cells(iSecondRowItems, iColItemsPostBU) = "See Row " & CStr(iRowItems)
            Cells(iSecondRowItems, iColItemsPostGL) = "See Row " & CStr(iRowItems)
    End If
        
    Next iSecondRowItems

GOTONextLine:
Next iRowItems

Cells(1, 1).Select

Set wsConClear = Nothing
Set wsItems = Nothing
Set rngGL = Nothing
Set rngFound = Nothing

End Sub


Sub Find_Mapping_Info_Step3()
'This function is to search profit center
Dim sProfitCFileFullName As String
Dim wkbProfitC As Workbook
Dim wsProfitC As Worksheet
Dim iMaxRowProfitC As Integer
Dim iRowProfitC As Integer
Dim sPCBU As String
Dim sPCGL As String
Dim sPCProfitCenter As String

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sItemsPostBU As String
Dim sItemsPostGL As String

Dim iFound As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

sProfitCFileFullName = GetWorkPath() & "\" & SubFolderMapping & "\" & "Profit Center.xlsx"
'Debug.Print sProfitCFileFullName

Set wkbProfitC = Workbooks.Open(sProfitCFileFullName)
wkbProfitC.Activate
Set wsProfitC = Worksheets(1)
wsProfitC.Select

lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowProfitC = lRealLastRow
If iMaxRowProfitC < 2 Then Exit Sub

ThisWorkbook.Activate
Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub

For iRowItems = 2 To iMaxRowItems
'For iRowItems = 14 To 14
    sPCProfitCenter = ""
    
    sItemsPostBU = wsItems.Cells(iRowItems, iColItemsPostBU)
    sItemsPostGL = wsItems.Cells(iRowItems, iColItemsPostGL)
    
    'If BU and GL are both non-empty
    If Replace(sItemsPostBU, " ", "") <> "" And Replace(sItemsPostGL, " ", "") <> "" Then
    
        For iRowProfitC = 2 To iMaxRowProfitC
            sPCBU = wsProfitC.Cells(iRowProfitC, iColProfitCBU)
            sPCGL = wsProfitC.Cells(iRowProfitC, iColProfitCGL)
        
            If sItemsPostBU = sPCBU And sItemsPostGL = sPCGL Then

                sPCProfitCenter = wsProfitC.Cells(iRowProfitC, iColProfitCPC)
                wsItems.Cells(iRowItems, iColItemsPostProfitC) = sPCProfitCenter
                Exit For
            End If
        Next iRowProfitC
    End If
    
Next iRowItems



wkbProfitC.Close SaveChanges:=False

Set wsItems = Nothing
Set wkbProfitC = Nothing
Set wsProfitC = Nothing

End Sub



Sub Find_Mapping_Info_Step4()

'This function is to find post key - 40 for debit 50 for credit, 21 or 31 for vendor code
Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sPostBU As String
Dim sPostGL As String
Dim sPostVendor As String
Dim dAMT As Double
Dim sPostKeyCode As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow
If iMaxRowItems < 2 Then Exit Sub

'For iRowItems = 3 To 3
For iRowItems = 2 To iMaxRowItems
    sPostKeyCode = ""
    sPostBU = Cells(iRowItems, iColItemsPostBU)
    sPostGL = Cells(iRowItems, iColItemsPostGL)
    sPostVendor = Cells(iRowItems, iColItemsPostVendor)
    dAMT = Cells(iRowItems, iColItemsAMT)
    
    If Replace(sPostBU, " ", "") = "" Then GoTo DONEXTLINE
    
    If dAMT < 0 Then
        sPostKeyCode = 50
    Else
        sPostKeyCode = 40
    End If
    
    'if it is a vendor code
    If Replace(sPostVendor, " ", "") <> "" Then
        If dAMT < 0 Then
            sPostKeyCode = 31
        Else
            sPostKeyCode = 21
        End If
    End If
    
    Cells(iRowItems, iColItemsPostKeyCode) = sPostKeyCode

DONEXTLINE:
Next iRowItems

Cells(1, 1).Select

Set wsItems = Nothing

End Sub

Sub Find_Mapping_Info_Step5_Email_to_Confirm()

Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
'Dim iGL As Integer
'Dim dAMT As Double
Dim sComments As String
'Dim sCommentsCompress As String
'Dim sBank As String
'Dim sBOAInfoA As String
'Dim sBOAinfoB As String
'Dim sKeyBankAcct As String
'Dim iPosStart As Integer
'Dim iLenInfo As Integer
'Dim SBankInfo As String
Dim sPostBU As String
Dim sPostGL As String
Dim sPostVendor As String
Dim sPostKeyCode As String
Dim sPostProfitCenter As String



Dim wsMappingEx As Worksheet
Dim iMaxRowMappingEx As Integer
Dim iRowMappingEx As Integer
Dim sMappingExtype As String
Dim sMappingExKeyWord As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long


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
    'sBank = ""
    'iGL = CInt(Cells(iRowItems, iColItemsGL))
    'dAMT = Cells(iRowItems, iColItemsAMT)
    sComments = Cells(iRowItems, iColItemsBankInfo)
    'sCommentsCompress = UCase(Replace(sComments, " ", ""))
    'sKeyBankAcct = Cells(iRowItems, iColItemsKeyBankAccount)
    'Debug.Print sComments
    
    For iRowMappingEx = 2 To iMaxRowMappingEx
    'For iRowMappingEx = 24 To 24
        sMappingExtype = wsMappingEx.Cells(iRowMappingEx, iColMapType)
        If UCase(Replace(sMappingExtype, " ", "")) = "EMAILCONFIRM" Then
            sMappingExKeyWord = wsMappingEx.Cells(iRowMappingEx, iColMapBankAcct)
            'Debug.Print sMappingExKeyWord
            
            If Bank_Info_Has_KeyWord(sComments, sMappingExKeyWord) Then
                'Debug.Print "Rresult:"
                'Debug.Print "To write Keyword:" & sMappingExKeyWord
                'Cells(iRowItems, iColItemsKeyBankAccount) = sMappingExKeyWord
                
                
                sPostBU = Cells(iRowItems, iColItemsPostBU)
                sPostGL = Cells(iRowItems, iColItemsPostGL)
                sPostVendor = Cells(iRowItems, iColItemsPostVendor)
                sPostProfitCenter = Cells(iRowItems, iColItemsPostProfitC)
                sPostKeyCode = Cells(iRowItems, iColItemsPostKeyCode)
                
                If sPostBU = "" Then
                    sPostBU = WaitToConfirmInfo
                Else
                    sPostBU = sPostBU & vbCrLf & WaitToConfirmInfo
                End If
                Cells(iRowItems, iColItemsPostBU) = sPostBU
        
                If sPostGL = "" Then
                    sPostGL = WaitToConfirmInfo
                Else
                    sPostGL = sPostGL & vbCrLf & WaitToConfirmInfo
                End If
                Cells(iRowItems, iColItemsPostGL) = sPostGL
        
                If sPostVendor = "" Then
                    sPostVendor = WaitToConfirmInfo
                Else
                    sPostVendor = sPostVendor & vbCrLf & WaitToConfirmInfo
                End If
                Cells(iRowItems, iColItemsPostVendor) = sPostVendor
        
                If sPostKeyCode = "" Then
                    sPostKeyCode = WaitToConfirmInfo
                Else
                    sPostKeyCode = sPostKeyCode & vbCrLf & WaitToConfirmInfo
                End If
                Cells(iRowItems, iColItemsPostKeyCode) = sPostKeyCode
        
                If sPostProfitCenter = "" Then
                    sPostProfitCenter = WaitToConfirmInfo
                Else
                    sPostProfitCenter = sPostProfitCenter & vbCrLf & WaitToConfirmInfo
                End If
                Cells(iRowItems, iColItemsPostProfitC) = sPostProfitCenter
                
                Exit For
            End If
        
        End If
    Next iRowMappingEx
    
Next iRowItems

End Sub

Sub Find_Mapping_Info_Step6_Format()
Dim wsItems As Worksheet
Set wsItems = Worksheets("2-Items to post")
wsItems.Select
Cells.Select
Selection.Columns.AutoFit
Cells(1, 7).Select
End Sub





Sub Initialize_Items_Sheet()
Dim wsItems As Worksheet
Dim lRealLastRow As Long
Dim lRealLastCol As Long
Dim iMaxRowItems As Integer

Set wsItems = Worksheets("2-Items to post")
wsItems.Select

Columns(iColItemsPostBU).Select
Selection.ClearContents
Selection.Interior.Pattern = xlNone
Selection.font.ColorIndex = xlAutomatic
Selection.HorizontalAlignment = xlCenter

Columns(iColItemsPostGL).Select
Selection.ClearContents
Selection.Interior.Pattern = xlNone
Selection.font.ColorIndex = xlAutomatic
Selection.HorizontalAlignment = xlCenter

Columns(iColItemsPostProfitC).Select
Selection.ClearContents
Selection.Interior.Pattern = xlNone
Selection.font.ColorIndex = xlAutomatic
Selection.HorizontalAlignment = xlCenter

Columns(iColItemsPostAssInfo).Select
Selection.ClearContents
Selection.Interior.Pattern = xlNone
Selection.font.ColorIndex = xlAutomatic
Selection.HorizontalAlignment = xlCenter


Columns(iColItemsPostCostCenter).Select
Selection.ClearContents
Selection.Interior.Pattern = xlNone
Selection.font.ColorIndex = xlAutomatic
Selection.HorizontalAlignment = xlCenter


Cells(1, iColItemsPostBU) = "BU"
Cells(1, iColItemsPostGL) = "GL"
Cells(1, iColItemsPostVendor) = "Vendor"
Cells(1, iColItemsPostProfitC) = "Profit Center"
Cells(1, iColItemsPostKeyCode) = "Key Code"
Cells(1, iColItemsPostAssInfo) = "Assignment"
Cells(1, iColItemsPostCostCenter) = "Cost Center"
Cells(1, iColItemsPostBU).Select


lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow

With Range(Cells(1, iColItemsPostBU), Cells(iMaxRowItems, iColItemsPostCostCenter)).Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = 0.799981688894314
    .PatternTintAndShade = 0
End With

Set wsItems = Nothing


End Sub



Attribute VB_Name = "Module02_FindOffsetItems"
Option Explicit

Public avTextOnly() As Variant
'       0           1           2
'   0   GL:
'   1   Amount:
'   2   If Offset

Public avPartGroup() As Variant
'       0           1           2           3
'   0   GL:
'   1   Leading_Number:
'   2   Amount:
'   3   If Offset:


Sub Find_Offset_Items()
'Empty_Assignment_NonEmpty_Text - Text Only

Call Create_Array_for_Text_Only_and_Part_Group
''Call show_Array_Value
Call Match_Both_Empty_Assignment_Text_to_Single_Line_Transaction_Both_NonEmpty_Assignment_Text
Call Match_Both_Empty_Assignment_Text_to_PART_Group
Call Match_Empty_Assignment_NonEmpty_Text_to_Signle_Line_Transaction_Both_NonEmpty_Assginment_Text
Call Match_Empty_Assignment_NonEmpty_Text_to_PART_Group

End Sub



'Step-1: Create Array for both Text Only transaction Group, and PART Group,as public variant
Sub Create_Array_for_Text_Only_and_Part_Group()

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim sGL As String
Dim sAss As String
Dim sText As String
Dim dAMT As Double
Dim sGLCurrent As String
Dim sLeadingNumber As String



'Dim avTextOnly() As Variant
Dim sGLTextOnly As String
Dim sAMTTextOnly As String
Dim iCountArrayTextOnly As Integer
Dim iFoundTextOnly As Integer
Dim iArrayTextOnly As Integer

'Dim avPartGroup() As Variant
Dim iCountArrayPartGroup As Variant
Dim iFoundLeadingNumber As Integer
Dim iArrayPartGroup As Integer


Dim lRealLastRow As Long
Dim lRealLastCol As Long


Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow
If iMaxRowSAP < 2 Then Exit Sub

ReDim Preserve avTextOnly(2, 0)
avTextOnly(0, 0) = "GL"
avTextOnly(1, 0) = 0#
avTextOnly(2, 0) = "If Offset"

ReDim Preserve avPartGroup(3, 0)
avPartGroup(0, 0) = "GL"
avPartGroup(1, 0) = "First 4 Number"
avPartGroup(2, 0) = 0#
avPartGroup(3, 0) = "If Offset"
Debug.Print LBound(avPartGroup, 2)
Debug.Print UBound(avPartGroup, 2)

iCountArrayTextOnly = 1
iCountArrayPartGroup = 1
sGLCurrent = ""


For iRowSAP = 2 To iMaxRowSAP
'For iRowSAP = 2 To 2
    sGL = Cells(iRowSAP, iColSAPGL)
    sAss = Cells(iRowSAP, iColSAPAss)
    sText = Cells(iRowSAP, iColSAPText)
    dAMT = Cells(iRowSAP, iColSAPAMT)
    
    'Text only transactions
    If Replace(sAss, " ", "") = "" And Replace(sText, " ", "") <> "" Then
        iFoundTextOnly = 0
                    
        For iArrayTextOnly = 0 To UBound(avTextOnly, 2)
            sGLTextOnly = avTextOnly(0, iArrayTextOnly)
            
            If sGL = sGLTextOnly Then
                iFoundTextOnly = 1
                Exit For
            End If
        Next iArrayTextOnly
        
        If iFoundTextOnly = 1 Then
            avTextOnly(1, iArrayTextOnly) = avTextOnly(1, iArrayTextOnly) + dAMT
        Else
            ReDim Preserve avTextOnly(2, iCountArrayTextOnly)
            avTextOnly(0, iCountArrayTextOnly) = sGL
            avTextOnly(1, iCountArrayTextOnly) = dAMT
            avTextOnly(2, iCountArrayTextOnly) = "Empty"
            iCountArrayTextOnly = iCountArrayTextOnly + 1
        End If
    End If
    
    'PART Group transactions
    If Replace(sAss, " ", "") <> "" And InStr(Replace(UCase(sText), " ", ""), "PART") > 0 And InStr(Replace(UCase(sText), " ", ""), "OF") > 0 Then
        
        sLeadingNumber = Leading_Number(sText)
        iFoundLeadingNumber = 0
        
        For iArrayPartGroup = 0 To UBound(avPartGroup, 2)
            If sGL = avPartGroup(0, iArrayPartGroup) And sLeadingNumber = avPartGroup(1, iArrayPartGroup) Then
                iFoundLeadingNumber = 1
                avPartGroup(2, iArrayPartGroup) = avPartGroup(2, iArrayPartGroup) + dAMT
                Exit For
           End If
        Next iArrayPartGroup
        
        If iFoundLeadingNumber = 0 Then
            ReDim Preserve avPartGroup(3, iCountArrayPartGroup)
            avPartGroup(0, iCountArrayPartGroup) = sGL
            avPartGroup(1, iCountArrayPartGroup) = sLeadingNumber
            avPartGroup(2, iCountArrayPartGroup) = dAMT
            avPartGroup(3, iCountArrayPartGroup) = "Empty"
            iCountArrayPartGroup = iCountArrayPartGroup + 1
        End If
    End If
    
    
Next iRowSAP

Dim iArray As Integer
For iArray = LBound(avTextOnly, 2) To UBound(avTextOnly, 2)
    Debug.Print iArray & "--" & avTextOnly(0, iArray) & "--" & avTextOnly(1, iArray) & "--" & avTextOnly(2, iArray)
Next iArray

For iArray = LBound(avPartGroup, 2) To UBound(avPartGroup, 2)
    Debug.Print iArray & "--" & avPartGroup(0, iArray) & "--" & avPartGroup(1, iArray) & "--" & avPartGroup(2, iArray) & "--" & avPartGroup(3, iArray)
Next iArray

End Sub


'step-2: Match transactions both empty for Assignment Field and Text Field
'with single line transation
Sub Match_Both_Empty_Assignment_Text_to_Single_Line_Transaction_Both_NonEmpty_Assignment_Text()

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim sGL As String
Dim sAss As String
Dim sText As String
Dim dAMT As Double
Dim sClearNote As String

Dim iRowSAPSecond As Integer
Dim sGLSecond As String
Dim sAssSecond As String
Dim sTextSecond As String
Dim dAMTSecond As Double
Dim sClearNoteSecond As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow
If iMaxRowSAP < 2 Then Exit Sub


For iRowSAP = 2 To iMaxRowSAP - 1
    sGL = Cells(iRowSAP, iColSAPGL)
    sAss = Cells(iRowSAP, iColSAPAss)
    sText = Cells(iRowSAP, iColSAPText)
    dAMT = Cells(iRowSAP, iColSAPAMT)
    sClearNote = Cells(iRowSAP, iColSAPClear)
    
    'Both Emptry for Assignment Field and Text Field, and not offeset yet
    If Replace(sAss, " ", "") = "" And Replace(sText, " ", "") = "" And Replace(sClearNote, " ", "") = "" Then
        
        For iRowSAPSecond = iRowSAP + 1 To iMaxRowSAP
                sGLSecond = Cells(iRowSAPSecond, iColSAPGL)
                sAssSecond = Cells(iRowSAPSecond, iColSAPAss)
                sTextSecond = Cells(iRowSAPSecond, iColSAPText)
                dAMTSecond = Cells(iRowSAPSecond, iColSAPAMT)
                sClearNoteSecond = Cells(iRowSAPSecond, iColSAPClear)
                
                ' Match with both non-empty for Assignment Field and Text Field, and not offeset yet
                'If Replace(sAssSecond, " ", "") <> "" And Replace(sTextSecond, " ", "") <> "" And Replace(sClearNoteSecond, " ", "") = "" Then
                If sGL = sGLSecond And Replace(sAssSecond, " ", "") <> "" And Replace(sTextSecond, " ", "") <> "" And Replace(sClearNoteSecond, " ", "") = "" Then
                    If Abs(dAMT + dAMTSecond) < 0.001 Then
                        Cells(iRowSAP, iColSAPClear) = "Offset"
                        Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = 15
                        Cells(iRowSAPSecond, iColSAPClear) = "Offset"
                        Range(Cells(iRowSAPSecond, 1), Cells(iRowSAPSecond, iColSAPPostKey)).Interior.ColorIndex = 15
                        
                        Exit For
                    End If
                    
                End If

        Next iRowSAPSecond
        
    End If
    
Next iRowSAP
End Sub


'step-3: Match transactions both empty for Assignment Field and Text Field
'but with PART Group
Sub Match_Both_Empty_Assignment_Text_to_PART_Group()

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim sGL As String
Dim sAss As String
Dim sText As String
Dim dAMT As Double
Dim sClearNote As String

Dim sGLPartGroup As String
Dim sLeadingNumberPartGroup As String
Dim dAMTPartGroup As Double
Dim sIfOffsetPartGroup As String
Dim iArray As Integer


Dim iRowSAPSecond As Integer
Dim sGLSecond As String
Dim sAssSecond As String
Dim sTextSecond As String
Dim dAMTSecond As Double
Dim sClearNoteSecond As String
Dim sLeadingNumberSecond As String

Dim iFound As Integer
Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow
If iMaxRowSAP < 2 Then Exit Sub

'Debug.Print LBound(avPartGroup, 1)
'Debug.Print UBound(avPartGroup, 1)

'If PART Group has no value, then exit (the first value in avPartGroup is ('GL", 0), so value acount is from "1"
If UBound(avPartGroup, 2) < 1 Then Exit Sub

For iRowSAP = 2 To iMaxRowSAP - 1
'For iRowSAP = 31 To 31
    sGL = Cells(iRowSAP, iColSAPGL)
    sAss = Cells(iRowSAP, iColSAPAss)
    sText = Cells(iRowSAP, iColSAPText)
    dAMT = Cells(iRowSAP, iColSAPAMT)
    sClearNote = Cells(iRowSAP, iColSAPClear)
    iFound = 0
    'Debug.Print dAMT
    
    'Both Emptry for Assignment Field and Text Field, and not offeset yet
    If Replace(sAss, " ", "") = "" And Replace(sText, " ", "") = "" And Replace(sClearNote, " ", "") = "" Then
        For iArray = 1 To UBound(avPartGroup, 2)
            sGLPartGroup = avPartGroup(0, iArray)
            sLeadingNumberPartGroup = avPartGroup(1, iArray)
            dAMTPartGroup = avPartGroup(2, iArray)
            sIfOffsetPartGroup = avPartGroup(3, iArray)
            
            If sGLPartGroup = sGL And Abs(dAMTPartGroup + dAMT) < 0.001 And sIfOffsetPartGroup <> "Offset" Then
                avPartGroup(3, iArray) = "Offset"
                iFound = 1
                Exit For
            End If
            
        Next iArray
    
        If iFound = 1 Then
            Cells(iRowSAP, iColSAPClear) = "Offset"
            Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = 15
            
            'Find transactions in PART Group, then mark "Offset"
            For iRowSAPSecond = 2 To iMaxRowSAP
                sGLSecond = Cells(iRowSAPSecond, iColSAPGL)
                sAssSecond = Cells(iRowSAPSecond, iColSAPAss)
                sTextSecond = Cells(iRowSAPSecond, iColSAPText)
                dAMTSecond = Cells(iRowSAPSecond, iColSAPAMT)
                sClearNoteSecond = Cells(iRowSAPSecond, iColSAPClear)
                sLeadingNumberSecond = Leading_Number(sTextSecond)
                
                If sGLSecond = sGLPartGroup And Replace(sAssSecond, " ", "") <> "" And sLeadingNumberSecond = sLeadingNumberPartGroup And Replace(sClearNoteSecond, " ", "") = "" Then
                    Cells(iRowSAPSecond, iColSAPClear) = "Offset"
                    Range(Cells(iRowSAPSecond, 1), Cells(iRowSAPSecond, iColSAPPostKey)).Interior.ColorIndex = 15
                End If
            Next iRowSAPSecond
        End If
    End If
    
Next iRowSAP

End Sub


'Step-4: Match transactions Text Only (Empty in Assignment Field, Non-Empty in Text Field)
'with single line transaction
Sub Match_Empty_Assignment_NonEmpty_Text_to_Signle_Line_Transaction_Both_NonEmpty_Assginment_Text()

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim sGL As String
Dim sAss As String
Dim sText As String
Dim dAMT As Double
Dim sClearNote As String

Dim iRowSAPSecond As Integer
Dim sGLSecond As String
Dim sAssSecond As String
Dim sTextSecond As String
Dim sClearNoteSecond As String


Dim sGLTextOnly As String
Dim dAMTTextOnly As Double
Dim sIfOffsetTextOnly As String
Dim iArray As Integer

Dim iFound As Integer
Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow
If iMaxRowSAP < 2 Then Exit Sub

If UBound(avTextOnly, 2) < 1 Then Exit Sub

For iArray = 1 To UBound(avTextOnly, 2)
    sGLTextOnly = avTextOnly(0, iArray)
    dAMTTextOnly = avTextOnly(1, iArray)
    Debug.Print dAMTTextOnly
    sIfOffsetTextOnly = avTextOnly(2, iArray)
    If sIfOffsetTextOnly = "Offset" Then GoTo CONTINUENEXTARRAY
    
    'Search for transaction non-Empty for Assignment Field and Text Field
    For iRowSAP = 2 To iMaxRowSAP
    'For iRowSAP = 57 To 58
        sGL = Cells(iRowSAP, iColSAPGL)
        sAss = Cells(iRowSAP, iColSAPAss)
        sText = Cells(iRowSAP, iColSAPText)
        dAMT = Cells(iRowSAP, iColSAPAMT)
        sClearNote = Cells(iRowSAP, iColSAPClear)
        iFound = 0
        
        If sGLTextOnly = sGL And Abs(dAMTTextOnly + dAMT) < 0.001 Then
            avTextOnly(2, iArray) = "Offset"
            iFound = 1
            Debug.Print iRowSAP & "--" & iFound
            Exit For
        End If
    Next iRowSAP
    
    If iFound = 1 Then
        Cells(iRowSAP, iColSAPClear) = "Offset"
        Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = 15
        
        'find all text only transaction, then mark "Offset"
        For iRowSAPSecond = 2 To iMaxRowSAP
            sGLSecond = Cells(iRowSAPSecond, iColSAPGL)
            sAssSecond = Cells(iRowSAPSecond, iColSAPAss)
            sTextSecond = Cells(iRowSAPSecond, iColSAPText)
            sClearNoteSecond = Cells(iRowSAPSecond, iColSAPClear)
            
            If sGLTextOnly = sGLSecond And Replace(sAssSecond, " ", "") = "" And Replace(sText, " ", "") <> "" And Replace(sClearNoteSecond, " ", "") = "" Then
                Cells(iRowSAPSecond, iColSAPClear) = "Offset"
                Range(Cells(iRowSAPSecond, 1), Cells(iRowSAPSecond, iColSAPPostKey)).Interior.ColorIndex = 15
            End If
        Next iRowSAPSecond
        
    End If
CONTINUENEXTARRAY:
Next iArray


End Sub


'Step-5:Match transactions Text Only (Empty in Assignment Field, Non-Empty in Text Field)
'with PART Group
Sub Match_Empty_Assignment_NonEmpty_Text_to_PART_Group()

Dim sGLTextOnly As String
Dim dAMTTextOnly As Double
Dim iArrayTextOnly As Integer

Dim sGLPartGroup As String
Dim sLeadingNumberPartGroup As String
Dim dAMTPartGroup As Double
Dim iArrayPartGroup As Integer

Dim wsSAP As Worksheet
Dim iMaxRowSAP As Integer
Dim iRowSAP As Integer
Dim sGL As String
Dim sAss As String
Dim sText As String
Dim dAMT As Double
Dim sClearNote As String
Dim sLeadingNumber As String

Dim iFound As Integer
Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow
If iMaxRowSAP < 2 Then Exit Sub

If UBound(avTextOnly, 2) < 1 Then Exit Sub
If UBound(avPartGroup, 2) < 1 Then Exit Sub

For iArrayTextOnly = 1 To UBound(avTextOnly, 2)
    sGLTextOnly = avTextOnly(0, iArrayTextOnly)
    dAMTTextOnly = avTextOnly(1, iArrayTextOnly)
    iFound = 0
    'Debug.Print sGLTextOnly & "--" & dAMTTextOnly

    For iArrayPartGroup = 1 To UBound(avPartGroup, 2)
        sGLPartGroup = avPartGroup(0, iArrayPartGroup)
        sLeadingNumberPartGroup = avPartGroup(1, iArrayPartGroup)
        dAMTPartGroup = avPartGroup(2, iArrayPartGroup)
        'Debug.Print sGLPartGroup & "--"; sLeadingNumberPartGroup & "--"; dAMTPartGroup
        
        If sGLPartGroup = sGLTextOnly And Abs(dAMTPartGroup + dAMTTextOnly) < 0.001 Then
            iFound = 1
            'Debug.Print "Found"
            Exit For
        End If
    Next iArrayPartGroup
        
    'If TextOnly matches with PART Group
    If iFound = 1 Then
        'Debug.Print "Found"
        'work on Text Only, to Loop SAP Sheet, to find all Text Only relarted transactions, then mark "Offset"
        For iRowSAP = 2 To iMaxRowSAP
            sGL = Cells(iRowSAP, iColSAPGL)
            sAss = Cells(iRowSAP, iColSAPAss)
            sText = Cells(iRowSAP, iColSAPText)
            sClearNote = Cells(iRowSAP, iColSAPClear)
                
            If sGL = sGLTextOnly And Replace(sAss, " ", "") = "" And Replace(sText, " ", "") <> "" And Replace(sClearNote, " ", "") = "" Then
                Cells(iRowSAP, iColSAPClear) = "Offset"
                Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = 15
            End If
        Next iRowSAP
            
        'Then work on PART Group, find all PART Group related transactions, then mark "Offset"
        For iRowSAP = 2 To iMaxRowSAP
            sGL = Cells(iRowSAP, iColSAPGL)
            sAss = Cells(iRowSAP, iColSAPAss)
            sText = Cells(iRowSAP, iColSAPText)
            sClearNote = Cells(iRowSAP, iColSAPClear)
            sLeadingNumber = Leading_Number(sText)
                
            'If sGL = sGLPartGroup And Replace(sAss, " ", "") <> "" And sLeadingNumber = sLeadingNumberPartGroup And Replace(sClearNote, " ", "") = "" Then
            If Replace(sAss, " ", "") <> "" And sLeadingNumber = sLeadingNumberPartGroup And Replace(sClearNote, " ", "") = "" Then
                
                Cells(iRowSAP, iColSAPClear) = "Offset"
                Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = 15
            End If

        Next iRowSAP
            
    End If
    
Next iArrayTextOnly

End Sub



Sub show_Array_Value()

Dim iArray As Integer

For iArray = LBound(avTextOnly, 2) To UBound(avTextOnly, 2)
    'Debug.Print iArray & "--" & avTextOnly(0, iArray) & "--" & avTextOnly(1, iArray)
Next iArray

For iArray = LBound(avPartGroup, 2) To UBound(avPartGroup, 2)
    'Debug.Print iArray & "--" & avPartGroup(0, iArray) & "--" & avPartGroup(1, iArray) & "--" & avPartGroup(2, iArray)
Next iArray

End Sub

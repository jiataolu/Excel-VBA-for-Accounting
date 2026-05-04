Attribute VB_Name = "Module021_OffsetAfterKyriba"
Option Explicit

Sub Matching_After_Kyriba()
Call Matching_After_Kyriba_1     ' 1:1 mappping
Call Matching_After_Kyriba_2_Interest_Income        'Interest income matching

End Sub

'GL-10301 matching
'SAP posting: Assignment not empty
'Kyriba posting: assignment empty, Text field is "Interest Income"
'3 (many) to 1 mapping

Sub Matching_After_Kyriba_2_Interest_Income()
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10301")
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10320")
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10322")
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10326")
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10327")
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10318")
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10325")
Call Matching_After_Kyriba_2_each_GL_Interest_Income("10303")


End Sub


Sub Matching_After_Kyriba_2_each_GL_Interest_Income(GLClearing As String)

'Dim GLClearing As String
'GLClearing = 10301

Dim dTotalAmtKyribaInterestIncome As Double
Dim dAmtSAPPosting As Double
Dim iFoundInterestIncome As Integer


Dim wsSAP As Worksheet
Dim iRowSAP As Integer
Dim iMaxRowSAP As Integer
Dim iRow2 As Integer
'Dim sSAPGL As String
Dim sSAPGL As String
Dim dSAPAmt As Double
Dim sSAPText As String
Dim sSAPAss As String

Dim sSAP2GL As Double
Dim dSAP2Amt As Double
Dim sSAP2Text As String
Dim sSAP2Ass As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow

dTotalAmtKyribaInterestIncome = 0
dAmtSAPPosting = 0

For iRowSAP = 2 To iMaxRowSAP
'For iRowSAP = 2 To 2

    
    If Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = xlNone And InStr(UCase(Cells(iRowSAP, iColSAPClear)), "OFFSET") = 0 Then
                
        sSAPGL = Cells(iRowSAP, iColSAPGL)
        sSAPAss = Cells(iRowSAP, iColSAPAss)
        sSAPText = Cells(iRowSAP, iColSAPText)
        dSAPAmt = Cells(iRowSAP, iColSAPAMT)
        
        'To find if it is Kyriba posting
        'Assignment field is empty, text field is not empty
        'If InStr(UCase(Replace(sSAPText, " ", "")), "WIRETYPE") > 0 Then GoTo CONTINUESAPLINE
        If Replace(sSAPAss, " ", "") <> "" Or Replace(sSAPText, " ", "") = "" Then GoTo CONTINUESAPLINE
        
        'If it is not interest income related. then skip
        If Replace(sSAPGL, " ", "") <> GLClearing Then GoTo CONTINUESAPLINE
        
        'If Text field does not contain "Interest Income", then skip
        If InStr(UCase(sSAPText), "INTEREST") = 0 Or InStr(UCase(sSAPText), "INCOME") = 0 Then GoTo CONTINUESAPLINE
        
        dTotalAmtKyribaInterestIncome = dTotalAmtKyribaInterestIncome + dSAPAmt
        
        'Debug.Print "Row-" & iRowSAP
        'Debug.Print "GL-" & sSAPGL
        'Debug.Print "Assignment-" & sSAPAss
        'Debug.Print "Text-" & sSAPText
        'Debug.Print "Amount-" & dSAPAmt
    End If
    
CONTINUESAPLINE:
Next iRowSAP

'Debug.Print dTotalAmtKyribaInterestIncome

If dTotalAmtKyribaInterestIncome = 0 Then GoTo FINISHINTERESTINCOME

iFoundInterestIncome = 0
'If multiple line of Kyriba's posting for interest income is found, then check the amount of SAP posting/
For iRow2 = 2 To iMaxRowSAP
    If Range(Cells(iRow2, 1), Cells(iRow2, iColSAPPostKey)).Interior.ColorIndex = xlNone And InStr(UCase(Cells(iRow2, iColSAPClear)), "OFFSET") = 0 Then
        
        'If wsSAP.Cells(iRow2, iColSAPGL) <> "10301" Then GoTo CONTINUESAPLINE
            
        sSAP2GL = Cells(iRow2, iColSAPGL)
        sSAP2Ass = Cells(iRow2, iColSAPAss)
        dSAP2Amt = Cells(iRow2, iColSAPAMT)
        sSAP2Text = Cells(iRow2, iColSAPText)
        'Debug.Print "GL2-" & sSAP2GL
        'Debug.Print "Assignment2-" & sSAP2Ass
        'Debug.Print "Text2-" & sSAP2Text
        'Debug.Print "Amount2-" & dSAP2Amt

            
        'if GL is not same, then skip
        If Replace(sSAP2GL, " ", "") <> GLClearing Then GoTo CONTINUEINNERLOOP
            
        'if Assignment field is empty, then skip it.
        If Replace(sSAP2Ass, " ", "") = "" Then GoTo CONTINUEINNERLOOP
            
        'to find SAP posting with "Wire type)
        If InStr(UCase(Replace(sSAP2Text, " ", "")), "WIRETYPE") = 0 Then GoTo CONTINUEINNERLOOP
            
        If Abs(dSAP2Amt + dTotalAmtKyribaInterestIncome) < 0.01 Then
            'Debug.Print "Found-" & iRow2
            'Debug.Print "SAP Amunt-" & dSAP2Amt
            iFoundInterestIncome = 1
            
            Debug.Print "Interest Income Highlight row-" & iRow2
            Cells(iRow2, iColSAPClear) = "Offset"
            Range(Cells(iRow2, 1), Cells(iRow2, iColSAPPostKey)).Interior.ColorIndex = 15
                    
            Exit For
        End If
    End If
CONTINUEINNERLOOP:

Next iRow2

'If find SAP  posting, then highlight all Kyriba posting lines
If iFoundInterestIncome = 0 Then GoTo FINISHINTERESTINCOME

For iRowSAP = 2 To iMaxRowSAP
'For iRowSAP = 2 To 2
    'sSAPGL = wsSAP.Cells(iRowSAP, iColSAPGL)
    'for all bank account
    'If wsSAP.Cells(iRowSAP, iColSAPGL) <> "10301" Then GoTo CONTINUESAPLINE
    
    If Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = xlNone And InStr(UCase(Cells(iRowSAP, iColSAPClear)), "OFFSET") = 0 Then
                
        sSAPGL = Cells(iRowSAP, iColSAPGL)
        sSAPAss = Cells(iRowSAP, iColSAPAss)
        sSAPText = Cells(iRowSAP, iColSAPText)
        dSAPAmt = Cells(iRowSAP, iColSAPAMT)
        
        'To find if it is Kyriba posting
        'Assignment field is empty, text field is not empty
        'If InStr(UCase(Replace(sSAPText, " ", "")), "WIRETYPE") > 0 Then GoTo CONTINUESAPLINE
        If Replace(sSAPAss, " ", "") <> "" Or Replace(sSAPText, " ", "") = "" Then GoTo CONTINUEHIGHLIGHTSAPLINE
        
        'If it is not interest income related. then skip
        If Replace(sSAPGL, " ", "") <> GLClearing Then GoTo CONTINUEHIGHLIGHTSAPLINE
        
        'If Text field does not contain "Interest Income", then skip
        If InStr(UCase(sSAPText), "INTEREST") = 0 Or InStr(UCase(sSAPText), "INCOME") = 0 Then GoTo CONTINUEHIGHLIGHTSAPLINE
        
        Debug.Print "Interest Income Highlight Row-" & iRowSAP
        Cells(iRowSAP, iColSAPClear) = "Offset"
        Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = 15
    End If
    
CONTINUEHIGHLIGHTSAPLINE:
Next iRowSAP

FINISHINTERESTINCOME:
End Sub













'GL-10301 matching
'SAP posting: Assignment not empty
'Kyriba posting: assignment empty, text field is not empty
'1 to 1 mapping
Sub Matching_After_Kyriba_1()

Dim wsSAP As Worksheet
Dim iRowSAP As Integer
Dim iMaxRowSAP As Integer
Dim iRow2 As Integer
'Dim sSAPGL As String
Dim sSAPGL As String
Dim dSAPAmt As Double
Dim sSAPText As String
Dim sSAPAss As String

Dim sSAP2GL As Double
Dim dSAP2Amt As Double
Dim sSAP2Text As String
Dim sSAP2Ass As String

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wsSAP = Worksheets("1-SAP")
wsSAP.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSAP = lRealLastRow

For iRowSAP = 2 To iMaxRowSAP
'For iRowSAP = 2 To 2
    'sSAPGL = wsSAP.Cells(iRowSAP, iColSAPGL)
    'for all bank account
    'If wsSAP.Cells(iRowSAP, iColSAPGL) <> "10301" Then GoTo CONTINUESAPLINE
    
    If Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = xlNone And InStr(UCase(Cells(iRowSAP, iColSAPClear)), "OFFSET") = 0 Then
                
        sSAPGL = Cells(iRowSAP, iColSAPGL)
        sSAPAss = Cells(iRowSAP, iColSAPAss)
        sSAPText = Cells(iRowSAP, iColSAPText)
        dSAPAmt = Cells(iRowSAP, iColSAPAMT)
        'Debug.Print "GL-" & sSAPGL
        'Debug.Print "Assignment-" & sSAPAss
        'Debug.Print "Text-" & sSAPText
        'Debug.Print "Amount-" & dSAPAmt
        
        'To find if it is Kyriba posting
        'Assignment field is empty, text field is not empty
        'If InStr(UCase(Replace(sSAPText, " ", "")), "WIRETYPE") > 0 Then GoTo CONTINUESAPLINE
        If Replace(sSAPAss, " ", "") <> "" Or Replace(sSAPText, " ", "") = "" Then GoTo CONTINUESAPLINE
        
        
        For iRow2 = 2 To iMaxRowSAP
        'For iRow2 = 19 To 19
            If Range(Cells(iRow2, 1), Cells(iRow2, iColSAPPostKey)).Interior.ColorIndex = xlNone And InStr(UCase(Cells(iRow2, iColSAPClear)), "OFFSET") = 0 Then
        
                'If wsSAP.Cells(iRow2, iColSAPGL) <> "10301" Then GoTo CONTINUESAPLINE
            
                sSAP2GL = Cells(iRow2, iColSAPGL)
                sSAP2Ass = Cells(iRow2, iColSAPAss)
                dSAP2Amt = Cells(iRow2, iColSAPAMT)
                sSAP2Text = Cells(iRow2, iColSAPText)
                'Debug.Print "GL2-" & sSAP2GL
                'Debug.Print "Assignment2-" & sSAP2Ass
                'Debug.Print "Text2-" & sSAP2Text
                'Debug.Print "Amount2-" & dSAP2Amt

            
                'if GL is not same, then skip
                If Replace(sSAP2GL, " ", "") <> Replace(sSAPGL, " ", "") Then GoTo CONTINUEINNERLOOP
            
                'if Assignment field is empty, then skip it.
                If Replace(sSAP2Ass, " ", "") = "" Then GoTo CONTINUEINNERLOOP
            
                'inner iteration is to find SAP posting with "Wire type), if GL is 10301
                'If Gl is not 10301, to find text field should not be empty
                If Replace(sSAP2GL, " ", "") = "10301" And InStr(UCase(Replace(sSAP2Text, " ", "")), "WIRETYPE") = 0 Then GoTo CONTINUEINNERLOOP
                If Replace(sSAP2GL, " ", "") <> "10301" And Replace(sSAP2Text, " ", "") = "" Then GoTo CONTINUEINNERLOOP
            
            
                If Abs(dSAP2Amt + dSAPAmt) < 0.01 Then
                    'Debug.Print "Found-" & iRowSAP & "-" & iRow2
                    
                    Cells(iRowSAP, iColSAPClear) = "Offset"
                    Range(Cells(iRowSAP, 1), Cells(iRowSAP, iColSAPPostKey)).Interior.ColorIndex = 15
                    
                    Cells(iRow2, iColSAPClear) = "Offset"
                    Range(Cells(iRow2, 1), Cells(iRow2, iColSAPPostKey)).Interior.ColorIndex = 15
                    
                    Exit For
                End If
            End If
CONTINUEINNERLOOP:
        Next iRow2
        
    End If
    


    


CONTINUESAPLINE:

Next iRowSAP


End Sub

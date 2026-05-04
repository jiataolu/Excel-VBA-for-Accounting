Attribute VB_Name = "Module11_ToProgramCase"
Option Explicit

'Sanjay says that we do not check 10901 by email anymore
Sub To_Program_GL_10901_Email_To_To_Confirm()


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
    sPostGL = Cells(iRowItems, iColItemsPostGL)
    If Replace(sPostGL, " ", "") = "10901" Then
                
        sPostBU = Cells(iRowItems, iColItemsPostBU)
        'sPostGL = Cells(iRowItems, iColItemsPostGL)
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
                
        
    End If
    
Next iRowItems

End Sub

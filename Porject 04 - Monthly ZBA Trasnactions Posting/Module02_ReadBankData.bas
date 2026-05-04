Attribute VB_Name = "Module02_ReadBankData"
Option Explicit

Sub Read_Bank_Data()
Call Read_Bank_Data_Step_1_Clear_JEData_Sheet
Call Read_Bank_Data_Step_2_Read_Column
Call Read_Bank_Data_Step_4_Extra_Columns
Call Read_Bank_Data_Step_5_Format_Sheet_JE_Data
Call Read_Bank_Data_Step_6_Find_Duplicate_ZBA
Call DeleteUnusedFormats
End Sub


Sub Read_Bank_Data_Step_1_Clear_JEData_Sheet()

Dim ws02JEData As Worksheet

Set ws02JEData = Worksheets(Sheet02Name_JEData)
ws02JEData.Select
Cells.Select
Selection.Delete

Cells.Interior.Pattern = xlNone
With Selection.Interior
    .Pattern = xlNone
    '.TintAndShade = 0
    '.PatternTintAndShade = 0
End With
Cells(1, 1).Select


End Sub

Sub Read_Bank_Data_Step_2_Read_Column()
Call Read_Column_from_Integrity_Statement_to_JE_Data("Value Date", 1)
Call Read_Column_from_Integrity_Statement_to_JE_Data("Account Code", 2)
Call Read_Column_from_Integrity_Statement_to_JE_Data("BT Code", 3)
Call Read_Column_from_Integrity_Statement_to_JE_Data("Ccy", 4)
Call Read_Column_from_Integrity_Statement_to_JE_Data("Reference", 5)
Call Read_Column_from_Integrity_Statement_to_JE_Data("Worksheet Category", 6)
Call Read_Column_from_Integrity_Statement_to_JE_Data("Amount", 7)

End Sub


'to add ZBA account, ZBA bank code, ZBA GL, Account GL
Sub Read_Bank_Data_Step_4_Extra_Columns()

Application.DisplayAlerts = False

Dim ws02JEData As Worksheet
Dim iMaxRowJEData As Long
Dim iRowJEData As Long
Dim sAccountBU As String
Dim sAccountGL As String

Dim sRef As String
Dim sRef1 As String
Dim sRef2 As String
Dim sZBABankAccount As String
Dim sZBABankCode As String
Dim sZBABU As String
Dim sZBAGL As String
Dim iPos1 As Integer
Dim iPos2 As Integer



Dim wkbMap As Workbook
Dim wsMap As Worksheet
Dim rngMapBankCode As Range
'Dim rngMapBankAccount As Range
Dim rngFound As Range
Dim iMaxRowMap As Long
Dim iRowMap As Long
Dim sMapBankAccount As String
Dim iFoundinMap As Integer

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set wkbMap = Workbooks.Open(Map_File_Full_Name, UpdateLinks:=False)
wkbMap.Activate
Set wsMap = Worksheets("Mapping Consolidated")
wsMap.Select
Set rngMapBankCode = Columns(SheetMapColBankCode)
'Set rngMapBankAccount = Columns("B")
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowMap = lRealLastRow


ThisWorkbook.Activate
Set ws02JEData = Worksheets(Sheet02Name_JEData)
ws02JEData.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowJEData = lRealLastRow


ws02JEData.Cells(1, Sheet02ColAccountBU) = "Account BU"
ws02JEData.Cells(1, Sheet02ColAccountGL) = "Account GL"
ws02JEData.Cells(1, Sheet02ColZBAAccount) = "ZBA Account"
ws02JEData.Cells(1, Sheet02ColZBABankCode) = "ZBA Bank Code"
ws02JEData.Cells(1, Sheet02ColZBABU) = "ZBA BU"
ws02JEData.Cells(1, Sheet02ColZBAGL) = "ZBA GL"


For iRowJEData = 2 To iMaxRowJEData
'For iRowJEData = 2 To 2
    
    'account BU
    sAccountBU = ws02JEData.Cells(iRowJEData, Sheet02ColAccountBankCode)

    Set rngFound = rngMapBankCode.Find(sAccountBU, LookIn:=xlValues, lookat:=xlWhole)
        
    If Not rngFound Is Nothing Then
        ws02JEData.Cells(iRowJEData, Sheet02ColAccountBU) = rngFound.Cells(1, SheetMapColBU - SheetMapColBankCode + 1)
        
        sAccountGL = rngFound.Cells(1, SheetMapColGL - SheetMapColBankCode + 1)
        sAccountGL = Replace(sAccountGL, " ", "")
        
        If Not Validate_GL(sAccountGL) Then sAccountGL = Replace(rngFound.Cells(1, SheetMapColVendor - SheetMapColBankCode + 1), " ", "")
        
        ws02JEData.Cells(iRowJEData, Sheet02ColAccountGL) = sAccountGL
        
    End If
    
    
    sRef = ws02JEData.Cells(iRowJEData, Sheet02ColRef)
    sRef = Right(sRef, Len(sRef) - InStr(sRef, " "))
    sZBABankAccount = sRef
    sRef1 = Left(sRef, InStr(sRef, " ") - 1)
    sRef2 = Right(sRef, Len(sRef) - InStr(sRef, " "))
    ws02JEData.Cells(iRowJEData, Sheet02ColZBAAccount) = sZBABankAccount

    For iRowMap = 2 To iMaxRowMap
    'For iRowMap = 146 To 146
        iFoundinMap = 0
        sMapBankAccount = wsMap.Cells(iRowMap, SheetMapColBankAccount)
        
        iPos1 = InStr(sMapBankAccount, sRef1)
        iPos2 = InStr(sMapBankAccount, sRef2)

        If iPos1 * iPos2 > 0 Then
            sZBABankCode = wsMap.Cells(iRowMap, SheetMapColBankCode)
            sZBABU = wsMap.Cells(iRowMap, SheetMapColBU)
            
            sZBAGL = Replace(wsMap.Cells(iRowMap, SheetMapColGL), " ", "")
            If Not Validate_GL(sZBAGL) Then sZBAGL = Replace(wsMap.Cells(iRowMap, SheetMapColVendor), " ", "")
            
            iFoundinMap = 1
            Exit For
        End If
    Next iRowMap
    
    If iFoundinMap = 1 Then
    
        ws02JEData.Cells(iRowJEData, Sheet02ColZBABankCode) = sZBABankCode
        ws02JEData.Cells(iRowJEData, Sheet02ColZBABU) = sZBABU
        ws02JEData.Cells(iRowJEData, Sheet02ColZBAGL) = sZBAGL
    End If
    
    
Next iRowJEData

wkbMap.Close savechanges:=False
Application.DisplayAlerts = True

Set rngMapBankCode = Nothing
Set rngFound = Nothing

ws02JEData.Cells(1, Sheet02ColZBADuplicate) = "Duplicate ZBA"


End Sub


Sub Read_Bank_Data_Step_5_Format_Sheet_JE_Data()
Dim ws02JEData As Worksheet
Set ws02JEData = Worksheets(Sheet02Name_JEData)

Cells.Select
Columns.AutoFit
Cells(1, 1).Select
Columns("A:F").HorizontalAlignment = xlCenter
Columns("H:M").HorizontalAlignment = xlCenter
Columns("G:G").Style = "Comma"

Range("B:B,H:H,I:I").Interior.ThemeColor = xlThemeColorAccent4
Range("B:B,H:H,I:I").Interior.TintAndShade = 0.799981688894314

Range("J:J,K:K,L:L,M:M").Interior.ThemeColor = xlThemeColorAccent5
Range("J:J,K:K,L:L,M:M").Interior.TintAndShade = 0.799981688894314


Cells(1, 1).Select

End Sub

Sub Read_Bank_Data_Step_6_Find_Duplicate_ZBA()

Dim ws02JEData As Worksheet
Dim iMaxRowJEData As Long
Dim iRowJEData As Long
Dim sBankCode As String
Dim sZBACode As String
Dim sZBADate As String
Dim dZBAAMT As Double

Dim iRow2 As Long
Dim sBankCode2 As String
Dim sZBACode2 As String
Dim sZBADate2 As String
Dim dZBAAMT2 As Double

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Set ws02JEData = Worksheets(Sheet02Name_JEData)
ws02JEData.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowJEData = lRealLastRow
If iMaxRowJEData < 3 Then Exit Sub

For iRowJEData = 2 To iMaxRowJEData - 1
    If Not IsEmpty(ws02JEData.Cells(iRowJEData, Sheet02ColZBADuplicate)) Then GoTo CONTINUENEXTROWDATA
    
    sBankCode = ws02JEData.Cells(iRowJEData, Sheet02ColAccountBankCode)
    sZBACode = ws02JEData.Cells(iRowJEData, Sheet02ColZBABankCode)
    sZBADate = ws02JEData.Cells(iRowJEData, Sheet02ColDate)
    dZBAAMT = ws02JEData.Cells(iRowJEData, Sheet02ColAmount)
    
    For iRow2 = iRowJEData + 1 To iMaxRowJEData
        If Not IsEmpty(ws02JEData.Cells(iRow2, Sheet02ColZBADuplicate)) Then GoTo CONTINUENEXTROWLEVEL2
        
        sBankCode2 = ws02JEData.Cells(iRow2, Sheet02ColAccountBankCode)
        sZBACode2 = ws02JEData.Cells(iRow2, Sheet02ColZBABankCode)
        sZBADate2 = ws02JEData.Cells(iRow2, Sheet02ColDate)
        dZBAAMT2 = ws02JEData.Cells(iRow2, Sheet02ColAmount)
        
        If (sZBADate = sZBADate2) And (sBankCode = sZBACode2) And (sZBACode = sBankCode2) And (dZBAAMT = dZBAAMT2 * -1) Then
            ws02JEData.Cells(iRowJEData, Sheet02ColZBADuplicate) = "O-" & CStr(iRow2)
            ws02JEData.Cells(iRow2, Sheet02ColZBADuplicate) = "D-" & CStr(iRowJEData)
            Exit For
        End If
CONTINUENEXTROWLEVEL2:
    Next iRow2
    
CONTINUENEXTROWDATA:
Next iRowJEData



End Sub


Sub Read_Column_from_Integrity_Statement_to_JE_Data(HeaderName As String, ColumnNumber As Integer)
'Dim HeaderName As String
'HeaderName = "Value Date"

'Dim ColumnNumber As Integer
'ColumnNumber = 1

Dim ws01Integrity As Worksheet
Dim iMaxRowIntegrity As Long
Dim iMaxColIntegrity As Integer
Dim iColIntegrity As Integer
Dim sHeader As String

Dim ws02JEData As Worksheet

Dim rngCopy As Range
Dim rngPaste As Range

Dim iColumnFound As Integer

Set ws02JEData = Worksheets(Sheet02Name_JEData)

Set ws01Integrity = Worksheets(Sheet01Name_Integrity)
ws01Integrity.Select
Dim lRealLastRow As Long
Dim lRealLastCol As Long
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxColIntegrity = lRealLastCol '

iColumnFound = 0
For iColIntegrity = 1 To iMaxColIntegrity
    sHeader = ws01Integrity.Cells(1, iColIntegrity)
    If LCase(sHeader) = LCase(HeaderName) Then
        'Debug.Print iColIntegrity
        ws01Integrity.Select
        Set rngCopy = Columns(iColIntegrity)
        
        ws02JEData.Select
        Set rngPaste = ws02JEData.Cells(1, ColumnNumber)
    
        rngCopy.Copy Destination:=rngPaste
        iColumnFound = 1
        Exit For
        
    End If
    
Next iColIntegrity

If iColumnFound = 0 Then
    ws02JEData.Cells(1, ColumnNumber) = HeaderName
    MsgBox ("Column " & HeaderName & " does not exisit, please check Integrity statement!")
End If


End Sub

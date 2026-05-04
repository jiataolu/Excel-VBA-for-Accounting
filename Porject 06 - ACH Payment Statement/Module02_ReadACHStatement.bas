Attribute VB_Name = "Module02_ReadACHStatement"
Option Explicit
Option Base 1

Sub Read_ACH()

Dim wsACHList As Worksheet
Dim iMaxRowACHList As Long
Dim iRowACHList As Long

Dim objFSO As Object
Dim objFolder As Object
Dim objFiles As Object
Dim objFile As Object

Dim sFullPathACH As String

Dim sFullPathLog As String
Dim objLog As Object
Dim iLogCount As Long


Dim lRealLastRow As Long
Dim lRealLastCol As Long

ThisWorkbook.Activate
Set wsACHList = Worksheets(SheetNameACHList)
wsACHList.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowACHList = lRealLastRow
If iMaxRowACHList > 2 Then
    Rows("2:" & iMaxRowACHList).Delete
    DeleteUnusedFormats
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

'check Log file
sFullPathLog = GetWorkPath & "\" & FileNameLog
If objFSO.FileExists(sFullPathLog) Then
    objFSO.DeleteFile sFullPathLog
End If
' Create new file
Set objLog = objFSO.CreateTextFile(sFullPathLog, True)
objLog.Close

Set objFolder = objFSO.Getfolder(GetWorkPath & FolderACHStatement)
Set objFiles = objFolder.Files

iLogCount = 0
For Each objFile In objFiles
    Set objLog = objFSO.openTextFile(sFullPathLog, 8, True)
    
    sFullPathACH = GetWorkPath & FolderACHStatement & "\" & objFile.Name
    Call Read_Write_ACH_Statement_by_File(sFullPathACH)
    
    iLogCount = iLogCount + 1
    objLog.writeline (CStr(iLogCount) & ".    " & sFullPathACH)
    objLog.Close
Next
End Sub


Sub Read_Write_ACH_Statement_by_File(PathACHState)

Application.ScreenUpdating = False
'Dim PathACHState As String
'PathACHState = "C:\Users\jlu\Lu work\01 - VBA\01 - Banking\06 - MMS-210020 & 210030\063 - MMS-210030 Bank Recon\02 - VBA Macro\ACH Statement - Excel\ACH_1115 FEB 01- FEB 08.xlsx"

Dim arrACHState As Variant
Dim arrACHList() As Variant

Dim wkbACHState As Workbook
Dim wsACHState As Worksheet
Dim iMaxRowACHState As Long
Dim iMaxColACHState As Long
Dim iRowACHState As Long
Dim iColACHState As Long
Dim sCompanyID As String
Dim iRecord As Long

Dim wsACHList As Worksheet
Dim iMaxRowSheetACHList As Long

Dim iMaxRowACHList As Long
Dim iMaxColACHList As Integer
Dim iRowACHList As Long

Dim lRealLastRow As Long
Dim lRealLastCol As Long

Dim sCKinfo As String

Set wkbACHState = Workbooks.Open(PathACHState)
Set wsACHState = Worksheets(1)
wsACHState.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowACHState = lRealLastRow
iMaxColACHState = lRealLastCol

arrACHState = Range(Cells(1, 1), Cells(iMaxRowACHState, iMaxColACHState)).Value

iMaxRowACHList = 0

Debug.Print iMaxRowACHState

For iRowACHState = 2 To iMaxRowACHState
    sCompanyID = Replace(CStr(arrACHState(iRowACHState, ColCompanyID)), " ", "")
    If sCompanyID = CompanyID Then iMaxRowACHList = iMaxRowACHList + 1
Next iRowACHState
'Debug.Print iMaxRowACHList
iMaxColACHList = iMaxColACHState - 1
'Debug.Print iMaxColACHList

ReDim arrACHList(1 To iMaxRowACHList, 1 To iMaxColACHList)
'Debug.Print LBound(arrACHList, 1)

iRowACHList = 0
sCKinfo = ""
iRecord = 0
Dim iCount As Long
iCount = 1
For iRowACHState = 2 To iMaxRowACHState

    'Debug.Print iRowACHState
    'iCount = iCount + 1
    'Debug.Print iCount
    sCompanyID = Replace(CStr(arrACHState(iRowACHState, ColCompanyID)), " ", "")
    'Debug.Print "CompanyID: " & sCompanyID
    
    'If the first row of transaction, then it is a new start
    'according to Recording true or false
    'to decide if to write CK info to previous transactions
    If sCompanyID <> "" And iRecord = 1 Then
        'start a new transaction, if last one has recorded cheque info, then write cheque info
        If sCKinfo <> "" Then
            'Debug.Print Left(sCKinfo, 20)
            'Debug.Print sCKinfo
            arrACHList(iRowACHList, iMaxColACHList) = RegExp_CK(sCKinfo)
        End If
    End If
    
    'if a new transaction, and company ID is 2023904428, then
    'start to record, and write into ACHList array
    If sCompanyID = CompanyID Then
        iRowACHList = iRowACHList + 1
        'Debug.Print "ACH List Row: ," & iRowACHList
        iRecord = 1
        
        For iColACHState = 1 To iMaxColACHState - 1
            arrACHList(iRowACHList, iColACHState) = arrACHState(iRowACHState, iColACHState)
        Next iColACHState
        
        sCKinfo = arrACHState(iRowACHState, iMaxColACHState)
        
        GoTo CONTINUEACHSTATELINE
    End If
    
    
    'if a new transacation, and company ID is not 2023904428
    'stop recording
    If sCompanyID <> CompanyID And sCompanyID <> "" Then
        iRecord = 0
        GoTo CONTINUEACHSTATELINE
    End If
    
    'if a continuos line of transaction, and if recording is true
    'then append CK info
    If sCompanyID = "" And iRecord = 1 Then
        sCKinfo = sCKinfo & arrACHState(iRowACHState, iMaxColACHState)
        GoTo CONTINUEACHSTATELINE
    End If

CONTINUEACHSTATELINE:
Next iRowACHState
If iRecord = 1 Then
    'Debug.Print sCKinfo
    arrACHList(iRowACHList, iMaxColACHList) = RegExp_CK(sCKinfo)
End If


ThisWorkbook.Activate
Set wsACHList = Worksheets(SheetNameACHList)
wsACHList.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowSheetACHList = lRealLastRow
Cells(iMaxRowSheetACHList + 1, 1).Resize(iMaxRowACHList, iMaxColACHList).Value = arrACHList
DeleteUnusedFormats

wkbACHState.Close savechanges:=False

Application.ScreenUpdating = True
End Sub

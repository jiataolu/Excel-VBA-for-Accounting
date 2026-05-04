Attribute VB_Name = "Module16_ReportFiles"
Option Explicit

Sub Generate_Daily_JE_File()
Application.DisplayAlerts = False

Dim wsSheet1 As Worksheet
Dim wsSheet2 As Worksheet
Dim wsSheet3 As Worksheet

Dim wsNewSheet1 As Worksheet
Dim wsNewSheet2 As Worksheet
Dim wsNewSheet3 As Worksheet


Dim wkbNewDailyJE As Workbook
Dim sDailyJEFileFullName As String

ThisWorkbook.Activate
Set wsSheet1 = Worksheets("1-SAP")
Set wsSheet2 = Worksheets("2-Items to post")
Set wsSheet3 = Worksheets("3 - C-SAP Standard Template")


sDailyJEFileFullName = GetWorkPath() & "\" & SubFolderOutput & "\" & FileNameDailyJE

If Dir(sDailyJEFileFullName) <> "" Then
    ' Delete the existing file
    Kill sDailyJEFileFullName
End If

Set wkbNewDailyJE = Workbooks.Add

wsSheet1.Copy After:=wkbNewDailyJE.Sheets(wkbNewDailyJE.Sheets.Count)
Set wsNewSheet1 = wkbNewDailyJE.Worksheets(wkbNewDailyJE.Sheets.Count)
wsNewSheet1.Name = "1-SAP"

wsSheet2.Copy After:=wkbNewDailyJE.Sheets(wkbNewDailyJE.Sheets.Count)
Set wsNewSheet2 = wkbNewDailyJE.Worksheets(wkbNewDailyJE.Sheets.Count)
wsNewSheet2.Name = "2-Items to post"

wsSheet3.Copy After:=wkbNewDailyJE.Sheets(wkbNewDailyJE.Sheets.Count)
Set wsNewSheet3 = wkbNewDailyJE.Worksheets(wkbNewDailyJE.Sheets.Count)
wsNewSheet3.Name = "3 - C-SAP Standard Template"

wkbNewDailyJE.Worksheets("Sheet1").Delete
wkbNewDailyJE.Worksheets("3 - C-SAP Standard Template").Select

wkbNewDailyJE.SaveCopyAs Filename:=sDailyJEFileFullName
wkbNewDailyJE.Close SaveChanges:=False

Application.DisplayAlerts = True
End Sub


Sub Generate_Adjusting_JE_File()
Application.DisplayAlerts = False

Dim wkbNewAdjustJE As Workbook
Dim wsNewSheetPendingList As Worksheet
Dim wsNewSheetJE As Worksheet


Dim wkbPending As Workbook
Dim sPendingFileFullName As String
Dim wsPending As Worksheet
Dim sAdjustJEFileFullName As String

Dim wsJE As Worksheet

sPendingFileFullName = GetWorkPath() & "\" & FileNamePending
Set wkbPending = Workbooks.Open(sPendingFileFullName)
Set wsPending = wkbPending.Worksheets("Pending")

sAdjustJEFileFullName = GetWorkPath() & "\" & SubFolderOutput & "\" & FileNameAdjustingJE
Set wkbNewAdjustJE = Workbooks.Add

wsPending.Copy After:=wkbNewAdjustJE.Sheets(wkbNewAdjustJE.Sheets.Count)
Set wsNewSheetPendingList = wkbNewAdjustJE.Worksheets(wkbNewAdjustJE.Sheets.Count)
wsNewSheetPendingList.Name = "Pending"

wkbPending.Close SaveChanges:=False


ThisWorkbook.Activate
Set wsJE = Worksheets("3 - C-SAP Standard Template")
wsJE.Copy After:=wkbNewAdjustJE.Sheets(wkbNewAdjustJE.Sheets.Count)
Set wsNewSheetJE = wkbNewAdjustJE.Worksheets(wkbNewAdjustJE.Sheets.Count)
wsNewSheetJE.Name = "3 - C-SAP Standard Template"

wkbNewAdjustJE.Worksheets("Sheet1").Delete


wkbNewAdjustJE.SaveCopyAs Filename:=sAdjustJEFileFullName
wkbNewAdjustJE.Close SaveChanges:=False

Application.DisplayAlerts = True

End Sub

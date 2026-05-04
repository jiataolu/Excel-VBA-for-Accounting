Attribute VB_Name = "Module090_Finalize"
Option Explicit


Sub Mapping_090_Finalize()
Dim wsFIS As Worksheet
Dim wsMap As Worksheet





'Remove working columns
Set wsFIS = Worksheets(SheetNameFIS)
wsFIS.Select
Columns(ColFISKeyNumber).Delete
Columns(ColFISRemark).Delete
Cells.Select
Selection.EntireColumn.AutoFit
Cells(1, 1).Select

Set wsMap = Worksheets(SheetNameMapping)
wsMap.Select

Columns(ColMapBankAcctKey).Delete
Columns(ColMapRemark).Delete


Columns(ColMapBankAcctFull).HorizontalAlignment = xlCenter
Columns(ColMapFISCode).HorizontalAlignment = xlCenter
Columns(ColMapKyribaCode).HorizontalAlignment = xlCenter
Columns(ColMapCry).HorizontalAlignment = xlCenter
Columns(ColMapERPSystem).HorizontalAlignment = xlCenter
Columns(ColMapFISBUCode).HorizontalAlignment = xlCenter
'Columns(ColMapSAPBUCode).HorizontalAlignment = xlCenter
Columns(ColMapFISSapGL).HorizontalAlignment = xlCenter
Columns(ColMapLocalBU).HorizontalAlignment = xlCenter
Columns(ColMapLocalGL).HorizontalAlignment = xlCenter
Columns(ColMapBUName).HorizontalAlignment = xlCenter
Columns(ColMapVendorCode).HorizontalAlignment = xlCenter
Columns(ColMapParentCode).HorizontalAlignment = xlCenter
Columns(ColMapProductCode).HorizontalAlignment = xlCenter
Columns(ColMapDataSource).HorizontalAlignment = xlCenter
Columns(ColMapCompanyName).HorizontalAlignment = xlLeft
Columns(ColMapOwnership).HorizontalAlignment = xlLeft



Cells.Select
Selection.EntireColumn.AutoFit
Columns(ColMapCompanyName).ColumnWidth = 20
Cells(1, 1).Select


'MsgBox "Deleted acounts - " & CStr(LineDeleted)
'MsgBox "New Added Accounts - " & CStr(LineNew)
End Sub


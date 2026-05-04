Attribute VB_Name = "Module030_format_FIS_PeopelSoft"
Option Explicit


Sub Mapping_030_Format_FIS_PS()

Dim wsFIS As Worksheet

Set wsFIS = Worksheets(SheetNameFIS)
wsFIS.Select

Columns(ColFISKeyNumber).HorizontalAlignment = xlCenter
Columns(ColFISProductCode).HorizontalAlignment = xlCenter
Columns(ColFISIsinFIS).HorizontalAlignment = xlCenter
Columns(ColFISIsinPS).HorizontalAlignment = xlCenter
Columns(ColFISRemark).HorizontalAlignment = xlLeft
Columns(ColFISBankAcct).HorizontalAlignment = xlCenter
Columns(ColFISFISCode).HorizontalAlignment = xlLeft
Columns(ColFISBUCode).HorizontalAlignment = xlCenter
Columns(ColFISSapGL).HorizontalAlignment = xlCenter
Columns(ColFISCurrency).HorizontalAlignment = xlCenter
Columns(ColFISCompanyName).HorizontalAlignment = xlLeft

End Sub


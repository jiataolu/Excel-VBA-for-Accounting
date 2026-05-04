Attribute VB_Name = "Module0_ribbon"
'Callback for customUI.onLoad
Sub Ribbon_Onload(ribbon As IRibbonUI)
End Sub

Sub ReadJEData(control As IRibbonControl)
Call Read_Bank_Data
Call Remove_Duplicate_Rows
End Sub

Sub GeneratePivotTable(control As IRibbonControl)
Call Generate_Pivot_Table
End Sub

Sub JEUpload(control As IRibbonControl)
Call Fill_JE_Template
Call Validation
Worksheets(Sheet05Name_JEUploadCAD).Select
End Sub


Attribute VB_Name = "Module0_ribbon"
'Callback for customUI.onLoad
Sub Ribbon_Onload(ribbon As IRibbonUI)
End Sub

Sub S4_Conversion(control As IRibbonControl)
Call Transfer_S4_Data
End Sub


Sub Read_PAP_Bank_Statement(control As IRibbonControl)
Call Final_Bank_Reconciliation
End Sub

Sub MSD_Report(control As IRibbonControl)
Call MSD_Clearing
End Sub

Sub SPS_Report(control As IRibbonControl)
Call SPS_Clearing
End Sub

Sub Wellca_Report(control As IRibbonControl)
Call Wellca_Clearing
End Sub

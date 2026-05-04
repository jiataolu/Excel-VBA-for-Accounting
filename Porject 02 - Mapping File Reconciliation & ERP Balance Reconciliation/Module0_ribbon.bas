Attribute VB_Name = "Module0_ribbon"
'Callback for customUI.onLoad
Sub Ribbon_Onload(ribbon As IRibbonUI)
End Sub

'Callback for Button001 onAction
Sub OnAction1_FISPeopleSoft(control As IRibbonControl)

Call Mapping_010_Clear_Initialize_FIS
Call Mapping_020_Combine_PeopleSoft_FIS
Call Mapping_030_Format_FIS_PS
Call Mapping_040_Consolidate_FIS_PS

End Sub


'Callback for Button002 onAction
Sub OnAction2_BankAccounts(control As IRibbonControl)
Call Mapping_050_Verfiy_Mapping_with_IFS_and_PeopleSoft
Call Mapping_060_Add_New_Lines
Call Mapping_070_Delete_Lines


End Sub

'Callback for Button003 onAction
Sub OnAction3_Dictionary(control As IRibbonControl)
Call Mapping_080_Read_Company_Code


End Sub

'Callback for Button004 onAction
Sub OnAction4_ReconInfo(control As IRibbonControl)
Call Mapping_085_Ownership
Call Mapping_090_Finalize
End Sub

'Callback for Button005 onAction
Sub onAction5_OutPutMapping(control As IRibbonControl)
Call Mapping_100_Output_Mapping_Report
End Sub








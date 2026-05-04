Attribute VB_Name = "Module0_ribbon"
'Callback for customUI.onLoad
Sub Ribbon_Onload(ribbon As IRibbonUI)
End Sub

Sub Act_GLBank(control As IRibbonControl)
Call Find_Unique_Bank_Date
Call Process_GL_Bank_Data

Call Generate_Pivot_Table_01_Clear_Pivot_Sheet
Call Generate_Pivot_Table_02_GL
Call Generate_Pivot_Table_03_Bank
'Call Generate_Pivot_Table_04_ACH

Call Matching_Pivot_Table
End Sub


Sub Act_GLACH(control As IRibbonControl)

Call Process_ACH1115_Data

Call Generate_Pivot_Table_01_Clear_Pivot_Sheet_GL_ACH
Call Generate_Pivot_Table_02_GL_ACH
Call Generate_Pivot_Table_04_ACH11151127


'Call Match_GL_ACH1115_by_Array_v3

End Sub

Sub Act_BankACH(control As IRibbonControl)

End Sub

Sub Act_JE_Upload(control As IRibbonControl)
Generate_MMS_JE
End Sub





Sub Act_Items_to_Post(control As IRibbonControl)

End Sub

Sub Act_Coding_Info(control As IRibbonControl)

End Sub



Sub Act_Validation(control As IRibbonControl)

End Sub

Sub Act_InquiryEmail(control As IRibbonControl)

End Sub

Sub Act_MoveToPending(control As IRibbonControl)

End Sub


Sub Act_JE_UploadPending(control As IRibbonControl)

End Sub

Sub Act_RemovePostedPending(control As IRibbonControl)

End Sub

Sub Act_Mapping_Update(control As IRibbonControl)

End Sub

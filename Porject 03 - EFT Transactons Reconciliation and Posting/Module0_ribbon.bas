Attribute VB_Name = "Module0_ribbon"
'Callback for customUI.onLoad
Sub Ribbon_Onload(ribbon As IRibbonUI)
End Sub

Sub Act_ReadSAP(control As IRibbonControl)
Call Read_SAP_File
Call Text_Field_Can_Not_Be_Empty
End Sub

Sub Act_OffSet(control As IRibbonControl)
Call Find_Offset_Items
Call Matching_After_Kyriba


Call Process_Kyriba_Bank_Statement
Call Activate_Offset_Items_to_Read
End Sub

Sub Act_Items_to_Post(control As IRibbonControl)
Call Filter_Items_to_Post
Call Find_Bank_Description
Call Find_Key_Bank_Info_and_Account

Call Format_Items_Sheet_By_Bank_Code
End Sub

Sub Act_Coding_Info(control As IRibbonControl)
Call Find_Mapping_Info_Step1    'find coding info in mapping exceptional sheet and
Call Find_Mapping_Info_Step2    'find profit center
Call Find_Mapping_Info_Step3    'find coding in same sheet
Call Find_Mapping_Info_Step4    'find post key 40 for Debit 50 for credit
Call Find_Mapping_Info_Step5_Email_to_Confirm       'to see if any transaction is to wait for email comfirmation
'Call To_Program_GL_10901_Email_To_To_Confirm
Call Find_Mapping_Info_Step6_Format

'To process FX transaction
Call Find_Mapping_Info_FX_Step1_Initialize_Items_Sheet_FX
Call Find_Mapping_Info_FX_Step2_Process_FX_Coding
End Sub

Sub Act_JE_Upload(control As IRibbonControl)
Call Fill_JE_Template
Call Fill_JE_Template_FX
Call Generate_Daily_JE_File
Worksheets("3 - C-SAP Standard Template").Select
'Call Validation
End Sub

Sub Act_Validation(control As IRibbonControl)
Call Validation
End Sub

Sub Act_InquiryEmail(control As IRibbonControl)
Call Send_Inquiry_Emails_for_Coding
End Sub

Sub Act_MoveToPending(control As IRibbonControl)
Call Move_to_Pending_File
End Sub


Sub Act_JE_UploadPending(control As IRibbonControl)
Call JE_Pending_Ready_to_Post
Call Generate_Adjusting_JE_File
End Sub

Sub Act_RemovePostedPending(control As IRibbonControl)
Call Remove_Posted_Pending_items
End Sub

Sub Act_Mapping_Update(control As IRibbonControl)
Call Mapping_File_Update
End Sub

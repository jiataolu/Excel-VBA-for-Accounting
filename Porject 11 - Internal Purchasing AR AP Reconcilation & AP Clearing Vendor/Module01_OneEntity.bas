Attribute VB_Name = "Module01_OneEntity"
Option Explicit

Sub MSD_Clearing()
Call PAP_for_One_Entity("MSD")
End Sub

Sub SPS_Clearing()
Call PAP_for_One_Entity("SPS")
End Sub

Sub Wellca_Clearing()
Call PAP_for_One_Entity("Well.ca")
End Sub

Sub PAP_for_One_Entity(Entity As String)

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

Call Read_Bank_Statement
Call Add_Entity_in_Bank_Statement(Entity)
Call Read_SAP_FBL5N(Entity)
Call Reconcile_PAP_invoices(Entity)
Call SPS_Discount_Info(Entity)
Call Reconcile_Amount_PAP_with_Bank_Statement(Entity)
Call Make_Validation_Sheet(Entity)
Call Output_report(Entity)

Worksheets("Validation").Select
Cells(1, 1).Select

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

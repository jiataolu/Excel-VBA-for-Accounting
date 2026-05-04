Attribute VB_Name = "Module90_PublicVar"
Option Explicit

Public Const SubFolderInput As String = "01 - Input"
Public Const SubFolderOutput As String = "02 - Output"
Public Const SubFolderMapping As String = "03 - Mapping & Other Supporting Files"

Public Const FileNameSAP As String = "SAP.XLSX"
'Public Const FileNameBS As String = "Statements_Bank"
Public Const FileNameBS As String = "reportcontent"
Public Const FileNameDailyJE As String = "01 - Daily JE.xlsx"
Public Const FileNameAdjustingJE As String = "02 - Adjusting JE.xlsx"
Public Const FileNamePending As String = "Pending Transactions.xlsx"

Public Const SuspenseAccount As String = "18970"
'Public Const SuspenseAccount As String = ""
Public Const WaitToConfirmInfo As String = "(Email to Confirm)"
Public Const WaitToConfirmGL As String = "10901"


Public Const iColSAPPostingDate As Integer = 2
Public Const iColSAPDocNumber As Integer = 3
Public Const iColSAPGL As Integer = 4
Public Const iColSAPAss As Integer = 5
Public Const iColSAPText As Integer = 6
Public Const iColSAPAMT As Integer = 7
Public Const iColSAPPostKey As Integer = 10
Public Const iColSAPClear As Integer = 11

Public Const iColItemsPostingDate As Integer = 1
Public Const iColItemsDocNumber As Integer = 2
Public Const iColItemsGL As Integer = 3
Public Const iColItemsAMT As Integer = 4
Public Const iColItemsBankInfo As Integer = 5
Public Const iColItemsKeyBankAccount As Integer = 6
Public Const iColItemsPostBU As Integer = 7
Public Const iColItemsPostGL As Integer = 8
Public Const iColItemsPostVendor As Integer = 9
Public Const iColItemsPostProfitC As Integer = 10
Public Const iColItemsPostKeyCode As Integer = 11
Public Const iColItemsPostAssInfo As Integer = 12
Public Const iColItemsPostCostCenter As Integer = 13



Public Const iColItemsPostCurrency As Integer = 14
Public Const iColItemsFXAmt As Integer = 15
Public Const iColItemsFXBU As Integer = 16
Public Const iColItemsFXGL As Integer = 17
Public Const iColItemsFXVendorCode As Integer = 18
Public Const iColItemsFXProfitC As Integer = 19
Public Const iColItemsFXKeyCode As Integer = 20
Public Const iColItemsFXAssInfo As Integer = 21
Public Const iColItemsFXCostCenter As Integer = 22


'Public Const iColBSInsertNetAmt As Integer = 7

Public Const iColBSAMT As Integer = 14
Public Const iColBSDepositAMT As Integer = 7
Public Const iColBSPaymentAMT As Integer = 6

Public Const iColBSBankCode As Integer = 1
Public Const iColBSComment As Integer = 10
Public Const iColBSBooked As Integer = 15


Public Const iColConcenClear As Integer = 2
Public Const iColConBankCode As Integer = 3

Public Const iColMapExcBankAcct As Integer = 2

Public Const iColMapType As Integer = 1
Public Const iColMapBankAcct As Integer = 2
Public Const iColMapBU As Integer = 7
Public Const iColMapGL As Integer = 8
Public Const iColMapCurrency As Integer = 5


Public Const iColProfitCBU As Integer = 1
Public Const iColProfitCGL As Integer = 2
Public Const iColProfitCPC As Integer = 3


Public Const iColPendingPostingDate As Integer = 1
Public Const iColPendingDocNumber As Integer = 2
Public Const iColPendingGL As Integer = 3
Public Const iColPendingAMT As Integer = 4
Public Const iColPendingBankInfo As Integer = 5
Public Const iColPendingKeyBankAcct As Integer = 6
Public Const iColPendingPostBU As Integer = 7
Public Const iColPendingPostGL As Integer = 8
Public Const iColPendingPostVendor As Integer = 9
Public Const iColPendingPostProfitCenter As Integer = 10
Public Const iColPendingPostKeyCode As Integer = 11
Public Const iColPendingPostAssInfo As Integer = 12
Public Const iColPendingPostCostCenter As Integer = 13
Public Const iColPendingJEPosted As Integer = 14

Public Const MainCompanyCode As Integer = 9000
Public Const MainGLFX As Long = 67023

Attribute VB_Name = "Module92_Public_Var"
Option Explicit

Public Const NameHD As String = "jlu"
Public Const NameHTTP As String = "jiatao_lu"

Public Const FileNameCashPosition As String = "To read 01 - Treasury Report.xlsx"
Public Const FileNameFCCS As String = "To read 02 - FCCS-SAP-JDE-NetSuite"


'CompanyCode file
Public Const FileNameCompanyCode = "To read 03 - Company Code.xlsx"
Public Const ColCompanyCodeCompanyCode As Integer = 1
Public Const ColCompanyCodeParentCode As Integer = 4
Public Const ColCompanyCodeBUName As Integer = 5
Public Const ColCompanyCodeVendorCode As Integer = 7
Public Const ColCompanyCodeERP As Integer = 16


'PeopleSoft File
Public Const FileNamePS As String = "To read 04 - PeopleSoft GL vs SAP mapping.xlsx"
Public Const ColPSBankName As Integer = 1
Public Const ColPSBankAcct As Integer = 3
Public Const ColPSProductCode As Integer = 7
Public Const ColPSBUCode As Integer = 9
Public Const ColPSSapGL As Integer = 10
Public Const ColPSRemark As Integer = 11



'Public Const SheetNameBankShortName As String = "Bank Abbreviation"
'Public Const ColBankShortNameBankName As Integer = 1

Public Const FIS4Header = "FISCodeKyribaCodeBUFISBUSAP"       'header information of 4 columns, with empty spaces removed

Public Const SheetNameFIS As String = "FIS & PeopleSoft"
Public Const ColFISFISCode As Integer = 1
Public Const ColFISKyribaCode As Integer = 2
Public Const ColFISBUCode As Integer = 3
Public Const ColFISSapGL As Integer = 4
Public Const ColFISCompanyName As Integer = 5
Public Const ColFISBankAcct As Integer = 6
Public Const ColFISCurrency As Integer = 7
Public Const ColFISProductCode As Integer = 8
Public Const ColFISIsinFIS As Integer = 9
Public Const ColFISIsinPS As Integer = 10
Public Const ColFISRemark As Integer = 11
Public Const ColFISKeyNumber As Integer = 12


'Sheet Mapping Consolidation
' initially, in Mapping sheet, there are two BU code column. Now one column "SAP BU" is not required.
' But we keep the vailabe name of FISBUCode for SAP BU Code0
Public Const SheetNameMapping As String = "Mapping Consolidated"
Public Const ColMapCompanyName As Integer = 1
Public Const ColMapBankAcctFull As Integer = 2
Public Const ColMapFISCode As Integer = 3
Public Const ColMapKyribaCode As Integer = 4
Public Const ColMapCry As Integer = 5
Public Const ColMapERPSystem As Integer = 6
Public Const ColMapFISBUCode As Integer = 7
'Public Const ColMapSAPBUCode As Integer = 8
Public Const ColMapFISSapGL As Integer = 8
Public Const ColMapLocalBU As Integer = 9
Public Const ColMapLocalGL As Integer = 10
Public Const ColMapBUName As Integer = 11
Public Const ColMapVendorCode As Integer = 12
Public Const ColMapParentCode As Integer = 13
Public Const ColMapProductCode As Integer = 14
Public Const ColMapDataSource As Integer = 15
Public Const ColMapOwnership As Integer = 16
Public Const ColMapComment As Integer = 17
Public Const ColMapRemark As Integer = 18
Public Const ColMapBankAcctKey As Integer = 19


Public Const SheetNameDeleted As String = "Deleted"
Public Const ColDeletedDeletedData As Integer = 19


Public Const SheetNameBUError As String = "BU Error"


Public Const FileNameRecon As String = "To read 05 - BlackLine Reconciliation List.xlsx"
Public Const ColReconBizUnit As Integer = 3
Public Const ColReconAccount As Integer = 5
Public Const ColReconTEAM As Integer = 6
Public Const ColReconReviewer As Integer = 7
Public Const ColReconApprover As Integer = 8
Public Const ColReconPreparer As Integer = 9
Public Const ColReconBU As Integer = 19
Public Const ColReconGL As Integer = 20
Public Const ColReconComboBUGL As Integer = 21


Public Const LenKeyBankAcctNo As Integer = 9


Public LineDeleted As Integer
Public LineNew As Integer



'Public Const iColMapBUName As Integer = 12



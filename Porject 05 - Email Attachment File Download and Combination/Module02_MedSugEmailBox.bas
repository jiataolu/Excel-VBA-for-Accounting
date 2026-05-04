Attribute VB_Name = "Module02_MedSugEmailBox"
Option Explicit

Sub Go_Through_MedSug_Email_Folder_and_Download()

'Call Initialize_Consolidated_CSV_file

Dim response As VbMsgBoxResult
    
' Display the message box with Yes and No options
response = MsgBox("Make sure folder " & Chr(34) & "Download Files - EFT Payment" & Chr(34) & " is empty." & Chr(10) & "Do you want to continue?", vbYesNo + vbQuestion, "Confirmation")
    
' Check the user's response
If response = vbNo Then Exit Sub


Dim outlookApp As Object
Dim namespace As Object


Dim rootFolder As Object
Dim botInBox12YearFolder As Object
Dim reportingFolder As Object
Dim medSurgFolder As Object


Dim items As Object
Dim email As Object
Dim oldReceivedDate As Date
Dim receivedDate As Date
Dim emailSubject As String
Dim startDate As Date
Dim endDate As Date
Dim attachment As Object
Dim categoryName As String

Dim iAttachCount As Integer
Dim sAttachmentNumber As String
Dim sAttachFileName As String
Dim sFullFileAttachSave As String
Dim wsEmail As Worksheet
Dim iRow As Long
Dim iCol As Integer
Dim sAttachSavePath As String

Dim today As Date
today = Date

' Set the date range
oldReceivedDate = "01/01/2000"
'startDate = DateAdd("m", -3, today)
'startDate = DateValue("12/21/2024")
'endDate = DateValue("11/18/2024")

startDate = InputBox("Please enter starting date MM/DD/YYYY:")

    
' Initialize Outlook application
Set outlookApp = CreateObject("Outlook.Application")
Set namespace = outlookApp.GetNamespace("MAPI")

' Specify the folder to search emails in (e.g., Inbox)
Set rootFolder = namespace.Folders("MVT Accounting Bank and Cash") ' Change this if needed
Set botInBox12YearFolder = rootFolder.Folders("Bot_Inbox-12Year")
Set reportingFolder = botInBox12YearFolder.Folders("Reporting")
Set medSurgFolder = reportingFolder.Folders("MedSurg")
'Set medSurgFolder = reportingFolder.Folders("00-Lu Macro Test")
    
    
' Get all items in the folder
Set items = medSurgFolder.items
    
' Sort items by received date in ascending order
items.Sort "[ReceivedTime]", True

' Specify the category name to be assigned
categoryName = "macro_process"


sAttachSavePath = GetWorkPath & "\Download Files - EFT Payment"
Set wsEmail = Worksheets("Log")
wsEmail.Select
Cells.Select
Selection.Delete
Cells(1, 1).Select
wsEmail.Cells(1, 1) = "Email Date"
wsEmail.Cells(1, 2) = "Email Subject"
wsEmail.Cells(1, 3) = "Attachment"
wsEmail.Cells(1, 4) = "Total AMT"


iRow = 1
iAttachCount = 0
' Loop through each email in the folder
For Each email In items
    ' Check if the item is a mail item
    If email.Class = olMail Then
        receivedDate = email.ReceivedTime
        If receivedDate <> oldReceivedDate Then
            iAttachCount = 0
            oldReceivedDate = receivedDate
        End If
        
        emailSubject = UCase(Replace(email.Subject, " ", ""))
        
        ' Check if the received date is within the specified range
        'If receivedDate <= startDate - 1 Or receivedDate >= endDate + 1 Then GoTo PROCESSNEXTEMAIL
        If receivedDate <= startDate - 1 Then GoTo PROCESSNEXTEMAIL

        If InStr(emailSubject, "SECURE:EDI") = 0 Or InStr(emailSubject, "EFTPAYMENT") = 0 Then GoTo PROCESSNEXTEMAIL
        If email.Categories = categoryName Then
            'Debug.Print "already processed : " & receivedDate & "-" & emailSubject
            GoTo PROCESSNEXTEMAIL
        End If
            
        'Debug.Print receivedDate & "-" & email.Subject

        email.Categories = categoryName
        email.Save
            


        If email.Attachments.Count > 0 Then
            'iAttachCount = 0
            'iCol = 2
            For Each attachment In email.Attachments
                sAttachFileName = attachment.Filename
                
                If InStr(UCase(sAttachFileName), "CSV") > 0 Then
                    iAttachCount = iAttachCount + 1
                    'iCol = iCol + 1
                    ' Save each attachment
                    
                    sAttachmentNumber = "A" & Format(CStr(iAttachCount), "00")
                    sFullFileAttachSave = sAttachSavePath & "\" & Year(receivedDate) & Format(Month(receivedDate), "00") & Format(Day(receivedDate), "00") & sAttachmentNumber & "-" & attachment.Filename
                
                    attachment.SaveAsFile sFullFileAttachSave
                    ' Print confirmation to the Immediate Window
                    'wsEmail.Cells(iRow, iCol) = sFullFileAttachSave
                    'Debug.Print "Attachment " & attachment.Filename & " saved."
                    
                    Call Read_CSV_Single(sFullFileAttachSave, Format(receivedDate, "YYYYMMDD"), sAttachmentNumber)
                    iRow = iRow + 1
                    wsEmail.Cells(iRow, 1) = Year(receivedDate) & Format(Month(receivedDate), "00") & Format(Day(receivedDate), "00")
                    wsEmail.Cells(iRow, 2) = email.Subject
                    wsEmail.Cells(iRow, 3) = sAttachmentNumber
                    wsEmail.Cells(iRow, 4) = CSVTotalAmount
                    
                    
                End If  ' If InStr(UCase(sAttachFileName), "CSV") > 0
            Next attachment
            
        End If  ' if email.Attachments.Count > 0
            
    End If  'end of email.Class = olMail
    
PROCESSNEXTEMAIL:
Next email
    
wsEmail.Select
Columns.AutoFit

' Release objects
Set attachment = Nothing
Set email = Nothing
Set items = Nothing
Set medSurgFolder = Nothing
Set reportingFolder = Nothing
Set botInBox12YearFolder = Nothing
Set rootFolder = Nothing
Set namespace = Nothing
Set outlookApp = Nothing
Set wsEmail = Nothing

End Sub




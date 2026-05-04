Attribute VB_Name = "Module18_InquiryEmail"
Option Explicit

Sub Send_Inquiry_Emails_for_Coding()
Dim wsItems As Worksheet
Dim iMaxRowItems As Integer
Dim iRowItems As Integer
Dim sBU As String
Dim sPostingDate As String
Dim sGL As String
'Dim dAMT As Double
Dim sAMT As String
Dim sDocNumber As String
Dim sBankDescription As String


Dim lRealLastRow As Long
Dim lRealLastCol As Long


Set wsItems = Worksheets("2-Items to post")
wsItems.Select
lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).Row
lRealLastCol = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
iMaxRowItems = lRealLastRow

If iMaxRowItems < 2 Then Exit Sub

For iRowItems = 2 To iMaxRowItems
'For iRowItems = 3 To 3
    sBU = wsItems.Cells(iRowItems, iColItemsPostBU)
    
    If Replace(sBU, " ", "") = "" Or InStr(sBU, WaitToConfirmInfo) > 0 Then
        sPostingDate = wsItems.Cells(iRowItems, iColItemsPostingDate)
        sGL = wsItems.Cells(iRowItems, iColItemsGL)
        sAMT = CStr(wsItems.Cells(iRowItems, iColItemsAMT))
        sDocNumber = CStr(wsItems.Cells(iRowItems, iColItemsDocNumber))
        sBankDescription = wsItems.Cells(iRowItems, iColItemsBankInfo)

        Call Send_Email_One_Line(sPostingDate, "9000", sGL, sAMT, sBankDescription)
    End If
Next iRowItems



End Sub


Sub Send_Email_One_Line(PostingDate As String, BU As String, GL As String, Amount As String, BankDescription)

Dim OutApp As Object  ' Outlook.Application
Dim OutMail As Object  ' Outlook.MailItem
Dim sEmailTo As String
Dim sMessage As String
Dim sSubject As String
Dim sSignature As String

Dim sPayorDepSubject As String
Dim sPayorDep As String
Dim sABSAmount As String

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

sABSAmount = CStr(Abs(CDbl(Amount)))

sEmailTo = ""

sSubject = Format(PostingDate, "MM/DD/YYYY") & ", " & "BU-" & BU & ", " & "GL-" & GL & ", "

If CDbl(Amount) < 0 Then

    sSubject = "Payment: " & sSubject & "(" & Format(sABSAmount, "$#,##0.00") & ")"
    sPayorDep = "have this payment. "
Else
    
    sSubject = "Deposit: " & sSubject & Format(sABSAmount, "$#,##0.00")
    sPayorDep = "receive this deposit. "
End If



sMessage = "<HTML>  <BODY style=font-size:11pt; font-family:Calibri >"
sMessage = sMessage & "<P> Good day, </P><P>We "
sMessage = sMessage & sPayorDep & "Detail is as following:</P><P></P>"
sMessage = sMessage & BankDescription
sMessage = sMessage & "<P>Could you please kindly check and provide coding information? Thank you!</P><P></P>"

sSignature = "<br><font color=Brown size=3><P>Sanjay Kotia<P><HTML>"
sMessage = sMessage & sSignature


With OutMail
    .To = sEmailTo
    .htmlbody = sMessage
    .Subject = sSubject
    '.attachments.Add sDocFileName
    '.attachments.Add sPDFFileName
    .Display
    .Save
    .Close olpromptforsave
End With
Set OutMail = Nothing
Set OutApp = Nothing

End Sub



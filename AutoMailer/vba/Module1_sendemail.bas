Attribute VB_Name = "Module1_sendemail"
Option Explicit
'====================================================
' Project     : Auto Mail Sending Utility
' Module      : Send Email
' Description : Sends emails to recipients listed in the "Data" sheet of the workbook using Outlook.
' Author      : Balaji
' Created     : 2026-10-18
' Version     : 1.0.0
' Last Update : 2026-10-18
'====================================================
Dim i As Long
Dim Sheet1 As Worksheet
Dim oMailMessage As Object
Dim oWkb As Workbook, oWks As Worksheet, lastrw As Long
Dim oMail As Object
Dim fd As FileDialog
Dim OutApp As Outlook.Application
Dim OutNS As Outlook.Namespace
Dim OutMail As Outlook.MailItem
Dim OutInspector As Outlook.Inspector
Dim SMTP_Config As Object, rng As Range, RWks As Worksheet, sTemplatepath As String
Dim S_MAILSUBJECT As String, S_FROMADDRESS As String, S_TOADDRESS As String, S_MAILBODY As String, S_CCADDRESS As String, S_BCCADDRESS As String
Dim I_IFR As Long


Public Sub Addattachement()

 Dim myAttachments As Outlook.Attachments


 Set myAttachments = OutMail.Attachments
 myAttachments.Add ThisWorkbook.Path & "\" & "InputFiles\DRAFT Digicel FT Comms - Statement Rec request Final.docx", olByValue, 1, "Attach"
 OutMail.Display

End Sub

Sub SendEmail()
    Application.ScreenUpdating = False
    sTemplatepath = "C:\Users\HP\AppData\Roaming\Microsoft\Templates"
        
        Set oWkb = ThisWorkbook
        Set oWks = oWkb.Sheets("Data")
        oWks.Activate
     '   Set RWks = oWkb.Sheets("Sheet2")
        lastrw = oWks.Range("A" & Rows.Count).End(xlUp).Row
        oWks.Range("A1:Z" & lastrw).AutoFilter Field:=25, Criteria1:="", Operator:=xlFilterValues
        oWks.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).Select
        I_IFR = ActiveCell.Row
        lastrw = oWks.Range("A" & Rows.Count).End(xlUp).Row
        For i = I_IFR To lastrw

            S_TOADDRESS = oWks.Range("S" & i)
            S_MAILSUBJECT = "Reference to your Profile posted on " & oWks.Range("X" & i) & " " & oWks.Range("C" & i) & " - " & oWks.Range("B" & i)
            
            'S_MAILBODY = oWks.Range("C" & i) & "/" & oWks.Range("Q" & 1) & "/" & oWks.Range("Q" & i)
                Set OutApp = CreateObject("Outlook.Application")
            
                Set OutNS = OutApp.GetNamespace("MAPI")
                OutNS.Logon
                    ' Set OutMail = OutApp.CreateItem(olMailItem)
                Set OutMail = OutApp.CreateItemFromTemplate(sTemplatepath & "\" & "Sai Sankara - Matrimony ID - 12442.oft")
                ' Addattachement
          
        
                With OutMail
                           .To = S_TOADDRESS
                         '  .CC = oWks.Range("T2") & ";" & oWks.Range("T3") & ";" & oWks.Range("T4") & ";" & oWks.Range("T5")
                           .Subject = S_MAILSUBJECT
                          ' .Body = S_MAILBODY
                           .Display
'                           .Send
                End With
                Application.Wait (Now + TimeValue("0:00:02"))
    
    oWks.Range("Y" & i) = "MailSent"
    oWks.Range("Z" & i) = Now
    Set OutMail = Nothing
    Set OutApp = Nothing
Next i
            
 oWks.ShowAllData
End Sub

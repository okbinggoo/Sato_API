
Imports System.Net.Mail
Imports KeepingList.fileOI
Public Class mail

    Public Shared Function sendmail(ByVal subject As String, ByVal body As String, ByVal functionName As String) As Boolean
        'Dim mailInstance As MailMessage = New MailMessage("connect-admin@murata.com", "xx", subject, body)
        Dim mailInstance As MailMessage = New MailMessage
        mailInstance.From = New MailAddress("GLSJOB-" & functionName & "@murata.com")

        Dim picTO As String()
        Dim picCC As String()
        Dim picBCC As String()

        Dim hashConf As Hashtable = readConfigFile()

        picTO = Split(hashConf.Item("mailTO"), ":")
        picCC = Split(hashConf.Item("mailCC"), ":")
        picBCC = Split(hashConf.Item("mailBCC"), ":")

        mailInstance.Subject = subject & "  >>  PLEASE CHECK DETAIL"
        mailInstance.Body = "GLS JOB IS ERROR<br><br><br>" & body & "<br><br><br>" & "Location Of Program >>> " & hashConf.Item("LOCATION_PGM")

        For Each item In picTO
            If item.ToString <> "" Then
                mailInstance.To.Add(item)
            End If
        Next
        For Each item In picCC
            If item.ToString <> "" Then
                mailInstance.CC.Add(item)
            End If
        Next
        For Each item In picBCC
            If item.ToString <> "" Then
                mailInstance.Bcc.Add(item)
            End If
        Next

        mailInstance.Priority = MailPriority.High
        mailInstance.IsBodyHtml = True
        'mailInstance.Attachments.Add(New Attachment("filename")) 'Optional  
        Dim mailSenderInstance As SmtpClient = New SmtpClient("172.24.128.80", 25) '25 is the port of the SMTP host  
        mailSenderInstance.Credentials = New System.Net.NetworkCredential("LoginAccout", "Password")
        mailSenderInstance.Send(mailInstance)
        mailInstance.Dispose() 'Please remember to dispose this object  
        sendmail = True
    End Function

End Class

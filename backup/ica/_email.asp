<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->  
<%
Dim objConfig   'As CDO.Configuration

Function PrepareConfigCdo()
	Dim Fields      'As ADODB.Fields
	Dim iResult
	On Error Resume Next
	
	' Get a handle on the config object and it's fields
	Set objConfig = CreateObject("CDO.Configuration")
	iResult = Err.Number
	iResult = iResult + Err.Number

	' Set config fields
	If iResult = 0 Then
		With objConfig.Fields
		    .Item(cdoSendUsingMethod) = cdoSendUsingPort 'cdoSendUsingPickup
		    .Item(cdoSMTPServer) = "mail.ibf.be"
		    .Item(cdoSMTPServerPort) = 25
		    .Item(cdoSMTPConnectionTimeout) = 60
		    .Item(cdoSMTPAuthenticate) = cdoNTLM
		    .Update
		    iResult = iResult + Err.Number
		End With

		If iResult<>0 Then
		With objConfig.Fields
		    .Item(cdoSendUsingMethod) = cdoSendUsingPickup
		    .Item(cdoSMTPServer) = "mail.ibf.be"
		    .Item(cdoSMTPServerPort) = 25
		    .Item(cdoSMTPConnectionTimeout) = 60
		    .Item(cdoSMTPAuthenticate) = cdoNTLM
		    .Update
		    iResult = iResult + Err.Number
		End With
		End If 
	End If
	On Error GoTo 0
PrepareConfigCdo=iResult
End Function


' Sending email via Exchange (CDO)
Function SendEmailViaCdo(sFrom, sTo, sSubject, sBody, sBcc)
	Dim objMessage  'As CDO.Message

	Dim iResult
	Dim sResultInfo

	'On Error Resume Next

	Set objMessage = CreateObject("CDO.Message")
	iResult = iResult + Err.Number

	If iResult = 0 Then
	Set objMessage.Configuration = objConfig

	With objMessage
	    '.DSNOptions = cdoDSNSuccessFailOrDelay
	    .MIMEFormatted = True

	    .From = sFrom
	    .To = sTo
	    If sBcc > "" Then
	        .BCC = sBcc
	    End If
	    .Subject = sSubject
	    .HTMLBody = sBody
	    .Send
	    
	    iResult = iResult + Err.Number
	End With

	End If

	Set objMessage = Nothing

	'On Error GoTo 0

	SendEmailViaCdo = iResult
End Function


' Sending email via Exchange (CDO)
Function SendEmailViaCdoSender(sSender, sFrom, sTo, sSubject, sBody, sBcc)
	Dim objMessage  'As CDO.Message

	Dim iResult
	Dim sResultInfo

	'On Error Resume Next

	Set objMessage = CreateObject("CDO.Message")
	iResult = iResult + Err.Number

	If iResult = 0 Then
	Set objMessage.Configuration = objConfig

	With objMessage
	    '.DSNOptions = cdoDSNSuccessFailOrDelay
	    .MIMEFormatted = True

	    .Sender = sSender
	    .From = sFrom
	    .To = sTo
	    If sBcc > "" Then
	        .BCC = sBcc
	    End If
	    .Subject = sSubject
	    .HTMLBody = sBody
	    .Send
	    
	    iResult = iResult + Err.Number
	End With

	End If

	Set objMessage = Nothing

	'On Error GoTo 0

	SendEmailViaCdoSender = iResult
End Function


Function SendEmailViaCdoNts(sSenderAddress, sReceiverAddress, sSubjectBody, sMessageBody, sMessageType)
Dim objNewMail
	Set objNewMail = Server.CreateObject("CDONTS.NewMail")

	objNewMail.BodyFormat = CdoBodyFormatHTML
	objNewMail.MailFormat = CdoMailFormatMime

	objNewMail.From = sSenderAddress
	objNewMail.To = sReceiverAddress
	objNewMail.Subject = sSubjectBody

	objNewMail.Body = sMessageBody
	objNewMail.Send

	Set objNewMail = Nothing
End Function

Function SendEmail(sSenderAddress, sReceiverAddress, sSubjectBody, sMessageBody, sMessageType)

	PrepareConfigCdo
	SendEmailViaCdo sSenderAddress, sReceiverAddress, sSubjectBody, sMessageBody, ""
	
End Function  

Function SendEmailWithSender(sSenderAddress, sFromAddress, sReceiverAddress, sSubjectBody, sMessageBody, sMessageType)

	PrepareConfigCdo
	SendEmailViaCdoSender sSenderAddress, sFromAddress, sReceiverAddress, sSubjectBody, sMessageBody, ""
	
End Function  
%>
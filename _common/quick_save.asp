<%
'--------------------------------------------------------------------
'
' CV registration.
' Short format. Saving all the data.
'
'--------------------------------------------------------------------
%>
<!--#include file="_data/datPsnTitle.asp"-->
<% 
'If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
'End If
CheckExpertID
%>
<!--#include file="../_common/expProfile.asp"-->
<%
Dim objUploadForm
Dim sUserLogin, sUserPassword, sUserPhone
Dim sFlagSelected, j
Dim objResult, iResult

Dim bEmailAlreadySent
%>
<html>
<head>
<title>Quick CV registration</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<% 
Set objUploadForm = Server.CreateObject("softartisans.fileup")
If objUploadForm.ContentDisposition = "form-data" Then
	If objUploadForm.TotalBytes > 3100000  Then 
		ShowMessageStart "error", 580 %>
			Your file is too big.</b><br>Please try to keep the size of the file within the allowed 3 Mb. Click back and try again.
		<% ShowMessageEnd 
	Else

	objResult=SaveExpertShortProfile(iExpertID, objUploadForm.Form)

	If sApplicationName="expert" Then
		If objResult(0)=0 Then
			If objResult(1)>0 Then
				iExpertID=objResult(1)
				iUserID=objResult(2)
				sUserLogin=objResult(3)
				sUserPassword=objResult(4)
	
				' Login active user
				objTempRs2=UpdateRecordSP("usp_LogSessionUser", _
					Array(Array(, adVarChar, 40, sSessionID), Array(, adInteger, , iUserID)))
			End If
		End If
	End If
	Set objResult=Nothing
	
	' Saving the file with CV
	iResult=SaveExpertDocument(iExpertID, "cv", objUploadForm.Form)
	
	If bEmailAlreadySent=False Then
		'PrepareEmailTemplate "emlExpCvvRegister", ";;sExpertFullName=" & sTitleLastName & ";;sUserLogin=" & sUserLogin & ";;sUserPassword=" & sUserPassword
		'SendEmail "info@assortis.com", sUserEmail, sEmailSubject, sEmailBody, "info"
		'SendEmailTemplate "assortissupport@ibf.be", "emlMntCvvRegister"
	End If
	
	If sApplicationName="expert" Then
		ShowMessageStart "info", 550 %>
			<b>Resume Management System</b><br><br>
			Thank you for having registered your CV. You will receive an electronic confirmation that your <span class="rs">CV has been successfully encoded</span> to the email address you currently have registered with us which is  
			&nbsp;<b><%=objUploadForm.form("exp_email")%></b>. <br><br>Your <b>username</b> and <b>password</b> will be contained within this email and will allow you to access your online profile.
			<br><br><br><img src="<% =sHomePath %>image/bte_continue.gif" border="0">
		<% ShowMessageEnd
	Else
		ShowMessageStart "info", 550 %>
			<b>Resume Management System</b><br><br>
			Thank you for having registered this CV. 
			<br><br><br><a href="<% =sApplicationHomePath %>"><img src="<% =sHomePath %>image/bte_continue.gif" border="0"></a>
		<% ShowMessageEnd
	End If

End If
End If
%>

<% CloseDBConnection %>
</body></html>

<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit 
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" %>
<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->
<!--#include file="_forms/frmInterface.asp"-->
<%
Dim sUserLogin, sUserPassword, sUserAccountDetails, bCorrectEmail

sUserEmail=CheckString(Request.Form("user_email"))
bCorrectEmail=0

If sUserEmail>"" Then
	Set objTempRs=GetDataRecordsetSP("usp_UsrPasswordSelect", Array( _
		Array(, adVarWChar, 250, sUserEmail)))
	If objTempRs.Eof Then
		bCorrectEmail=0
	Else
		bCorrectEmail=1
		sUserAccountDetails=""
		sUserLogin=objTempRs(0)
		sUserPassword=objTempRs(1)

		sUserAccountDetails=sUserAccountDetails & "Login: <b>" & sUserLogin & "</b><br>"
		sUserAccountDetails=sUserAccountDetails & "Password: <b>" & sUserPassword & "</b><br>"
		sUserAccountDetails=sUserAccountDetails & "<br>"

		PrepareEmailTemplate "emlUsrForgotMultiPassword.htm", ";;sUserAccountDetails=" & sUserAccountDetails & ";;sUserIpAddress=" & sUserIpAddress
		SendEmail sEmailCvipSystem, sUserEmail, sEmailSubject, sEmailBody, "info"
	End If
	objTempRs.Close
	Set objTempRs=Nothing
End If   
%>

<html>
<head>
<title>Forgot your password?</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>

	<% If bCorrectEmail=1 Then
		ShowMessage "Your login and password has now been sent to your e-mail address. Check your email and click on continue.", "info", 450
		%>
		<br>
		<div align="center"><a href="login.asp<% =sParams %>"><img src="<% =sHomePath %>image/bte_continue.gif" name="Login" border=0 alt="Continue" vspace=3></div>
	<% Else
		If sUserEmail>"" Then
			ShowMessage "The e-mail address you supplied is not available in our database.", "error", 450
		End If

		ShowInputFormHeader 450, "FORGOTTEN PASSWORD OR LOGIN" %>
		<form method="post" action="<% =sScriptFileName %><%=sParams%>" name="LoginForm">
		<% ShowInputFormElement 450, "E-mail address", "text", "user_email", "", 50, 255, 0, ""
		ShowInputFormButton 450, "Submit", sHomePath & "image/bte_submit.gif" %>
		</form>
		<% ShowInputFormFooter 450
	End If %>

</body>
</html>
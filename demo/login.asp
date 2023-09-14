<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit 
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" %>
<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->
<!--#include file="_forms/frmInterface.asp"-->
<%
Dim sLoginName, sLoginPassword, sLoginEmail
If Request.Form("")>"" Then
	sLoginName=CheckString(Request.Form("login_name"))
	sLoginPassword=CheckString(Request.Form("login_pwd"))
	sLoginEmail=CheckString(Request.Form("login_email"))
End If

Dim sMessage
Dim sCameFromForgotPassword
sCameFromForgotPassword=Request.QueryString("forgot")

If InStr(sScriptFullName, "apply/")>0 Then
	sUserType="expert"
Else
	sUserType="staff"
End If
%>
<html>
<head>
<title>Login</title>
<meta name="robots" content="noarchive">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<%=sHomePath%>/styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
function validateForm() {
	var f=document.forms[0];
	if (!(f)) {
		return false; }
	if (!checkTextFieldValue(f.login_name, "", "<% =GetLabel(sCvLanguage, "Please fill in your login name") %>", 1)) { return false }
	if (!checkTextFieldValue(f.login_pwd, "", "<% =GetLabel(sCvLanguage, "Please fill in your password") %>", 1)) { return false }
	f.submit();
}
</script>
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<!-- Login -->
	<% If sError="NoMatch" then
		sMessage=GetLabel(sCvLanguage, "The login or password you supplied are not correct")
		ShowMessage sMessage, "error", 450
	ElseIf sCameFromForgotPassword="ok" Then
		sMessage=GetLabel(sCvLanguage, "Your login and password has now been sent")
		ShowMessage sMessage, "info", 450
	ElseIf sUserType="expert" Then
		sMessage=GetLabel(sCvLanguage, "Please enter your login name and password") & "."
		ShowMessage sMessage, "info", 450
	Else
		sMessage=GetLabel(sCvLanguage, "Please enter your login name and password") & "."
		ShowMessage sMessage, "info", 450
	End If
	%>
	
	<%
	ShowInputFormHeader 360, "LOGIN"
	%>
	<form method="post" action="<%=sHomePath %>login_do.asp<% =sParams %>" name="LoginForm" onSubmit="validateForm(); return false;">
	<%
	ShowInputFormSpacer 360, 1
	ShowInputFormElement 360, GetLabel(sCvLanguage, "Login name"), "text", "login_name", sLoginName, 100, 200, 0, ""
	ShowInputFormElement 360, GetLabel(sCvLanguage, "Password"), "password", "login_pwd", "", 100, 200, 0, ""
	ShowInputFormSpacer 360, 1
	ShowInputFormButton 360, "Login", sHomePath & "image/bte_login.gif"
	%>
	</form>
	<%
	ShowInputFormFooter 360 
	%>

	<% If sUserType="expert" Then %>
		<% If sCameFromForgotPassword<>"ok" Then %>
		<br><div style="text-align: center;"><div style="width:450px; margin-left:auto;  margin-right:auto; text-align: left;">
		<p class="txt"><img src="<% =sHomePath %>image/x.gif" width=35 height=1 align="left"><a href="login_forgot.asp<%=sParams%>"><% =GetLabel(sCvLanguage, "Forgot your login or password?") %></a></p>
		</div></div><br>
		<% End If %>

	<% End If %>	
	
<% CloseDBConnection %>
</body>
</html>

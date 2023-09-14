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
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!--
function validateForm() {
	var f=document.forms[0];
	if (!(f)) {
		return false; }
	if (!checkTextFieldValue(f.login_name, "", "Please fill in your login name", 1)) { return false }
	if (!checkTextFieldValue(f.login_pwd, "", "Please fill in your password", 1)) { return false }
	f.submit();
}
//-->
</script>
</head>
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">

<!-- Login -->
	<% If sError="NoMatch" then
		sMessage="The login or password you supplied are not correct. They are case sensitive: check the CAPS LOCK key. Please verify the punctuation and spaces as well."
		ShowMessage sMessage, "error", 450
	ElseIf sCameFromForgotPassword="ok" Then
		sMessage="Your login and password has now been sent to your <nobr>email</nobr> address. <b>Please check your email and login.</b>"
		ShowMessage sMessage, "info", 450
	ElseIf sUserType="expert" Then
		sMessage="Please enter your login name and password."
		ShowMessage sMessage, "info", 450
	End If
	%>
	
	<form method="post" action="<%=sHomePath %>login_do.asp<% =sParams %>" name="LoginForm" onSubmit="validateForm(); return false;">
	<%
	ShowInputFormHeader 360, "Login"
	%>
	<form method="post" action="<%=sHomePath & "login.asp" & sParams%>" name="LoginForm" onSubmit="Login(); return false;">
	<%
	ShowInputFormElement 360, "User&nbsp;name", "text", "login_name", "", 50, 280, 0, ""
	ShowInputFormElement 360, "Password", "password", "login_pwd", "", 50, 280, 0, "class=""last"""
	ShowInputFormFooter 360 
	%>
	<div class="button first login">
		<input type="image" class="login" src="../image/bte_login.gif" name="Login" alt="Login">
	</div>
	</form>

</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->

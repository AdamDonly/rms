<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit
Response.Buffer=True
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Confirmation
'
'--------------------------------------------------------------------
%>
<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../_forms/frmInterface.asp"-->
<!--#include file="../../_common/expProfile.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams

Dim sUserLogin, sUserPassword, sUserPhone
Dim bEmailAlreadySent

bEmailAlreadySent=False
%>

<html>
<head>
<title>Curriculum Vitae. Confirmation.</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
	<%
	If bEmailAlreadySent=False Then
		LoadExpertProfile(iExpertID)

		PrepareEmailTemplate "emlExpertRegisteredCV.htm", ";;sExpertFullName=" & sFullName
		SendEmail sEmailCvipSystem, sUserEmail, sEmailSubject, sEmailBody, "info"
	End If
	
	ShowMessageStart "info", 550 %>
		<b><% =sApplicationTitle %></b><br><br>
		Thank you for having registered your CV. You will receive an electronic confirmation to the email address you currently have registered with us which is  
		&nbsp;<b><% =sUserEmail %></b>
	<% ShowMessageEnd %>

<% CloseDBConnection %>
</body></html>

<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include file="../../../_common/invite.asp"-->
<%
Dim sUserLogin, sUserPassword, sUserPhone

Dim bEmailAlreadySent
If Request.QueryString("sent")>0 Then
	bEmailAlreadySent=1
Else
	bEmailAlreadySent=0
End If
%>
<html>
<head>
<title>Expert invitation</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<% 

LoadExpertProfile iExpertID
LoadExpertAccountDetails iExpertID
sFullName=Trim(sFirstName & " " & sLastName)

	If bEmailAlreadySent=0 Then
		PrepareEmailTemplate "emlExpertInvite.htm", ";;sExpertFullName=" & sFullName & ";;sUserLogin=" & sUserLogin & ";;sUserPassword=" & sUserPassword & ";;sSystemUrl=" & "http://cvip.assortis.com" & sHomePath & "apply/" 
		SendEmail sEmailCvipSystem, sUserEmail, sEmailSubject, sEmailBody, "info"
		SendEmail sEmailCvipSystem, sEmailClient, sEmailSubject, sEmailBody, "info"
		Response.Redirect sScriptFileName & ReplaceUrlParams(sParams, "sent=1")
	End If
	
	ShowMessageStart "info", 550 %>
		<b>Resume Management System</b><br><br>
		Thank you for inviting a new expert.<br><br>
		Invitation was sent to an expert by the following email: <b><% =sUserEmail %></b><br>
		And copy of the invitation was sent to <b><% =sEmailClient %></b> as well.
		<br><br><br><a href="<% =sApplicationHomePath %>"><img src="<% =sHomePath %>image/bte_continue.gif" border="0"></a>
	<% ShowMessageEnd
%>

<% CloseDBConnection %>
</body></html>

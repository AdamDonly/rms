<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Check CV & send an email
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="cv_data.asp"-->
<%
' Check user's access rights
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else
	Response.Redirect sHomePath & "apply/"
End If
CheckExpertID()

' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register/register.asp" & sParams)
%>
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<!--#include virtual="/_common/_class/status_cv.asp"-->
<!--#include virtual="/_common/_class/expert.status_cv.asp"-->
<%
Dim objExpertStatusCV
Set objExpertStatusCV = New CExpertStatusCV
objExpertStatusCV.Expert.ID=iExpertID
objExpertStatusCV.LoadData
%>
<!--#include virtual="/_common/register/update_status.asp"-->
<html>
<head>
<title>CV management</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 7 %>

<br>
<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
	<%
	If sApplicationName="external" Then
		If bEmailExpertAccountSent = 2 Then
			If Not (Len(sUserEmail)>0 And InStr(sUserEmail, "@")>0) Or IsEmpty(sUserEmail) Or IsNull(sUserEmail) Then
				ShowMessageStart "error", 550
				Response.Write "<p><b>There is no valid email registered for " & ReplaceIfEmpty(Trim(sFullName), "the expert") & ".<b></p>"
			Else
				ShowMessageStart "info", 550
				%>
				<p>Login details were not sent to <% =sFullName %> yet!</p>
				<p><a href="register_confirm.asp?id=<% =iExpertID %>">Send personal login and password to the expert now</a>.</p>
				<%
			End If
			ShowMessageEnd
		ElseIf bEmailExpertAccountSent = 1 And sAction<>"resent" Then
			If Not (Len(sUserEmail)>0 And InStr(sUserEmail, "@")>0) Or IsEmpty(sUserEmail) Or IsNull(sUserEmail) Then
				ShowMessageStart "error", 550
				Response.Write "<p><b>There is no valid email registered for " & ReplaceIfEmpty(Trim(sFullName), "the expert") & ".<b></p>"
			Else
				ShowMessageStart "info", 550
				%>
				<p>Login details were already sent to <% =sFullName %>.</p>
				<p>If requested, you can</p>
				<p><a href="register_confirm.asp?id=<% =iExpertID %>">Resend welcome message and login details</a> in case expert's email wasn't correct, or</p> 
				<p><a href="password_confirm.asp?id=<% =iExpertID %>">Resend just the login details</a> to the expert.</p>
				<%
			End If
			ShowMessageEnd
		End If
	End If
	%>

	<!--#include file="register6_data.asp"-->
	<br>

</td>
<td width="5%">&nbsp;&nbsp;</td>
<td width="20%" valign="top">
	<!-- Feature boxes -->
	<img src="<% =sHomePath %>image/x.gif" width=1 height=23><br />

	<% 
	'ShowFeatureBoxHeader("CV status") 
	%>
  	<table width="176" border="0" cellpadding="0" cellspacing="0">
	<form method="post" action="<% =sScriptFullName %>">
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttltop.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center">CV status</p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
	
	<p class="sml" style="padding: 2px 5px;">Currently the CV status is</p>
	<div align="center"><select name="cv_status" id="cv_status" style="width: 152px;">
	<option value="0">New CV</option>
	<% 
	Dim objStatusCVList
	Set objStatusCVList = New CStatusCVList
	objStatusCVList.LoadData
	objStatusCVList.ShowSelectItems(objExpertStatusCV.Status.ID)
	%>
	</select></div>
	<div align="center"><input type="image" src="<% =sHomePath %>image/bte_updatestatus.gif" vspace="4" border="0" alt="Update status"></a></div>
	<% ShowFeatureBoxDelimiter %>
	<p class="sml" style="padding: 2px 5px;"><% If Len(sComments)<1 Then %>Some comments about this CV?<% Else %><% =sComments %><% End If %></p>
	<div align="center"><a href="../register/comments.asp?id=<%=iExpertID%>"><img src="<% =sHomePath %>image/bte_editcomments.gif" vspace="4" border="0" alt="Edit comments"></a></div>
	<% ShowFeatureBoxFooterWithFormFooter %>
	<br>	
	
	<% ShowFeatureBoxHeader("CV options") %>
	<p class="sml" style="padding: 2px 5px;">Some information from the CV<br>is missing?</p>
	<div align="center"><a href="register.asp?id=<%=iExpertID%>"><img src="<% =sHomePath %>image/bte_updatethiscv152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxDelimiter %>
	<p class="sml" style="padding: 2px 5px;">If this CV is a duplicate and expert has another CV</p>
	<div align="center"><a href="../manage/cv_hide.asp?id=<%=iExpertID%>"><img src="<% =sHomePath %>image/bte_hideexpert.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxDelimiter %>
	<p class="sml" style="padding: 2px 5px;">If expert asked to be removed or it's a fake CV</p>
	<div align="center"><a href="../manage/cv_remove.asp?id=<%=iExpertID%>"><img src="<% =sHomePath %>image/bte_removeexpert.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxFooter %>
	<br />
	
	
	<% ShowFeatureBoxHeader("CV formats") %>
	<p class="sml" style="padding: 2px 5px;">To view this CV in different formats, to save or to print it</p>
	<div align="center"><a href="<% =sApplicationHomePath %>view/cv_view.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "idexpert"), "id=" & iExpertID) %>"><img src="<% =sHomePath %>image/bte_formatcv152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxFooter %>
	<br />
	
</td>
</tr>

<% CloseDBConnection %>
</body>
</html>

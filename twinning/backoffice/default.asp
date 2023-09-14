<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../_forms/frmInterface.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<%
' Check UserID 
CheckUserLogin sScriptFullNameAsParams

sParams=ReplaceUrlParams(sParams, "idproject")
sParams=ReplaceUrlParams(sParams, "idexpert")
%>
<html>
<head>
<title><% =sApplicationTitle %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<br><br>
<table width="100%" cellpadding=0 cellspacing=1 border=0 align="center">
<tr>
<td width="7%">&nbsp;</td>
<td width="50%" valign="top">
</td>
<td width="5%">&nbsp;</td>
<td width="30%" valign="top">
<h3>Search for experts</h3>
<p>
<a href="search/exp_search.asp<% =sParams %>">Search for experts</a>
</p><br />
<% If iUserAdmin=1 Then %>
<h3>Register new expert</h3>
<p>
<a href="register/register.asp<% =sParams %>">Complete CV registration</a>
</p><br />
<h3>Manage the database</h3>
<p>
<a href="manage/cv_list.asp<% =sParams %>">List of all experts visible in the database</a>
<br><br>
<a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=registeredweek") %>">List of experts registered this week</a>
<br><br>
<a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=registeredmonth") %>">List of experts registered this month</a>
<br><br>
<a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=updatedover12&ord=E") %>">List of experts with CVs not updated for the past 12 months</a>
<br><br><br>
<a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=deleted") %>">List of deleted experts</a>
<% End If %>
</td>
<td width="7%">&nbsp;</td>
</tr>
</table>
</body>
</html>

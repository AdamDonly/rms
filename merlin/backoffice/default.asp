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
<h3>Register new project</h3>
	<p class="project"><a href="project/register.asp<% =AddUrlParams(sParams, "project_type=tendering") %>">New project</a></p>

<br />
<h3>Projects. Tendering</h3>
	<%
	Dim objProjectList
	Set objProjectList = New CProjectList
	objProjectList.LoadDataByStatusKeywords "111,112,121,122", "", ""
	For i=0 To objProjectList.Count-1
		%>
		<p class="project"><a href="project/details.asp<% =AddUrlParams(sParams, "idproject=" & objProjectList.Item(i).ID) %>"><% =objProjectList.Item(i).Title %></a></p>
		<%
	Next
	objProjectList.Clear
	%>
	<!--
	<p class="project"><a href="project/register.asp<% =AddUrlParams(sParams, "project_type=tendering") %>">New project</a></p>
	-->

<br />
<h3>Projects. Running</h3>
	<%
	objProjectList.LoadDataByStatusKeywords "201,202,203", "", ""
	For i=0 To objProjectList.Count-1
		%>
		<p class="project"><a href="project/details.asp<% =AddUrlParams(sParams, "idproject=" & objProjectList.Item(i).ID) %>"><% =objProjectList.Item(i).Title %></a></p>
		<%
	Next
	objProjectList.Clear
	%>
	<!--
	<p class="project"><a href="project/register.asp<% =AddUrlParams(sParams, "project_type=running") %>">New project</a></p>
	-->

<br />
<h3>Projects. Closed</h3>
	<%
	objProjectList.LoadDataByStatusKeywords "301", "", ""
	For i=0 To objProjectList.Count-1
		%>
		<p class="project"><a href="project/details.asp<% =AddUrlParams(sParams, "idproject=" & objProjectList.Item(i).ID) %>"><% =objProjectList.Item(i).Title %></a></p>
		<%
	Next
	objProjectList.Clear
	%>
	<!--
	<p class="project"><a href="project/register.asp<% =AddUrlParams(sParams, "project_type=closed") %>">New project</a></p>
	-->
<br />

	<%
	objProjectList.LoadDataByStatusKeywords "117,118,119,127,128,129", "", ""
	If objProjectList.Count>0 Then %>
<h3>Projects. Inactive</h3>
	<% 
	End If
	For i=0 To objProjectList.Count-1
		%>
		<p class="project"><a href="project/details.asp<% =AddUrlParams(sParams, "idproject=" & objProjectList.Item(i).ID) %>"><% =objProjectList.Item(i).Title %></a></p>
		<%
	Next
	objProjectList.Clear
	
Set objProjectList = Nothing
%>
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
<% End If %>
<h3>Manage the database</h3>
<p>
<a href="manage/cv_list.asp<% =sParams %>">List of all experts visible in the database</a>
<br><br>
<% If iUserAdmin=1 Then %>
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

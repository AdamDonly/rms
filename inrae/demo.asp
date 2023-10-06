<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->
<!--#include file="_forms/frmInterface.asp"-->
<%
AbandonSession()
%>
<html>
<head>
<title>CVIP start page - Demo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<br><br>
<table width="100%" cellpadding=0 cellspacing=1 border=0 align="center">
<tr>
<td width="15%">&nbsp;</td>
<td width="30%" valign="top">
<h3>Internal view</h3>
<p>These pages are visible only for your staff.</p>
<br><br><br><br><br>
<p>
<a href="backoffice/<% =sParams %>">CVIP internal tools</a>
</p>
</td>
<td width="5%">&nbsp;</td>
<td width="30%" valign="top">
<h3>View for experts</h3>
<p>The link for CV registration could be shown on public website, announced to your experts through email campaigns, added in company emails signatures.</p>
<br>
<p>There are 2 types of the experts registration:</p>
<br><p>
<a href="apply/quick.asp<% =sParams %>">Quick CV registration</a>
<br><br>
<a href="apply/register.asp<% =sParams %>">Complete CV registration</a>
</p>
</td>
<td width="15%">&nbsp;</td>
</tr>
</table>
</body>
</html>

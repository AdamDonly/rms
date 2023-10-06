<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit
'--------------------------------------------------------------------
'
' Access denied
'
'--------------------------------------------------------------------
%>
<!--#include file="dbc.asp"-->
<!--#include file="_forms/frmInterface.asp"-->

<html>
<head>
<title>Access is denied!</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="stylesheet" TYPE="text/css" href="styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<br><br><br><br><br><br><br>
<%
	ShowMessage "Access is denied!", "error", 150
%>

<% CloseDBConnection %>
</body>
</html>

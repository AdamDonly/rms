<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<%
If sDefaultRegistration="QUICK" Then
	Response.Redirect "quick.asp"
Else
	Response.Redirect "register.asp"
End If
%>
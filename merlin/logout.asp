<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->

<%
Session.Abandon
iUserID=0
iMemberID=0
iExpertID=0
sSessionID=""
sCookiesSessionID=""
Response.Cookies("SessionID")=""

Response.Redirect(sApplicationHomePath & "default.asp")
%>